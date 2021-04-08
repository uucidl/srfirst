// # Todo App
//
// An experiment in designing an app starting first from screen-reader support, before thinking about the GUI.
//

#define _CRT_SECURE_NO_WARNINGS
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include "wyhash.h"

#include <algorithm>
#include <array>
#include <iterator>
#include <string>
#include <vector>

#include <cstdarg>
#include <cstdio>
#include <cstdlib>
#include <cmath>

#include <Objbase.h>
#pragma comment(lib, "Ole32.lib")

#include <unknwn.h>

#include <Windows.h>


/// General utilities
#pragma region Utils

#define STRINGIFY_INNER(s) # s
#define STRINGIFY(s) STRINGIFY_INNER(s)
#define VERIFY(expr) do { auto r = (expr); if (!bool(r)) { \
  auto LastError = GetLastError(); auto LastErrorAsHRESULT = HRESULT_FROM_WIN32(LastError); \
  ::log("%s:%d: VERIFY(%s) failed. (GetLastError() returns %#x)\n", __FILE__, __LINE__, STRINGIFY(expr), LastErrorAsHRESULT); \
  if (::IsDebuggerPresent()) { ::DebugBreak(); } \
  std::exit(1); \
} } while(0)

#define VERIFYHR(expr) do { auto hr = (expr); VERIFY(SUCCEEDED(hr)); } while(0)

#define COMPLETE_SWITCH_BEGIN                                 \
  __pragma(warning(push))                                    \
  __pragma(warning(error : 4061; error : 4060; error : 4062))

#define COMPLETE_SWITCH_END \
  __pragma(warning(pop))


static FILE* g_output_log = fopen("log.txt", "wb");
//stdout;

void log(char const* fmt, ...);


struct ComScope {
  ComScope() { VERIFYHR(::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED)); }
  ~ComScope() { ::CoUninitialize(); }
};

template <typename T>
struct ComOwner {
  T* ptr;

  ComOwner() : ptr(nullptr) {}
  ComOwner(T* p) : ptr(p) {};
  ~ComOwner() { 
    if (ptr) { ptr->Release(); }
    ptr = nullptr;
  }
  ComOwner(const ComOwner& Other) {
    *this = Other;
  }
  ComOwner(ComOwner&& Other) {
    *this = std::move(Other);
  }

  operator bool() { return bool(ptr); }
  operator T* () { return ptr;  }

  ComOwner& operator=(const ComOwner& Other) {
    if (ptr) ptr->Release();
    ptr = Other.ptr;
    if (ptr) ptr->AddRef();
  }

  ComOwner& operator=(ComOwner&& other) {
    this->ptr = other.ptr;
    other.ptr = nullptr;
    return *this;
  }

  T** Slot() {
    if (this->ptr) { this->ptr->Release(); }
    this->ptr = nullptr;
    return &this->ptr;
  }

  template <typename Q>
  HRESULT QueryInterface(Q** pRetVal) {
    if (!ptr) { *pRetVal = nullptr; return E_POINTER; }
    return ptr->QueryInterface(__uuidof(Q), (void**)pRetVal);
  }
};

void
logv(char const* fmt, va_list args) {
  std::vfprintf(g_output_log, fmt, args);
  std::fflush(g_output_log);
}

void
log(char const* fmt, ...) {
  std::va_list args;
  va_start(args, fmt);
  ::logv(fmt, args);
  va_end(args);
}

uint64_t inline
bits(uint64_t x, uint64_t start, uint64_t num) {
  VERIFY(start + num < 64);
  return (x >> start) & ((1ULL << num) - 1);
}

bool inline
bit(uint64_t x, uint64_t bit) {
  VERIFY(bit < 64);
  return x & (1ULL << bit);
}

uint64_t
hash(size_t num_bytes, void const* bytes) {
  return wyhash(bytes, num_bytes, 0, _wyp);
}

RECT
intersection(RECT const a, RECT const b) {
  return {
    .left = std::max(a.left, b.left), .top = std::max(a.top, b.top),
    .right = std::min(a.right, b.right), .bottom = std::max(a.bottom, b.bottom),
  };
}

bool
contains(RECT const r, POINT const p) {
  if (p.y < r.top) return false;
  if (p.y >= r.bottom) return false;
  if (p.x < r.left) return false;
  if (p.x >= r.right) return false;
  return true;
}

RECT
operator+ (RECT const a, POINT b) {
  return {
    .left = a.left + b.x, .top = a.top + b.y,
    .right = a.right + b.x, .bottom = a.bottom + b.y,
  };
}

#pragma endregion Utils

/// Implementation of the Ui framework:
#pragma region Ui

struct UiVirtualKeyId {
  int x;
};

struct DigitalButton {
  bool is_down = false;
  bool pressed = false;
  bool released = false;
};

#define IdFormat "%#x"

struct RootProvider;

struct Ui {
  /// identifies a node in the ui tree.
  /// 0 => root node
  /// -1 => invalid node.
  using Id = std::uint32_t;

  HWND hwnd;
  ComOwner<RootProvider> root_provider;

  struct {
    bool updated = false;
    DigitalButton keys_per_vk[256];
    DigitalButton shift_key;
  } inputs;

  struct {
    Ui::Id id = 0;
    bool updated = false;
  } focus;

  enum Type {
    kNone,   // invalid
    kText,   // static label / paragraph
    kButton, // button that can be activated
    kPane,   // a grouping of elements
  };

  int depth_for_adding_nodes = 0;

  // Nodes with their properties as separate arrays, APL-style.
  // Elements are ordered in depth-first traversal.
  std::vector<Id>           node_ids; // in presentation order.
  std::vector<std::wstring> node_names;
  std::vector<Type>         node_type;
  std::vector<Id>           node_parent;
  std::vector<int>          node_depth;
  std::vector<RECT>         node_rect;
  
  struct {
    std::vector<Id> ids;
    std::vector<DigitalButton> state;
  } buttons;
};

Ui g_ui;

bool
valid_id(Ui::Id id) {
  return 0 < id && id < Ui::Id(-1);
}

size_t
ui_get_index(Ui::Id id) {
  VERIFY(valid_id(id));
  static struct {
    Ui::Id id[2];
    size_t index[2];
  } fingers = {};
  static int next = 0;

  if (id == fingers.id[0]) {
    return fingers.index[0];
  }
  else if (id == fingers.id[1]) {
    return fingers.index[1];
  }

  log("ui_get_index finger cache miss, for id %#llx (cached ids: %#llx %#llx)\n", id, fingers.id[0], fingers.id[1]);

  size_t index;
  for (index = 0; index < g_ui.node_ids.size(); index++) {
    if (g_ui.node_ids[index] == id) break;
  }
  VERIFY(index < g_ui.node_ids.size()); // is this a case that needs instead to be legitimately handled, like if we have elements that disappear?
  fingers.id[next & 1] = id;
  fingers.index[next & 1] = index;
  next++;

  return index;
}

char const*
type_desc(Ui::Type type) {
COMPLETE_SWITCH_BEGIN
  switch (type) {
  case Ui::Type::kNone: return "(invalid)";
  case Ui::Type::kButton: return "Button";
  case Ui::Type::kPane: return "Pane";
  case Ui::Type::kText: return "Text";
  }
COMPLETE_SWITCH_END
  VERIFY(0);
  return "(unknown)";
}

// Focus

void
ui_update_focus(Ui& ui, Ui::Id new_id) {
  auto old_id = ui.focus.id;
  ui.focus.id = new_id;
  ui.focus.updated = new_id != old_id;
}


// Inputs

void
ui_update(DigitalButton* button, bool is_down) {
  auto was_down = button->is_down;
  button->is_down = is_down;
  button->pressed = is_down && !was_down;
  button->released = was_down && !is_down;
}

bool
ui_on_press(const DigitalButton button) {
  return button.pressed;
}

bool
ui_on_press(UiVirtualKeyId key) {
  return ui_on_press(g_ui.inputs.keys_per_vk[key.x]);
}

bool
ui_on_press(int key) {
  VERIFY(0 <= key && key < 256);
  return ui_on_press(UiVirtualKeyId(key));
}

bool
ui_down(UiVirtualKeyId key) {
  return g_ui.inputs.keys_per_vk[key.x].is_down;
}

bool
ui_down(int key) {
  VERIFY(0 <= key && key < 256);
  return ui_down(UiVirtualKeyId(key));
}

POINT
ui_point_from_screen_point(const Ui& ui, double x, double y) {
  POINT LeftTop = { .x = 0, .y = 0 };
  VERIFY(::ClientToScreen(ui.hwnd, &LeftTop));

  return {
    .x = LONG(std::round(x - LeftTop.x)),
    .y = LONG(std::round(y - LeftTop.y)),
  };
}

Ui::Id
ui_search_deepest_node_containing(POINT pt) {
  auto depth = 0;
  Ui::Id id = 0;
  for (size_t i = 0; i < g_ui.node_rect.size(); i++) {
    auto d = g_ui.node_depth[i];
    if (d < depth) continue;

    auto r = g_ui.node_rect[i];
    if (!contains(r, pt)) continue;

    depth = d;
    id = g_ui.node_ids[i];
  }
  return id;
}

Ui::Id
ui_prev_sibling(const Ui& ui, Ui::Id id) {
  VERIFY(valid_id(id));
  auto index = ui_get_index(id);
  if (index == 0) return -1;
  auto rf = std::make_reverse_iterator(ui.node_depth.begin() + index - 1);
  auto rl = ui.node_depth.rend();
  auto pos = std::find(rf, rl, ui.node_depth[index]);
  if (pos == rl) return -1;
  return ui.node_ids[std::distance(ui.node_depth.begin(), pos.base())];
}

Ui::Id
ui_next_sibling(const Ui& ui, Ui::Id id) {
  VERIFY(valid_id(id));
  auto index = ui_get_index(id);
  if (index + 1 == ui.node_depth.size()) return -1;

  auto f = ui.node_depth.begin();
  auto l = ui.node_depth.end();
  auto pos = std::find(f + index + 1, l, ui.node_depth[index]);
  if (pos == ui.node_depth.end()) return -1;
  return ui.node_ids[std::distance(f, pos)];
}

Ui::Id
ui_first_child(const Ui& ui, Ui::Id id) {
  VERIFY(valid_id(id));
  auto index = ui_get_index(id);
  if (index + 1 == ui.node_depth.size()) return -1;

  auto f = ui.node_depth.begin();
  auto l = ui.node_depth.end();
  auto pos = std::find(f + index + 1, l, ui.node_depth[index] + 1);
  if (pos == ui.node_depth.end()) return -1;
  return ui.node_ids[std::distance(f, pos)];
}

Ui::Id
ui_last_child(const Ui& ui, Ui::Id id) {
  VERIFY(valid_id(id));
  auto index = ui_get_index(id);
  if (index + 1 == ui.node_depth.size()) return -1;

  auto f = ui.node_depth.begin();
  auto l = ui.node_depth.end();
  l = std::find(f + index + 1, l, ui.node_depth[index]);
  auto rf = std::make_reverse_iterator(l);
  auto rl = std::make_reverse_iterator(f + index + 1);
  auto pos = std::find(rf, rl, ui.node_depth[index] + 1);
  if (pos == rl) return -1;
  return ui.node_ids[std::distance(f, pos.base())];
}

// Describing the UI tree

void
ui_begin() {
  auto& ui = g_ui;
  
  // for now we're recreating the structure each time, which doesn't allow detecting structural changes, which will be necessary later on.
  /* remove buttons that no longer exist */ {
    std::vector<Ui::Id> live_nodes = ui.node_ids;
    std::sort(live_nodes.begin(), live_nodes.end());

    std::vector<std::pair<Ui::Id, size_t>> button_nodes;
    for (size_t i = 0; i < ui.buttons.ids.size(); i++) {
      button_nodes.push_back({ ui.buttons.ids[i], i });
    }
    std::sort(button_nodes.begin(), button_nodes.end(), [](const auto& a, const auto& b) {
      return a.first < b.first;
    });

    std::vector<std::pair<Ui::Id, size_t>> buttons_to_remove;
    struct LessThan {
      bool operator()(const std::pair<Ui::Id, size_t> a, const Ui::Id& b) {
        return a.first < b;
      }
      bool operator()(const Ui::Id& a, const std::pair<Ui::Id, size_t> b) {
        return a < b.first;
      }
    };

    std::set_difference(button_nodes.begin(), button_nodes.end(), live_nodes.begin(), live_nodes.end(), std::back_inserter(buttons_to_remove), LessThan());
    for (auto [id, index] : buttons_to_remove) {
      VERIFY(ui.buttons.ids[index] == id);
      ui.buttons.ids.erase(ui.buttons.ids.begin() + index);
      ui.buttons.state.erase(ui.buttons.state.begin() + index);
    }
  }

  ui.node_ids.clear();
  ui.node_names.clear();
  ui.node_type.clear();
  ui.node_depth.clear();
  ui.node_parent.clear();
  ui.node_rect.clear();
}

void ui_uia_raise_events_for_updates(const Ui& ui);

void
ui_end() {

  // Global input handlers, such as for focus changes:
  auto& ui = g_ui;
  VERIFY(ui.focus.id == 0 || std::ranges::end(ui.node_ids) != std::ranges::find(ui.node_ids, ui.focus.id));

  auto& inputs = ui.inputs;
  bool focus_next = false;
  bool focus_prev = false;
  if (inputs.updated) {
    focus_next = ui_on_press(VK_DOWN)
      || (!inputs.shift_key.is_down && ui_on_press({ VK_TAB }));
    focus_prev = ui_on_press(VK_UP)
      || (inputs.shift_key.is_down && ui_on_press(VK_TAB));

    if (!ui.focus.id && !ui.node_ids.empty() && (focus_next || focus_prev)) {
      ui_update_focus(ui, ui.node_ids[0]);
    } else if (focus_next) {
      log("User wants to focus the next element (keyboard)\n");
      auto index = ui_get_index(ui.focus.id);
      if (index + 1 < ui.node_ids.size()) {
        index++;
        ui_update_focus(ui, ui.node_ids[index]);
      }
    } else if (focus_prev) {
      log("User wants to focus the previous element (keyboard)\n");
      auto index = ui_get_index(ui.focus.id);
      if (index > 0) {
        index--;
        ui_update_focus(ui, ui.node_ids[index]);
      }
    }
  }

  ui_uia_raise_events_for_updates(ui);

  // reset button triggers:
  for (auto& state : ui.buttons.state) {
    state.pressed = false;
    state.released = false;
  }
  ui.focus.updated = false;

  inputs.updated = false;
  VERIFY(g_ui.depth_for_adding_nodes == 0); // Unbalanced?
}

size_t
ui_search_parent_index_for_adding(const Ui& ui) {
  auto index = ui.node_ids.size();
  auto depth = ui.depth_for_adding_nodes;
  auto parent_depth = depth - 1;
  auto parent_pos = std::find(ui.node_depth.rbegin(), ui.node_depth.rend(), parent_depth);
  VERIFY(parent_pos != ui.node_depth.rend());
  auto relative_index = std::distance(ui.node_depth.rbegin(), parent_pos);
  auto parent_index = index - 1 - relative_index;
  return parent_index;
}

Ui::Id
ui_named_element(wchar_t const* name, Ui::Type type, wchar_t const* text) {
  auto& ui = g_ui;
  auto index = ui.node_ids.size();
  auto depth = ui.depth_for_adding_nodes;

  Ui::Id parent_id = -1;
  if (depth == 0) {
    parent_id = 0;
  }
  else {
    auto parent_index = ui_search_parent_index_for_adding(ui);
    parent_id = g_ui.node_ids[parent_index];
  }

  auto num_bytes = wcslen(name) * sizeof name[0];
  auto genid = hash(num_bytes, name);
  genid = wyhash64(genid, parent_id);

  auto id = Ui::Id(bits(genid, 0, 32));

  VERIFY(valid_id(id));
  VERIFY(ui.node_ids.end() == std::find(ui.node_ids.begin(), ui.node_ids.end(), id));
  ui.node_ids.push_back(id);
  ui.node_names.push_back(text ? text : name);
  ui.node_type.push_back(type);
  ui.node_depth.push_back(depth);
  ui.node_parent.push_back(parent_id);
  ui.node_rect.push_back({});

  // TODO(nil): debug/hack
  ui.node_rect[index].right = 200;
  ui.node_rect[index].bottom = 200;
  return id;
}

Ui::Id
ui_text_paragraph(wchar_t const* content) {
  return ui_named_element(content, Ui::Type::kText, nullptr);
}


struct ButtonResult {
  Ui::Id id;
  bool activated;
};

ButtonResult
ui_button(wchar_t const* name, wchar_t const* text = nullptr) {
  auto& ui = g_ui;
  auto id = ui_named_element(name, Ui::Type::kButton, text);
  
  auto button_pos = std::find(ui.buttons.ids.begin(), ui.buttons.ids.end(), id);
  size_t button_index;
  if (button_pos == ui.buttons.ids.end()) {
    button_index = ui.buttons.ids.size();
    ui.buttons.ids.push_back(id);
    ui.buttons.state.push_back({});
  }
  else {
    button_index = std::distance(ui.buttons.ids.begin(), button_pos);
  }
  auto& state = ui.buttons.state[button_index];

  bool is_down = g_ui.inputs.updated && g_ui.focus.id == id && g_ui.inputs.keys_per_vk[VK_RETURN].is_down;
  ui_update(&state, is_down);

  if (state.released) {
    return { id, true };
  }
  return { id, false };
}

Ui::Id
ui_pane_begin(wchar_t const* name) {
  auto id = ui_named_element(name, Ui::Type::kPane, nullptr);
  g_ui.depth_for_adding_nodes++;
  return id;
}

void
ui_pane_end(Ui::Id pane) {
  auto pane_index = ui_search_parent_index_for_adding(g_ui);
  VERIFY(pane == g_ui.node_ids[pane_index]);
  g_ui.depth_for_adding_nodes--;
  VERIFY(g_ui.depth_for_adding_nodes == g_ui.node_depth[pane_index]);
}

void
ui_log_structure() {
  log("UI Tree:\n");
  for (size_t i = 0; i < g_ui.node_ids.size(); i++) {
    unsigned long long id = g_ui.node_ids[i];
    int depth = g_ui.node_depth[i];
    auto type = g_ui.node_type[i];
    const auto& name = g_ui.node_names[i];
    log("%*s", 2 + 4 * int(depth), "");
    log("node: type(%s)", type_desc(type));
    log(" " IdFormat, id);
    if (g_ui.focus.id == id) {
      log("*");
    }
    log(" (%ls)\n", name.data());
  }
  log("\n");
}

#pragma endregion Ui

/// UI Automation part of the Ui framework:
#pragma region UI_UIA

void main_update();

#include <uiautomationclient.h>
#include <uiautomationcore.h>
#include <uiautomationcoreapi.h>
#pragma comment(lib, "Uiautomationcore.lib")

// Some horrible macros to make defining COM objects more tolerable. I'll find
// an alternative some other day.

#define LOG_DEFINE_METHOD(cls_)          \
  void log(char const* fmt, ...) {       \
    ::log("this(%p) " #cls_ "::", this);  \
    std::va_list args;                   \
    va_start(args, fmt);                 \
    ::logv(fmt, args);                   \
    va_end(args);                        \
  }

#define IUNKNOWN_DEFS \
  ULONG STDMETHODCALLTYPE AddRef() override { return ++reference_count; }  \
  ULONG STDMETHODCALLTYPE Release() override { return --reference_count; } \
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override; \
  \
  ULONG reference_count = 1

#define COM_REQUIRE_PTR(x) do { if (!x) { return E_POINTER; } else { (*x) = {}; } } while(0)

UiaRect
ui_screen_rect(const Ui& ui, Ui::Id id) {
  POINT LeftTop = { .x = 0, .y = 0 };
  VERIFY(::ClientToScreen(ui.hwnd, &LeftTop));
  auto r = ui.node_rect[ui_get_index(id)];
  return {
    .left = double(LeftTop.x + r.left),
    .top = double(LeftTop.y + r.top),
    .width = double(r.right - r.left),
    .height = double(r.bottom - r.top),
  };
}

IRawElementProviderFragment*
create_element_provider(Ui::Id id);

struct RootProvider : public IRawElementProviderSimple
  , public IRawElementProviderFragmentRoot
  , public IRawElementProviderFragment {
  // IRawElementProviderFragmentRoot
  HRESULT STDMETHODCALLTYPE ElementProviderFromPoint(double x, double y, IRawElementProviderFragment** pRetVal) override {
    log("%s\n", __func__);
    COM_REQUIRE_PTR(pRetVal);
    auto pt = ui_point_from_screen_point(g_ui, x, y);
    auto id = ui_search_deepest_node_containing(pt);
    if (id) {
      *pRetVal = create_element_provider(id);
    }
    else {
      this->AddRef();
      *pRetVal = static_cast<IRawElementProviderFragment*>(this);
    }
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE GetFocus(IRawElementProviderFragment** pRetVal) {
    COM_REQUIRE_PTR(pRetVal);
    if (g_ui.focus.id) {
      *pRetVal = create_element_provider(g_ui.focus.id);
    }
    return S_OK;
  }

  // IRawElementProviderFragment
  HRESULT STDMETHODCALLTYPE get_BoundingRectangle(UiaRect* pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    RECT ClientRect;
    VERIFY(::GetClientRect(g_ui.hwnd, &ClientRect));
    POINT LeftTop = { .x = ClientRect.left, .y = ClientRect.top };
    VERIFY(::ClientToScreen(g_ui.hwnd, &LeftTop));
    pRetVal->left = double(LeftTop.x);
    pRetVal->top = double(LeftTop.y);
    pRetVal->width = double(ClientRect.right) - ClientRect.left;
    pRetVal->height = double(ClientRect.bottom) - ClientRect.top;
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE get_FragmentRoot(IRawElementProviderFragmentRoot** pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    this->AddRef();
    *pRetVal = this;
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE GetEmbeddedFragmentRoots(SAFEARRAY** pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE GetRuntimeId(SAFEARRAY** pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE Navigate(NavigateDirection direction, IRawElementProviderFragment** pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    Ui::Id found_id = -1;
    const auto find_at_depth = [](auto First, auto Last, int depth) -> Ui::Id {
      auto first_pos = std::find(First, Last, 0);
      if (first_pos == Last) return -1;
      auto index = std::distance(First, first_pos);
      return g_ui.node_ids[std::distance(First, first_pos)];
    };
COMPLETE_SWITCH_BEGIN
      switch (direction) {
      case NavigateDirection_FirstChild: {
        auto f = g_ui.node_depth.begin();
        auto l = g_ui.node_depth.end();
        auto pos = std::find(f, l, 0);
        found_id = pos == l ? -1 : g_ui.node_ids[std::distance(g_ui.node_depth.begin(), pos)];
      } break;
      case NavigateDirection_LastChild: {
        auto f = g_ui.node_depth.rbegin();
        auto l = g_ui.node_depth.rend();
        auto pos = std::find(f, l, 0);
        found_id = pos == l ? -1 : g_ui.node_ids[std::distance(g_ui.node_depth.begin(), pos.base())];
      } break;

      case NavigateDirection_Parent: break; // nothing to return, per spec
      case NavigateDirection_NextSibling: break; // nothing to return, per spec
      case NavigateDirection_PreviousSibling: break; // nothing to return, per spec
      }
COMPLETE_SWITCH_END
    if (valid_id(found_id)) {
      *pRetVal = create_element_provider(found_id);
    }
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE SetFocus() override {
    return S_OK;
  }

  // IRawElementProviderSimple
  HRESULT STDMETHODCALLTYPE get_HostRawElementProvider(IRawElementProviderSimple** pRetVal) override {
    return UiaHostProviderFromHwnd(g_ui.hwnd, pRetVal);
  }
  HRESULT STDMETHODCALLTYPE get_ProviderOptions(ProviderOptions* pRetVal) override {
    return ProviderOptions_ServerSideProvider | ProviderOptions_UseComThreading;
  }
  HRESULT STDMETHODCALLTYPE GetPatternProvider(PATTERNID patternId, IUnknown** pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE GetPropertyValue(PROPERTYID propertyId, VARIANT* pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    return S_OK;
  }

  LOG_DEFINE_METHOD(RootProvider)
  IUNKNOWN_DEFS;
};

struct AnyElementProvider : public IRawElementProviderSimple, public IRawElementProviderFragment {
  AnyElementProvider(Ui::Id id) : id(id) {}

  // IRawElementProviderFragment
  HRESULT STDMETHODCALLTYPE get_BoundingRectangle(UiaRect* pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    auto ScreenRect = ui_screen_rect(g_ui, id);
    *pRetVal = ScreenRect;
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE get_FragmentRoot(IRawElementProviderFragmentRoot** pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    return g_ui.root_provider.QueryInterface<IRawElementProviderFragmentRoot>(pRetVal);
  }

  HRESULT STDMETHODCALLTYPE GetEmbeddedFragmentRoots(SAFEARRAY** pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE GetRuntimeId(SAFEARRAY** pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    std::array ids{ int(UiaAppendRuntimeId), int(id) };
    auto psa = ::SafeArrayCreateVector(VT_I4, 0, LONG(ids.size()));
    if (!psa) return E_OUTOFMEMORY;

    for (LONG i = 0; i < LONG(ids.size()); i++) {
      VERIFYHR(::SafeArrayPutElement(psa, &i, &(ids[i])));
    }
    *pRetVal = psa;
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE Navigate(NavigateDirection direction, IRawElementProviderFragment** pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    Ui::Id found_id = -1;

    const auto this_index = ui_get_index(id);
COMPLETE_SWITCH_BEGIN
      switch (direction) {
      case NavigateDirection_Parent: {
        found_id = g_ui.node_parent[this_index];
      } break;

      case NavigateDirection_NextSibling: {
        found_id = ui_next_sibling(g_ui, id);
      } break;

      case NavigateDirection_PreviousSibling: {
        found_id = ui_prev_sibling(g_ui, id);
      } break;

      case NavigateDirection_FirstChild: {
        found_id = ui_first_child(g_ui, id);
      } break;

      case NavigateDirection_LastChild: {
        found_id = ui_last_child(g_ui, id);
      } break;
      }
COMPLETE_SWITCH_END

    if (found_id == this->id) return S_OK;
    
    if (found_id == 0) {
      return g_ui.root_provider.QueryInterface<IRawElementProviderFragment>(pRetVal);
    }
    else if (valid_id(found_id)) {
      *pRetVal = create_element_provider(found_id);
    }

    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE SetFocus() override {
    ui_update_focus(g_ui, id);
    main_update(); // TODO(nil): or post a message.
    return S_OK;
  }

  // IRawElementProviderSimple
  HRESULT STDMETHODCALLTYPE get_HostRawElementProvider(IRawElementProviderSimple** pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE get_ProviderOptions(ProviderOptions* pRetVal) override {
    return ProviderOptions_ServerSideProvider | ProviderOptions_UseComThreading;
  }
  HRESULT STDMETHODCALLTYPE GetPatternProvider(PATTERNID patternId, IUnknown** pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    // TODO(nil): implement button
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE GetPropertyValue(PROPERTYID propertyId, VARIANT* pRetVal) override {
    COM_REQUIRE_PTR(pRetVal);
    auto this_index = ui_get_index(id);
    switch (propertyId) {
    case UIA_NamePropertyId: {
      pRetVal->vt = VT_BSTR;
      pRetVal->bstrVal = ::SysAllocString(g_ui.node_names[this_index].data());
    } break;
    case UIA_IsKeyboardFocusablePropertyId: {
      pRetVal->vt = VT_BOOL;
      pRetVal->boolVal = VARIANT_TRUE;
    } break;
    case UIA_HasKeyboardFocusPropertyId: {
      pRetVal->vt = VT_BOOL;
      pRetVal->boolVal = g_ui.focus.id == this->id ? VARIANT_TRUE : VARIANT_FALSE;
    } break;
    case UIA_ControlTypePropertyId: {
      pRetVal->vt = VT_I4;
      COMPLETE_SWITCH_BEGIN
      switch (g_ui.node_type[this_index]) {
      case Ui::Type::kNone: VERIFY(0);  break;
      case Ui::Type::kPane: pRetVal->lVal = UIA_PaneControlTypeId; break;
      case Ui::Type::kText: pRetVal->lVal = UIA_TextControlTypeId; break;
      case Ui::Type::kButton: pRetVal->lVal = UIA_ButtonControlTypeId; break;
      }
      COMPLETE_SWITCH_END
    } break;
    }

    return S_OK;
  }

  LOG_DEFINE_METHOD(AnyElementProvider)
  IUNKNOWN_DEFS;

  Ui::Id id;
};

HRESULT
RootProvider::QueryInterface(REFIID riid, void** ppvObject) {
  auto result = [&]() -> std::pair<char const*, void*> {
    // supported interfaces:
    if (riid == IID_IRawElementProviderSimple) return { "IRawElementProviderSimple", static_cast<IRawElementProviderSimple*>(this) };
    if (riid == IID_IRawElementProviderFragmentRoot) return { "IRawElementProviderFragmentRoot", static_cast<IRawElementProviderFragmentRoot*>(this) };
    if (riid == IID_IRawElementProviderFragment) return { "IRawElementProviderFragment", static_cast<IRawElementProviderFragment*>(this) };
    // not supported, but logged
    if (riid == IID_IAccIdentity) return { "IAccIdentity", nullptr };

    return { nullptr, nullptr };
  }();

  if (riid == IID_IUnknown) {
    // should always be supported
    result = { "IUnknown", this };
  }

  COM_REQUIRE_PTR(ppvObject);
  LPOLESTR riid_string;
  VERIFYHR(::StringFromIID(riid, &riid_string));
  if (false) {
    log("%s %u (%ls)\n", __func__, riid.Data1, riid_string);
  }

  *ppvObject = result.second; // it's important to also assign it a value even if E_NOTINTERFACE
  if (!result.second) {
    if (false) {
      if (result.first) { log("  missing %s interface (not supported)\n", result.first); }
    }
    return E_NOINTERFACE;
  }

  AddRef();
  if (false) {
    log("  supported_interface %s\n", result.first);
  }
  return S_OK;
}

HRESULT
AnyElementProvider::QueryInterface(REFIID riid, void** ppvObject) {
  auto result = [&]() -> std::pair<char const*, void*> {
    // supported interfaces:
    if (riid == IID_IRawElementProviderSimple) return { "IRawElementProviderSimple", static_cast<IRawElementProviderSimple*>(this) };
    if (riid == IID_IRawElementProviderFragment) return { "IRawElementProviderFragment", static_cast<IRawElementProviderFragment*>(this) };
    // not supported, but logged
    if (riid == IID_IAccIdentity) return { "IAccIdentity", nullptr };

    return { nullptr, nullptr };
  }();

  if (riid == IID_IUnknown) {
    // should always be supported
    result = { "IUnknown", this };
  }

  COM_REQUIRE_PTR(ppvObject);
  LPOLESTR riid_string;
  VERIFYHR(::StringFromIID(riid, &riid_string));
  if (false) {
    log("%s %u (%ls)\n", __func__, riid.Data1, riid_string);
  }

  *ppvObject = result.second; // it's important to also assign it a value even if E_NOTINTERFACE
  if (!result.second) {
    if (false) {
      if (result.first) { log("  missing %s interface (not supported)\n", result.first); }
    }
    return E_NOINTERFACE;
  }

  AddRef();
  if (false) {
    log("  supported_interface %s\n", result.first);
  }
  return S_OK;
}

IRawElementProviderFragment*
create_element_provider(Ui::Id id) {
  return new AnyElementProvider(id);
}


void
ui_uia_raise_events_for_updates(const Ui& ui) {
  if (!::UiaClientsAreListening()) return;

  if (ui.focus.updated) {
    ComOwner p = create_element_provider(g_ui.focus.id);
    ComOwner<IRawElementProviderSimple> sp;
    VERIFYHR(p.QueryInterface(sp.Slot()));
    VERIFYHR(UiaRaiseAutomationEvent(sp, UIA_AutomationFocusChangedEventId));
  }

  for (size_t i = 0; i < ui.buttons.state.size(); i++) {
    const auto& state = ui.buttons.state[i];
    const auto id = ui.buttons.ids[i];
    if (state.released) {
      // button was activated.
      ComOwner p = create_element_provider(id);
      ComOwner<IRawElementProviderSimple> sp;
      VERIFYHR(p.QueryInterface(sp.Slot()));
      UiaRaiseAutomationEvent(sp, UIA_Invoke_InvokedEventId);
    }
  }
}

#pragma endregion UI_UIA

/// Actual application
#pragma region TodoApp

#include "TodoAppResources.h"

INT_PTR about_dlgproc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
  switch (uMsg) {
  case WM_INITDIALOG: return TRUE;
  case WM_COMMAND: // fallthrough
  case WM_CLOSE: {
    VERIFY(::EndDialog(hwnd, 0));
    return TRUE;
  } break;
  }
  return FALSE;
}

void
main_update() {
  // Ui State:
  static auto show_content = false;

  auto content_need_refresh = true;
  while (content_need_refresh) {
    content_need_refresh = false;
    ui_begin();
    if (auto pane = ui_pane_begin(L"Main")) {
      auto show_content_button = ui_button(L"Content Toggle", show_content ? L"Hide Content" : L"Show Content");
      if (show_content_button.activated) { show_content = !show_content; }
      if (show_content) {
        auto id = ui_text_paragraph(L"Lorem ipsum...");
        if (show_content_button.activated) {
          ui_update_focus(g_ui, id); // focus on the content being shown for the first time...
        }
        if (ui_button(L"Done").activated) {
          show_content = false;
          ui_update_focus(g_ui, show_content_button.id);
          content_need_refresh = true; // necessary because the name of the toggle button needs to change.

        }
      }
      ui_text_paragraph(L"You may close this app with the next button.");
      if (ui_button(L"Close application.").activated) {
        log("User requested to close the application by pressing the button.\n");
        VERIFY(::DestroyWindow(g_ui.hwnd));
      }
      ui_pane_end(pane);
    }
    ui_end();
  }
  ui_log_structure();
}

LRESULT CALLBACK
main_window_proc(_In_ HWND hwnd, _In_ UINT uMsg, _In_ WPARAM wParam, _In_ LPARAM lParam) {
  switch (uMsg) {
  case WM_DESTROY: { ::PostQuitMessage(0); } break;
  case WM_COMMAND: {
    switch (LOWORD(wParam)) {
    case ID_HELP_ABOUT: {
      ::DialogBoxW(nullptr, MAKEINTRESOURCE(IDD_ABOUT_DIALOG), hwnd, (DLGPROC)about_dlgproc);
      return 0;
    } break;
    case ID_FILE_EXIT: { VERIFY(::DestroyWindow(hwnd)); } break;
    }
  } break;

  case WM_GETOBJECT: {
    if (UiaRootObjectId == (DWORD)lParam) {
      if (!g_ui.root_provider) {
        g_ui.root_provider = new RootProvider;
        main_update();
      }
      return ::UiaReturnRawElementProvider(hwnd, wParam, lParam, g_ui.root_provider.ptr);
    }
  } break;
  case WM_KEYDOWN: // fallthrough
  case WM_KEYUP: {
    BYTE keys[256];
    VERIFY(::GetKeyboardState(keys));
    auto& dest = g_ui.inputs;
    for (size_t i = 0; i < std::size(keys); i++) {
      ui_update(&dest.keys_per_vk[i], bit(keys[i], 7));
    }
    ui_update(&dest.shift_key, bit(keys[VK_SHIFT], 7));

    auto this_key = LOWORD(wParam);
    VERIFY(ui_down(this_key) == (uMsg == WM_KEYDOWN));
    VERIFY((uMsg == WM_KEYDOWN) || ((uMsg == WM_KEYUP) && !ui_down(this_key)));

    g_ui.inputs.updated = true;
    main_update();
  } break;
  }

  return DefWindowProcW(hwnd, uMsg, wParam, lParam);
}

int __stdcall
wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd) {
  ComScope com;

  WNDCLASSW cls = { .lpfnWndProc = main_window_proc, .lpszMenuName = MAKEINTRESOURCEW(IDR_MENU1), .lpszClassName = L"TodoAppMainClass", };
  VERIFY(::RegisterClassW(&cls));

  auto Window = ::CreateWindowW(cls.lpszClassName, L"TodoApp", WS_CLIPCHILDREN | WS_GROUP | WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, nullptr, nullptr, nullptr, 0);
  VERIFY(Window);
  ::ShowWindow(Window, SW_SHOWNORMAL);

  g_ui.hwnd = Window;

  for (;;) {
    MSG msg;
    switch (::GetMessageW(&msg, nullptr, 0, 0)) {
    case -1: VERIFY(0); break;
    case 0: goto end; // WM_QUIT was received.
    default: break;
    }
    ::TranslateMessage(&msg);
    ::DispatchMessageW(&msg);
  }

end:
  return 0;
}

#pragma endregion TodoApp