// # SRFirst "Screen Reader First"
//
// An experiment in designing an app starting first from screen-reader support, before thinking about the GUI.
//

// First test it with the Accessibility Insight for Windows app, then test it with a screen-reader such as NVDA, Jaws or Narrator.
//
// I initially thought that elements shall have no graphical representation at all on the screen. It must be noted however that screen-readers do use the Mouse to select on hover certain elements. A visually-impaired user might use this to "feel" and "scan" the user interface. So probably elements should nevertheless have a position and take some amount of space, and allocate individual space for individual elements to give them a unique presence.

// QueryInterface and unknown GUIDs.
// ---------------------------------
// About the GUIds received by QueryInterface:
// You can find the corresponding interface name by searching in the registry (regedit.exe) at the key Computer\HKEY_CLASSES_ROOT\Interface
//
// There is one directory per GUID, and the interface is in there.
//   for instance {62F62F5A-5EC0-48F5-A032-783445DF9A89} on my machine is registered to IRawElementProviderComponent (for which there is no documentation)

// Fragment root
// -------------
// This is the UIA provider that sits on top of our hierarchy.

// Narrator (Microsoft)
// --------------------
// When I run it in scan mode, it manipulates the focus as it goes from element to element, once it starts reading (Caps_Lock + R)
//
// When I am not in scan mode, when I press (Caps_Lock + R) to read, it overrides the default focused element by calling SetFocus to start from its own idea of where to start. Why?
//
// NVDA too seems to not always follow my app's notion of focus, and I don't know why. I often have to forcefully Press Alt once to move to the system menu, then Alt again to move back from it to make the main pane the active one.

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#define _CRT_SECURE_NO_WARNINGS

#include "wyhash.h"

#include <Windows.h>

#include <objbase.h>
#pragma comment(lib, "Ole32.lib")

#include <uiautomationclient.h>
#include <uiautomationcore.h>
#include <uiautomationcoreapi.h>
#pragma comment(lib, "Uiautomationcore.lib")

#include <algorithm>
#include <cstdarg>
#include <cstdio>
#include <cstdlib>
#include <functional>
#include <string>
#include <unordered_map>
#include <vector>

// 1. Utils

#define STRINGIFY_INNER(s) # s
#define STRINGIFY(s) STRINGIFY_INNER(s)
#define VERIFY(expr) do { auto r = (expr); if (!bool(r)) { \
  auto LastError = GetLastError(); auto LastErrorAsHRESULT = HRESULT_FROM_WIN32(LastError); \
  log("%s:%d: VERIFY(%s) failed. (GetLastError() returns %#x)\n", __FILE__, __LINE__, STRINGIFY(expr), LastErrorAsHRESULT); \
  if (::IsDebuggerPresent()) { ::DebugBreak(); } \
  std::exit(1); \
} } while(0)

#define VERIFYHR(expr) do { auto hr = (expr); VERIFY(SUCCEEDED(hr)); } while(0)


static FILE* g_output_log = fopen("log.txt", "wb");
//stdout;

void
logv(char const* fmt, va_list args) {
    std::vfprintf(g_output_log, fmt, args);
    std::fflush(g_output_log);
}

void
log(char const* fmt, ...) {
  std::va_list args;
  va_start(args, fmt);
  logv(fmt, args);
  va_end(args);
}

uint64_t
bits(uint64_t x, uint64_t start, uint64_t num) {
  VERIFY(start + num < 64);
  return (x >> start)& ((1ULL << num) - 1);
}

uint64_t
hash(size_t num_bytes, void const* bytes) {
  return wyhash(bytes, num_bytes, 0, _wyp);
}

// 2. Actual program

// Sits at the top of the window and delivers the accessible ui to its client.
struct RootProvider : public IRawElementProviderSimple
  , public IRawElementProviderFragmentRoot
  , public IRawElementProviderFragment {

    void log(char const* fmt, ...) {
        ::log("this(%p) RootProvider::", this);
        std::va_list args;
        va_start(args, fmt);
        ::logv(fmt, args);
        va_end(args);
    }

  // IRawElementProviderFragmentRoot
  HRESULT STDMETHODCALLTYPE ElementProviderFromPoint(double x, double y, IRawElementProviderFragment** pRetVal) override;
  HRESULT STDMETHODCALLTYPE GetFocus(IRawElementProviderFragment** pRetVal) override;

  // IRawElementProviderFragment
  HRESULT STDMETHODCALLTYPE get_BoundingRectangle(UiaRect* pRetVal) override;
  HRESULT STDMETHODCALLTYPE get_FragmentRoot(IRawElementProviderFragmentRoot** pRetVal) override;
  HRESULT STDMETHODCALLTYPE GetEmbeddedFragmentRoots(SAFEARRAY** pRetVal) override;
  HRESULT STDMETHODCALLTYPE GetRuntimeId(SAFEARRAY** pRetVal) override;
  HRESULT STDMETHODCALLTYPE Navigate(NavigateDirection direction, IRawElementProviderFragment** pRetVal) override;
  HRESULT STDMETHODCALLTYPE SetFocus() override;

  // IRawElementProviderSimple
  HRESULT STDMETHODCALLTYPE get_HostRawElementProvider(IRawElementProviderSimple** pRetVal) override;
  HRESULT STDMETHODCALLTYPE get_ProviderOptions(ProviderOptions* pRetVal) override;
  HRESULT STDMETHODCALLTYPE GetPatternProvider(PATTERNID patternId, IUnknown** pRetVal) override;
  HRESULT STDMETHODCALLTYPE GetPropertyValue(PROPERTYID propertyId, VARIANT* pRetVal) override;

  // IUnknown interface:
  ULONG STDMETHODCALLTYPE AddRef() override { return ++reference_count; }
  ULONG STDMETHODCALLTYPE Release() override { return --reference_count; }
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;

  ULONG reference_count = 1;
};

static HWND g_hwnd;
static RootProvider* g_root_provider;

enum MenuId : UINT {
  MenuId_None,
  MenuId_File_Quit,
};

LRESULT CALLBACK main_window_proc(
  _In_ HWND   hwnd,
  _In_ UINT   uMsg,
  _In_ WPARAM wParam,
  _In_ LPARAM lParam
);

void ui_describe();
void ui_focus_next();
void ui_focus_prev();
void ui_activate();

int __stdcall
WinMain(
  HINSTANCE hInstance,
  HINSTANCE hPrevInstance,
  LPSTR     lpCmdLine,
  int       nShowCmd
) {
  if (false) {
    VERIFY(::SetConsoleCP(CP_UTF8));
    VERIFY(::SetConsoleOutputCP(CP_UTF8)); // Only works in Windows 10.
    std::printf("Author: Nicolas Léveillé. 2021-03.\n");
  }
  log("START: Starting SRFirst\n");
  VERIFYHR(::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED));
  
  WNDCLASSW Class = {
    .lpfnWndProc = main_window_proc,
    .lpszClassName = L"SRFirstMainClass",
  };
  VERIFY(::RegisterClassW(&Class));

  auto main_menu = ::CreateMenu(); // @tag{leak} intentional, the main menu lasts for the entire lifetime of the app.
  VERIFY(main_menu);
  {
    // URL(https://docs.microsoft.com/en-us/windows/win32/menurc/about-menus#standard-keyboard-interface)
    std::vector<HMENU> menu_stack;
    std::vector<UINT> pos_stack;
    menu_stack.push_back(main_menu);
    pos_stack.push_back(0);
    const auto BeginTopLevelMenu = [&](wchar_t const* title) -> HMENU {
      const auto submenu = ::CreateMenu();
      VERIFY(submenu);
      MENUITEMINFOW info = {
        .cbSize = sizeof(MENUITEMINFOW),
        .fMask = MIIM_STRING | MIIM_SUBMENU,
        .hSubMenu = submenu,
        .dwTypeData = const_cast<wchar_t*>(title),
      };
      VERIFY(::InsertMenuItemW(menu_stack.back(), pos_stack.back()++, TRUE, &info));
      menu_stack.push_back(submenu);
      pos_stack.push_back(0);
      return submenu;
    };
    const auto EndTopLevelMenu = [&]() {
      menu_stack.pop_back();
    };
    const auto PushEntry = [&](MenuId id, wchar_t const* title) {
      MENUITEMINFOW info = {
        .cbSize = sizeof(MENUITEMINFOW),
        .fMask = MIIM_ID | MIIM_STRING,
        .wID = id,
        .dwTypeData = const_cast<wchar_t*>(title),
      };
      VERIFY(::InsertMenuItemW(menu_stack.back(), pos_stack.back()++, TRUE, &info));
    };
    
    BeginTopLevelMenu(L"&File"); // & marks the mnemonic key used for keyboard access.
    PushEntry(MenuId_File_Quit, L"&Quit");
    EndTopLevelMenu();
    
    VERIFY(menu_stack.size() == 1);
  }

  auto Window = ::CreateWindowW(Class.lpszClassName, L"SRFirst", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, nullptr, main_menu, nullptr, 0);
  VERIFY(Window);
  g_hwnd = Window;
  VERIFY(::ShowWindow(Window, SW_SHOWNORMAL) == 0);
  ::SetFocus(Window);
  ui_describe();

  for (;;) {
    MSG msg;
    switch (::GetMessageW(&msg, nullptr, 0, 0)) {
    case -1: VERIFY(0); break;
    case 0: goto end; // WM_QUIT was received.
    default: break;
    }
    TranslateMessage(&msg);
    ::DispatchMessageW(&msg);
  }
end:
  VERIFYHR(::UiaDisconnectAllProviders());
  if (g_root_provider) {
      VERIFY(g_root_provider->Release() == 0);
      g_root_provider = nullptr;
  }
  ::CoUninitialize();
  log("END: Ended.\n");
  return 0;
}

LRESULT CALLBACK
main_window_proc(
  _In_ HWND   hwnd,
  _In_ UINT   uMsg,
  _In_ WPARAM wParam,
  _In_ LPARAM lParam
) {
    switch (uMsg) {
    case WM_CLOSE: {
        log("WM_CLOSE received\n");
        ::DestroyWindow(g_hwnd);
        return 0;
    } break;
    case WM_DESTROY: {
        log("WM_DESTROY received\n");
        // Microsoft recommends making this call from the WM_DESTROY message handler of the window that returns the UI Automation providers.
        // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcoreapi/nf-uiautomationcoreapi-uiareturnrawelementprovider)
        ::UiaReturnRawElementProvider(hwnd, 0, 0, NULL);
        ::PostQuitMessage(0);
        return 0;
    } break;
    //  2.1- Menu handling. URL(https://docs.microsoft.com/en-us/windows/win32/menurc/about-menus#messages-used-with-menus)
    case WM_COMMAND: {
      log("WM_COMMAND received with command: %#lx\n", long(wParam));
      switch ((MenuId)LOWORD(wParam)) {
      case MenuId_File_Quit: ::DestroyWindow(g_hwnd); return 0;
      }
      
    } break;
    case WM_GETOBJECT: {
      switch ((DWORD)lParam) {
      case UiaRootObjectId: {
        log("WM_GETOBJECT received for UiaAutomation with params: %u %u\n", wParam, lParam);
        if (!g_root_provider) {
          g_root_provider = new RootProvider;
        }
        return ::UiaReturnRawElementProvider(hwnd, wParam, lParam, g_root_provider);
      } break;

      default: break;
      }
    } 
    case WM_KEYDOWN: {
      if (0 == ((lParam >> 30) & 1 /* first transition bit*/)) {
        switch (wParam) {
        case VK_TAB: {
          log("User pressed <Tab> to change focus.\n");
          BYTE Keys[256];
          VERIFY(::GetKeyboardState(Keys));
          if (Keys[VK_SHIFT] & (1 << 7)) {
            log("  <Shift-Tab>\n");
            ui_focus_prev();
          }
          else {
            ui_focus_next();
          }
          return 0;
        } break;
        case VK_RETURN: {
          log("User pressed <Return> to activate primary action.\n");
          ui_activate();
          return 0;
        } break;
        default: {
          log("WM_KEYDOWN received: %#lx (unmapped)\n", long(wParam));
        } break;
        }
      }
    } break;
    case WM_CHAR: {
      log("WM_CHAR with character code %llx (unmapped)\n", long(wParam));
      return 0;
    } break;
    case WM_KILLFOCUS: { log("WM_KILLFOCUS received towards %ul\n", ULONG(wParam));  } break;
    case WM_SETFOCUS: {  log("WM_SETFOCUS received\n");   } break;
  }
  return ::DefWindowProcW(hwnd, uMsg, wParam, lParam);
}

struct UiTree {
  using Id = std::uint64_t;
  // Id == -1 => invalid_id
  // Id == 0  => root
  // ...

  enum class Type {
      kNone,
      kText,
      kDocument,
      kButton,
  };

  // Nodes with their properties as separate arrays, APL-style.
  std::vector<Id>           node_ids; // in presentation order.
  std::vector<std::wstring> node_names;
  std::vector<Type>         node_type;
  std::vector<Id>           node_parent;
  std::vector<int>          node_depth;

  std::unordered_map<Id, std::function<void()>> actions;

  Id focused_id = 0;

  int depth_for_adding_element = 0;
};

bool
valid_id(UiTree::Id id) {
  return 0 < id && id < UiTree::Id(-1);
}

void ui_set_focus_to(UiTree::Id id);
void ui_activate(UiTree::Id);

static UiTree g_ui;

bool
exists_id(UiTree::Id id) {
  return valid_id(id) && g_ui.node_ids.end() != std::find(g_ui.node_ids.begin(), g_ui.node_ids.end(), id);
}


IRawElementProviderFragment* create_element_provider(UiTree::Id element_id);
ITextProvider* create_element_text_provider(UiTree::Id element_id);
IValueProvider* create_element_value_provider(UiTree::Id element_id);
IInvokeProvider* create_element_invoke_provider(UiTree::Id element_id);

size_t ui_get_index(UiTree::Id element_id);

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

  LPOLESTR riid_string;
  VERIFYHR(::StringFromIID(riid, &riid_string));
  log("%s %u (%ls)\n", __func__, riid, riid_string);
  if (!ppvObject) return E_POINTER;

  *ppvObject = result.second; // it's important to also assign it a value even if E_NOTINTERFACE
  if (!result.second) {
    if (result.first) { log("  missing %s interface (not supported)\n", result.first); }
    return E_NOINTERFACE;
  }

  AddRef();
  log("  supported_interface %s\n", result.first);
  return S_OK;
}

HRESULT
RootProvider::get_HostRawElementProvider(IRawElementProviderSimple** pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementprovidersimple-get_hostrawelementprovider)
  return UiaHostProviderFromHwnd(g_hwnd, pRetVal);
}

HRESULT
RootProvider::get_ProviderOptions(ProviderOptions* pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementprovidersimple-get_provideroptions)
  return ProviderOptions_ServerSideProvider | ProviderOptions_UseComThreading;
}

HRESULT
RootProvider::GetPatternProvider(PATTERNID patternId, IUnknown** pRetVal) {
  log("%s %d\n", __func__, patternId);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementprovidersimple-getpatternprovider)
  if (!pRetVal) return E_POINTER;
  *pRetVal = nullptr;
  return S_OK;
}

HRESULT
RootProvider::GetPropertyValue(PROPERTYID propertyId, VARIANT* pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementprovidersimple-getpropertyvalue)
  if (!pRetVal) return E_POINTER;

  pRetVal->vt = VT_EMPTY;

  switch (propertyId) {
  case UIA_NamePropertyId: {
    pRetVal->vt = VT_BSTR;
    pRetVal->bstrVal = ::SysAllocString(L"SRFirstRoot");
  } break;

  case UIA_ProviderDescriptionPropertyId: {
    pRetVal->vt = VT_BSTR;
    pRetVal->bstrVal = ::SysAllocString(L"UU::AnyElementProvider");
  } break;
  case UIA_ClassNamePropertyId: {
    pRetVal->vt = VT_BSTR;
    pRetVal->bstrVal = ::SysAllocString(L"UU::AnyElementProvider");
  } break;
  }
  return S_OK;
}

HRESULT
RootProvider::get_BoundingRectangle(UiaRect* pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-get_boundingrectangle)
  if (!pRetVal) return E_INVALIDARG;
  RECT ClientRect;
  VERIFY(::GetClientRect(g_hwnd, &ClientRect));
  POINT LeftTop = { .x = ClientRect.left, .y = ClientRect.top };
  VERIFY(::ClientToScreen(g_hwnd, &LeftTop));
  pRetVal->left = double(LeftTop.x);
  pRetVal->top = double(LeftTop.y);
  pRetVal->width = double(ClientRect.right) - ClientRect.left;
  pRetVal->height = double(ClientRect.bottom) - ClientRect.top;
  return S_OK;
}

HRESULT
RootProvider::get_FragmentRoot(IRawElementProviderFragmentRoot** pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-get_fragmentroot)
  if (!pRetVal) return E_INVALIDARG;
  this->AddRef();
  *pRetVal = this;
  return S_OK;
}

HRESULT
RootProvider::GetEmbeddedFragmentRoots(SAFEARRAY** pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-getembeddedfragmentroots)
  if (!pRetVal) return E_INVALIDARG;
  *pRetVal = nullptr;
  return S_OK;
}

HRESULT
RootProvider::GetRuntimeId(SAFEARRAY** pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-getruntimeid)
  if (!pRetVal) return E_INVALIDARG;
  *pRetVal = nullptr;
  return S_OK;
}

HRESULT
RootProvider::Navigate(NavigateDirection direction, IRawElementProviderFragment** pRetVal) {
  log("%s %d\n", __func__, direction);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-navigate)
  if (!pRetVal) return E_INVALIDARG;
  *pRetVal = nullptr;

  UiTree::Id element_id = -1;

  switch (direction) {
  case NavigateDirection_FirstChild: {
    log("  first-child(Root)\n");
    if (!g_ui.node_ids.empty()) {
      for (size_t index = 0; index < g_ui.node_ids.size(); index++) {
        if (g_ui.node_depth[index] == 0) {
          VERIFY(g_ui.node_parent[index] == 0);
          element_id = g_ui.node_ids[index];
          break;
        }
      }
    }
  } break;
  case NavigateDirection_LastChild: {
    log("  last-child(Root)\n");
    if (!g_ui.node_ids.empty()) {
      for (size_t ri = g_ui.node_ids.size(); ri > 0; ri--) {
        auto index = ri - 1;
        if (g_ui.node_depth[index] == 0) {
          VERIFY(g_ui.node_parent[index] == 0);
          element_id = g_ui.node_ids[index];
          break;
        }
      }
    }
  } break;

  default: break;
  }

  if (valid_id(element_id)) {
    *pRetVal = create_element_provider(element_id); // TODO(nil): cache it somehow. 
  }

  return S_OK;
}

HRESULT
RootProvider::SetFocus() {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-setfocus)
  ::SetFocus(g_hwnd);
  return S_OK;
}

HRESULT
RootProvider::GetFocus(IRawElementProviderFragment** pRetVal) {
  log("%s (%#llx)\n", __func__, g_ui.focused_id);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragmentroot-getfocus)
  if (!pRetVal) return E_POINTER;

  *pRetVal = nullptr;
  if (g_ui.focused_id) {
    *pRetVal = create_element_provider(g_ui.focused_id);
  }

  return S_OK;
}

HRESULT
RootProvider::ElementProviderFromPoint(double x, double y, IRawElementProviderFragment** pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragmentroot-elementproviderfrompoint)
  if (!pRetVal) return E_POINTER;
  *pRetVal = nullptr;

  if (g_ui.focused_id) {
      *pRetVal = create_element_provider(g_ui.focused_id);
  }
  else {
      this->AddRef();
      *pRetVal = (IRawElementProviderFragment*)this;
  }

  return S_OK;
}

struct AnyElementProvider : public IRawElementProviderSimple, public IRawElementProviderFragment {
  // IRawElementProviderFragment
  HRESULT STDMETHODCALLTYPE get_BoundingRectangle(UiaRect* pRetVal) override;
  HRESULT STDMETHODCALLTYPE get_FragmentRoot(IRawElementProviderFragmentRoot** pRetVal) override;
  HRESULT STDMETHODCALLTYPE GetEmbeddedFragmentRoots(SAFEARRAY** pRetVal) override;
  HRESULT STDMETHODCALLTYPE GetRuntimeId(SAFEARRAY** pRetVal) override;
  HRESULT STDMETHODCALLTYPE Navigate(NavigateDirection direction, IRawElementProviderFragment** pRetVal) override;
  HRESULT STDMETHODCALLTYPE SetFocus() override;

  // IRawElementProviderSimple
  HRESULT STDMETHODCALLTYPE get_HostRawElementProvider(IRawElementProviderSimple** pRetVal) override;
  HRESULT STDMETHODCALLTYPE get_ProviderOptions(ProviderOptions* pRetVal) override;
  HRESULT STDMETHODCALLTYPE GetPatternProvider(PATTERNID patternId, IUnknown** pRetVal) override;
  HRESULT STDMETHODCALLTYPE GetPropertyValue(PROPERTYID propertyId, VARIANT* pRetVal) override;

  // IUnknown interface:
  ULONG STDMETHODCALLTYPE AddRef() override { return ++reference_count; }
  ULONG STDMETHODCALLTYPE Release() override { return --reference_count; }
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;

  void log(char const* fmt, ...) {
    ::log("this(%p, id=%#llx) AnyElementProvider::", this, this->id);
    std::va_list args;
    va_start(args, fmt);
    ::logv(fmt, args);
    va_end(args);
  }

  ULONG reference_count = 1;
  UiTree::Id id = -1;
};


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

  LPOLESTR riid_string;
  VERIFYHR(::StringFromIID(riid, &riid_string));
  log("%s %u (%ls)\n", __func__, riid, riid_string);
  if (!ppvObject) return E_POINTER;

  *ppvObject = result.second; // it's important to also assign it a value even if E_NOTINTERFACE
  if (!result.second) {
    if (result.first) { log("  missing %s interface (not supported)\n", result.first); }
    return E_NOINTERFACE;
  }

  AddRef();
  log("  supported_interface %s\n", result.first);
  return S_OK;
}

HRESULT
AnyElementProvider::get_HostRawElementProvider(IRawElementProviderSimple** pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementprovidersimple-get_hostrawelementprovider)
  if (!pRetVal) return E_INVALIDARG;
  *pRetVal = nullptr;
  return S_OK;
}

HRESULT
AnyElementProvider::get_ProviderOptions(ProviderOptions* pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementprovidersimple-get_provideroptions)
  return ProviderOptions_ServerSideProvider | ProviderOptions_UseComThreading;
}

HRESULT
AnyElementProvider::GetPatternProvider(PATTERNID patternId, IUnknown** pRetVal) {
  log("%s %d\n", __func__, patternId);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementprovidersimple-getpatternprovider)
  if (!pRetVal) return E_POINTER;
  *pRetVal = nullptr;

  auto index = ui_get_index(this->id);
  VERIFY(g_ui.node_type[index] != UiTree::Type::kNone);

  auto type = g_ui.node_type[index];
  char const* pattern = nullptr;
  switch (type) {
  case UiTree::Type::kDocument: {
    if (patternId == UIA_TextPatternId) {
      *pRetVal = create_element_text_provider(this->id);
    }
    else if (patternId == UIA_ValuePatternId) {
      *pRetVal = create_element_value_provider(this->id);
    }
  case UiTree::Type::kText: {
    if (patternId == UIA_TextPatternId) {
      *pRetVal = create_element_text_provider(this->id);
    }
    if (false) {
      // If I implement the value pattern, then narrator says "<name of item> Text <name of item>" which is redundant. (at least in scan mode)
      if (patternId == UIA_ValuePatternId) {
        *pRetVal = create_element_value_provider(this->id);
      }
    }
  }
  } break;

  case UiTree::Type::kButton : {
    if (patternId == UIA_InvokePatternId) {
      *pRetVal = create_element_invoke_provider(this->id);
    }
  } break;
  }

  switch (patternId) {
  case UIA_ValuePatternId: { pattern = "Value";  } break;
  case UIA_TextPatternId: { pattern = "Text";  } break;
  case UIA_InvokePatternId: { pattern = "Invoke"; } break;
  case UIA_ExpandCollapsePatternId: { pattern = "ExpandCollapse"; } break;
  case UIA_GridItemPatternId: { pattern = "GridItem"; } break;
  case UIA_GridPatternId: { pattern = "Grid";  } break;
  case UIA_RangeValuePatternId: { pattern = "RangeValue";  } break;
  case UIA_ScrollItemPatternId: { pattern = "ScrollItem";  } break;
  case UIA_ScrollPatternId: { pattern = "Scroll";  } break;
  case UIA_SelectionItemPatternId: { pattern = "SelectionItem"; } break;
  case UIA_SelectionPatternId: { pattern = "Selection"; } break;
  case UIA_TableItemPatternId: { pattern = "TableItem"; } break;
  case UIA_TablePatternId: { pattern = "Table"; } break;
  case UIA_TogglePatternId: { pattern = "Toggle"; } break;
  case UIA_WindowPatternId: { pattern = "Window"; } break;
  case UIA_TextChildPatternId: { pattern = "TextChild";  } break;
  case UIA_DragPatternId: { pattern = "Drag";  } break;
  case UIA_SpreadsheetItemPatternId: { pattern = "SpreadsheetItemPattern";  } break;
  }

  if (!*pRetVal) {
    log("  %s pattern not supported.\n", pattern? pattern : "");
  }
  else {
    log("  %s pattern supported.\n", pattern);
  }

  return S_OK;
}

HRESULT
AnyElementProvider::GetPropertyValue(PROPERTYID propertyId, VARIANT* pRetVal) {
  log("%s(%d)\n", __func__, propertyId);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementprovidersimple-getpropertyvalue)
  if (!pRetVal) return E_POINTER;

  pRetVal->vt = VT_EMPTY;

  auto index = ui_get_index(this->id);
  VERIFY(g_ui.node_type[index] != UiTree::Type::kNone);

  auto type = g_ui.node_type[index];

  char const* propname = nullptr;

  switch (propertyId) {
  case UIA_NamePropertyId: {
    pRetVal->vt = VT_BSTR;
    pRetVal->bstrVal = ::SysAllocString(g_ui.node_names[index].data());
    propname = "Name";
  } break;

  case UIA_ControlTypePropertyId: {
    pRetVal->vt = VT_I4;
    switch (type) {
    case UiTree::Type::kText: { pRetVal->lVal = UIA_TextControlTypeId; } break;
    case UiTree::Type::kDocument: { pRetVal->lVal = UIA_DocumentControlTypeId; } break;
    case UiTree::Type::kButton: { pRetVal->lVal = UIA_ButtonControlTypeId;  } break;
    default: VERIFY(0); // Implement this missing type
    }
    propname = "ControlType";
  } break;

  case UIA_IsControlElementPropertyId: {
    propname = "IsControlElement";
    pRetVal->vt = VT_BOOL;
    pRetVal->boolVal = VARIANT_TRUE;
  } break;

  case UIA_IsContentElementPropertyId: {
    propname = "IsContentElement";
    pRetVal->vt = VT_BOOL;
    pRetVal->boolVal = VARIANT_TRUE;
  } break;
  case UIA_IsEnabledPropertyId: {
    pRetVal->vt = VT_BOOL;
    pRetVal->boolVal = VARIANT_TRUE;
    propname = "IsEnabled";
  } break;
  case UIA_IsKeyboardFocusablePropertyId: {
    pRetVal->vt = VT_BOOL;
    pRetVal->boolVal = VARIANT_TRUE;
    propname = "IsKeyboardFocusable";
  } break;

  case UIA_LabeledByPropertyId: {
      if (type == UiTree::Type::kDocument) {
          pRetVal->vt = VT_BSTR;
          pRetVal->bstrVal = ::SysAllocString(g_ui.node_names[index].data());
          propname = "LabeledBy";
      }
  } break;

  case UIA_NativeWindowHandlePropertyId: {
    propname = "NativeWindowHandle";
  } break;
  case UIA_FrameworkIdPropertyId: { propname = "FrameworkId";  } break;
  case UIA_AutomationIdPropertyId: { propname = "AutomationId";  } break;
  case UIA_ProcessIdPropertyId: { propname = "ProcessId";  } break;
  case UIA_HelpTextPropertyId: { propname = "HelpText";  } break;
  case UIA_AccessKeyPropertyId: { propname = "AccessKey"; } break;

  case UIA_ProviderDescriptionPropertyId: {
    pRetVal->vt = VT_BSTR;
    pRetVal->bstrVal = ::SysAllocString(L"UU::RootProvider");
    propname = "ProviderDescription";
  } break;
  case UIA_ClassNamePropertyId: {
    pRetVal->vt = VT_BSTR;
    pRetVal->bstrVal = ::SysAllocString(L"UU::RootProvider");
    propname = "ClassNameDescription";
  } break;

  }

  if (pRetVal->vt != VT_EMPTY) {
    log("  supported_property %s\n", propname);
  }
  else {
    if (propname) { log("  unsupported_property %s\n", propname); }
  }
  
  return S_OK;
}

HRESULT
AnyElementProvider::get_BoundingRectangle(UiaRect* pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-get_boundingrectangle)
  if (!pRetVal) return E_INVALIDARG;
  RECT ClientRect;
  VERIFY(::GetClientRect(g_hwnd, &ClientRect));
  POINT LeftTop = { .x = ClientRect.left, .y = ClientRect.top };
  VERIFY(::ClientToScreen(g_hwnd, &LeftTop));
  pRetVal->left = double(LeftTop.x);
  pRetVal->top = double(LeftTop.y);
  pRetVal->width = double(ClientRect.right) - ClientRect.left;
  pRetVal->height = double(ClientRect.bottom) - ClientRect.top;
  return S_OK;
}

HRESULT
AnyElementProvider::get_FragmentRoot(IRawElementProviderFragmentRoot** pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-get_fragmentroot)
  if (!pRetVal) return E_INVALIDARG;
  g_root_provider->AddRef();
  *pRetVal = g_root_provider;
  return S_OK;
}

HRESULT
AnyElementProvider::GetEmbeddedFragmentRoots(SAFEARRAY** pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-getembeddedfragmentroots)
  if (!pRetVal) return E_INVALIDARG;
  *pRetVal = nullptr;
  return S_OK;
}

HRESULT
AnyElementProvider::GetRuntimeId(SAFEARRAY** pRetVal) {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-getruntimeid)
  if (!pRetVal) return E_INVALIDARG;

  LONG ids[] = { UiaAppendRuntimeId, LONG(bits(this->id, 0, 32)) };
  auto num_ids = sizeof ids / sizeof ids[0];

  log("  id: UiAppendRuntimeId.%#llx\n", int(ids[1]));

  SAFEARRAY* psa = ::SafeArrayCreateVector(VT_I4, 0, LONG(num_ids));
  if (psa == NULL) {
    return E_OUTOFMEMORY;
  }

  for (size_t i = 0; i < num_ids; i++) {
    LONG idx = (LONG)i;
    VERIFYHR(::SafeArrayPutElement(psa, &idx, &(ids[i])));
  }

  *pRetVal = psa;
  return S_OK;
}

HRESULT
AnyElementProvider::Navigate(NavigateDirection direction, IRawElementProviderFragment** pRetVal) {
  log("%s %d\n", __func__, direction);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-navigate)
  if (!pRetVal) return E_INVALIDARG;
  *pRetVal = nullptr;

  UiTree::Id element_id = -1;

  auto index = ui_get_index(this->id);
  auto this_depth = g_ui.node_depth[index];
  auto this_parent = g_ui.node_parent[index];
  auto navtype = "unknown";
  switch (direction) {
  case NavigateDirection_Parent: {
      navtype = "parent";
      element_id = g_ui.node_parent[index];
      VERIFY(element_id == this_parent);
  } break;
  case NavigateDirection_NextSibling: {
    navtype = "next-sibling";
    for (auto i = index + 1; i < g_ui.node_ids.size() && g_ui.node_depth[i] >= this_depth; i++) {
      if (g_ui.node_depth[i] == this_depth) {
        index = i;
        break;
      }
    }
    element_id = g_ui.node_ids[index];
    VERIFY(element_id == this->id || g_ui.node_parent[index] == this_parent);
  } break;
  case NavigateDirection_PreviousSibling: { 
    navtype = "prev-sibling";
    for (auto ri = index; ri > 0 && g_ui.node_depth[ri-1] >= this_depth; ri--) {
      auto i = ri - 1;
      if (g_ui.node_depth[i] == this_depth) {
        index = i;
        break;
      }
    }
    element_id = g_ui.node_ids[index];
    VERIFY(element_id == this->id || g_ui.node_parent[index] == this_parent);
  } break;
  case NavigateDirection_FirstChild: {
      navtype = "first-child";
      for (auto i = index + 1; i < g_ui.node_ids.size() && g_ui.node_depth[i] >= this_depth + 1; i++) {
        if (g_ui.node_depth[i] == this_depth + 1) {
          index = i;
          VERIFY(g_ui.node_parent[index] == this->id);
          break;
        }
      }
      element_id = g_ui.node_ids[index];
      VERIFY(element_id == this->id || g_ui.node_parent[index] == this->id);
  } break;
  case NavigateDirection_LastChild: {
      navtype = "last-child";
      for (auto i = index + 1; i < g_ui.node_ids.size() && g_ui.node_depth[i] >= this_depth + 1; i++) {
        if (g_ui.node_depth[i] == this_depth + 1) {
          index = i;
          VERIFY(g_ui.node_parent[index] == this->id);
        }
      }
      element_id = g_ui.node_ids[index];
      VERIFY(element_id == this->id || g_ui.node_parent[index] == this->id);
  } break;
  }

  log("  Navigating (%s) from element %#llx to %#llx\n", navtype, this->id, element_id);

  if (element_id == 0) {
    g_root_provider->AddRef();
    *pRetVal = g_root_provider;
  } else if (valid_id(element_id) && element_id != this->id) {
    *pRetVal = create_element_provider(element_id);
  }

  return S_OK;
}

HRESULT
AnyElementProvider::SetFocus() {
  log("%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-setfocus)
  ui_set_focus_to(this->id);
  return S_OK;
}

struct AnyElementValueProvider : public IValueProvider {
  // IValueProvider:
  HRESULT STDMETHODCALLTYPE get_IsReadOnly(BOOL* pRetVal);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-ivalueprovider-get_isreadonly)
  HRESULT STDMETHODCALLTYPE get_Value(BSTR* pRetVal);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-ivalueprovider-get_value)
  HRESULT STDMETHODCALLTYPE SetValue(LPCWSTR val);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-ivalueprovider-setvalue)

  // IUnknown interface:
  ULONG STDMETHODCALLTYPE AddRef() override { return ++reference_count; }
  ULONG STDMETHODCALLTYPE Release() override { return --reference_count; }
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;

  void log(char const* fmt, ...) {
    ::log("this(%p, id=%#llx) AnyElementValueProvider::", this, this->id);
    std::va_list args;
    va_start(args, fmt);
    ::logv(fmt, args);
    va_end(args);
  }

  ULONG reference_count = 1;
  UiTree::Id id = -1;
};

HRESULT
AnyElementValueProvider::get_IsReadOnly(BOOL* pRetVal) {
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-ivalueprovider-get_isreadonly)
  log("%s\n", __func__);
  if (!pRetVal) return E_INVALIDARG;
  *pRetVal = TRUE;
  return S_OK;
}

HRESULT
AnyElementValueProvider::get_Value(BSTR* pRetVal) {
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-ivalueprovider-get_value)
  log("%s\n", __func__);
  if (!pRetVal) return E_INVALIDARG;
  *pRetVal = ::SysAllocString(g_ui.node_names[ui_get_index(this->id)].data());
  return S_OK;
}

HRESULT
AnyElementValueProvider::SetValue(LPCWSTR val) {
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-ivalueprovider-setvalue)
  log("%s\n", __func__);
  return E_ACCESSDENIED;
}

HRESULT
AnyElementValueProvider::QueryInterface(REFIID riid, void** ppvObject) {
  auto result = [&]() -> std::pair<char const*, void*> {
    // supported interfaces:
    if (riid == IID_IValueProvider) return { "IValueProvider", static_cast<IValueProvider*>(this) };
    return { nullptr, nullptr };
  }();

  if (riid == IID_IUnknown) {
    // should always be supported
    result = { "IUnknown", this };
  }

  LPOLESTR riid_string;
  VERIFYHR(::StringFromIID(riid, &riid_string));
  log("%s %u (%ls)\n", __func__, riid, riid_string);
  if (!ppvObject) return E_POINTER;

  *ppvObject = result.second; // it's important to also assign it a value even if E_NOTINTERFACE
  if (!result.second) {
    if (result.first) { log("  missing %s interface (not supported)\n", result.first); }
    return E_NOINTERFACE;
  }

  AddRef();
  log("  supported_interface %s\n", result.first);
  return S_OK;
}



struct AnyElementTextProvider : public ITextProvider {
  // ITextProvider
  HRESULT STDMETHODCALLTYPE get_DocumentRange(ITextRangeProvider** pRetVal) override;
  HRESULT STDMETHODCALLTYPE get_SupportedTextSelection(SupportedTextSelection* pRetVal) override;
  HRESULT STDMETHODCALLTYPE GetSelection(SAFEARRAY** pRetVal) override;
  HRESULT STDMETHODCALLTYPE GetVisibleRanges(SAFEARRAY** pRetVal) override;
  HRESULT STDMETHODCALLTYPE RangeFromChild(IRawElementProviderSimple* childElement, ITextRangeProvider** pRetVal) override;
  HRESULT RangeFromPoint(UiaPoint point, ITextRangeProvider** pRetVal) override;

  // IUnknown interface:
  ULONG STDMETHODCALLTYPE AddRef() override { return ++reference_count; }
  ULONG STDMETHODCALLTYPE Release() override { return --reference_count; }
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;

  void log(char const* fmt, ...) {
    ::log("this(%p, id=%#llx) AnyElementTextProvider::", this, this->id);
    std::va_list args;
    va_start(args, fmt);
    ::logv(fmt, args);
    va_end(args);
  }

  ULONG reference_count = 1;
  UiTree::Id id = -1;
};

HRESULT 
AnyElementTextProvider::QueryInterface(REFIID riid, void** ppvObject) {
  auto result = [&]() -> std::pair<char const*, void*> {
    // supported interfaces:
    if (riid == IID_IValueProvider) return { "ITextProvider", (ITextProvider*)this };
    return { nullptr, nullptr };
  }();

  if (riid == IID_IUnknown) {
    // should always be supported
    result = { "IUnknown", this };
  }

  LPOLESTR riid_string;
  VERIFYHR(::StringFromIID(riid, &riid_string));
  log("%s %u (%ls)\n", __func__, riid, riid_string);
  if (!ppvObject) return E_POINTER;

  *ppvObject = result.second; // it's important to also assign it a value even if E_NOTINTERFACE
  if (!result.second) {
    if (result.first) { log("  missing %s interface (not supported)\n", result.first); }
    return E_NOINTERFACE;
  }

  AddRef();
  log("  supported_interface %s\n", result.first);
  return S_OK;
}

HRESULT
AnyElementTextProvider::get_DocumentRange(ITextRangeProvider** pRetVal) {
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-itextprovider-get_documentrange)
  log("%s (unsupported)\n", __func__);
  if (!pRetVal) return E_INVALIDARG;
  *pRetVal = nullptr;
  return S_OK;
}

HRESULT 
AnyElementTextProvider::get_SupportedTextSelection(SupportedTextSelection* pRetVal) {
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-itextprovider-get_supportedtextselection)
  log("%s\n", __func__);
  if (!pRetVal) return E_INVALIDARG;

  *pRetVal = SupportedTextSelection_None;
  return S_OK;
}

HRESULT
AnyElementTextProvider::GetSelection(SAFEARRAY** pRetVal) {
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-itextprovider-getselection)
  log("%s\n", __func__);
  if (!pRetVal) return E_INVALIDARG;

  *pRetVal = nullptr;
  return S_OK;
}

HRESULT
AnyElementTextProvider::GetVisibleRanges(SAFEARRAY** pRetVal) {
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-itextprovider-getvisibleranges)
  log("%s\n", __func__);
  if (!pRetVal) return E_INVALIDARG;

  // TODO(nil): we need to provide ranges here, which is possible once we implement a ITextRangeProvider.

  *pRetVal = nullptr;
  return E_NOTIMPL;

}

HRESULT
AnyElementTextProvider::RangeFromChild(IRawElementProviderSimple* childElement, ITextRangeProvider** pRetVal) {
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-itextprovider-rangefromchild)
  log("%s\n", __func__);
  if (!childElement) return E_INVALIDARG;
  if (!pRetVal) return E_INVALIDARG;

  auto p = dynamic_cast<AnyElementProvider*>(childElement);
  log("  id=%#llx\n", p->id);
  *pRetVal = nullptr;
  return E_NOTIMPL;
}

HRESULT
AnyElementTextProvider::RangeFromPoint(UiaPoint point, ITextRangeProvider** pRetVal) {
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-itextprovider-rangefrompoint)
  log("%s\n", __func__);
  if (!pRetVal) return E_INVALIDARG;

  log("  {%f %f}\n", point.x, point.y);
  *pRetVal = nullptr;
  return E_NOTIMPL;
}

struct AnyElementInvokeProvider : public IInvokeProvider {
  // IInvokeProvider interface:
  HRESULT STDMETHODCALLTYPE Invoke() override;

  // IUnknown interface:
  ULONG STDMETHODCALLTYPE AddRef() override { return ++reference_count; }
  ULONG STDMETHODCALLTYPE Release() override { return --reference_count; }
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;

  void log(char const* fmt, ...) {
    ::log("this(%p, id=%#llx) AnyElementInvokeProvider::", this, this->id);
    std::va_list args;
    va_start(args, fmt);
    ::logv(fmt, args);
    va_end(args);
  }

  ULONG reference_count = 1;
  UiTree::Id id = -1;
};

HRESULT 
AnyElementInvokeProvider::Invoke() {
  log("%s\n", __func__);
  ui_activate(this->id);
  return S_OK;
}

HRESULT
AnyElementInvokeProvider::QueryInterface(REFIID riid, void** ppvObject) {
  auto result = [&]() -> std::pair<char const*, void*> {
    // supported interfaces:
    if (riid == IID_IInvokeProvider) return { "IInvokeProvider", static_cast<IInvokeProvider*>(this) };
    return { nullptr, nullptr };
  }();

  if (riid == IID_IUnknown) {
    // should always be supported
    result = { "IUnknown", this };
  }

  LPOLESTR riid_string;
  VERIFYHR(::StringFromIID(riid, &riid_string));
  log("%s %u (%ls)\n", __func__, riid, riid_string);
  if (!ppvObject) return E_POINTER;

  *ppvObject = result.second; // it's important to also assign it a value even if E_NOTINTERFACE
  if (!result.second) {
    if (result.first) { log("  missing %s interface (not supported)\n", result.first); }
    return E_NOINTERFACE;
  }

  AddRef();
  log("  supported_interface %s\n", result.first);
  return S_OK;
}

IRawElementProviderFragment*
create_element_provider(UiTree::Id element_id) {
  VERIFY(exists_id(element_id));

  auto p = new AnyElementProvider;
  p->id = element_id;

  return p;
}

ITextProvider*
create_element_text_provider(UiTree::Id element_id) {
  VERIFY(exists_id(element_id));

  auto p = new AnyElementTextProvider;
  p->id = element_id;

  return p;
}

IValueProvider*
create_element_value_provider(UiTree::Id element_id) {
  VERIFY(exists_id(element_id));

  auto p = new AnyElementValueProvider;
  p->id = element_id;

  return p;
}

IInvokeProvider*
create_element_invoke_provider(UiTree::Id element_id) {
  VERIFY(exists_id(element_id));

  auto p = new AnyElementInvokeProvider;
  p->id = element_id;
  return p;
}

UiTree::Id
ui_named_element(wchar_t const* name, UiTree::Type type) {
  auto index = g_ui.node_ids.size();
  auto depth = g_ui.depth_for_adding_element;

  UiTree::Id parent_id = -1;
  if (depth == 0) {
    parent_id = 0;
  }
  else {
    auto parent_depth = depth - 1;
    auto parent_pos = std::find(g_ui.node_depth.rbegin(), g_ui.node_depth.rend(), parent_depth);
    VERIFY(parent_pos != g_ui.node_depth.rend());
    auto relative_index = std::distance(g_ui.node_depth.rbegin(), parent_pos);
    auto parent_index = index - 1 - relative_index;
    parent_id = g_ui.node_ids[parent_index];
  }

  auto num_bytes = wcslen(name) * sizeof name[0];
  auto id = hash(num_bytes, name);
  id = wyhash64(id, parent_id);

  log("%*snode: %d %#llx (%ls)\n", int(depth), "", type, id, name);
  VERIFY(valid_id(id));
  VERIFY(g_ui.node_ids.end() == std::find(g_ui.node_ids.begin(), g_ui.node_ids.end(), id));
  g_ui.node_ids.push_back(id);
  g_ui.node_names.push_back(name);
  g_ui.node_type.push_back(type);
  g_ui.node_depth.push_back(depth);
  g_ui.node_parent.push_back(parent_id);
  return id;
}

UiTree::Id
ui_document(wchar_t const* text) {
    return ui_named_element(text, UiTree::Type::kDocument);
}

UiTree::Id
ui_text_paragraph(wchar_t const* text) {
    return ui_named_element(text, UiTree::Type::kText);
}

UiTree::Id
ui_button(wchar_t const* text, std::function<void()> action) {
  auto id = ui_named_element(text, UiTree::Type::kButton);
  VERIFY(g_ui.actions.insert_or_assign(id, action).second);
  return id;
}

void
ui_describe() {
  log("ui_describe: START\n");
  UiTree::Id fid = {};
  ui_document(L"Main");
  {
    g_ui.depth_for_adding_element++;

    ui_text_paragraph(L"This is the first paragraph.");
    fid = ui_text_paragraph(L"Hello, Dreamer of dreams.");
    ui_text_paragraph(L"Yet another paragraph");
 
    g_ui.depth_for_adding_element--;
  }
  ui_button(L"Minimize Application", []() { VERIFY(::CloseWindow(g_hwnd)); });
  ui_button(L"Close Application", []() { ::SendMessage(g_hwnd, WM_CLOSE, 0, 0); }); // A thread cannot use DestroyWindow to destroy a window created by a different thread.

  log("ui_describe: END\n");

  log("g_ui.node_ids.size() = %zu\n", g_ui.node_ids.size());

  // Initialize focus
  if (g_ui.focused_id == 0 && !g_ui.node_ids.empty()) {
      ui_set_focus_to(fid);
  }
}

size_t
ui_get_index(UiTree::Id id) {
  VERIFY(valid_id(id));
  static struct {
    UiTree::Id id[2];
    size_t index[2];
  } fingers;
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

// TODO(nil): embed in single describe function.
void
ui_focus_next() {
  auto index = ui_get_index(g_ui.focused_id);
  if (index >= g_ui.node_ids.size() - 1) return;
  
  index++;
  ui_set_focus_to(g_ui.node_ids[index]);
}

void
ui_focus_prev() {
  auto index = ui_get_index(g_ui.focused_id);
  if (index <= 0) return;
  
  index--;
  ui_set_focus_to(g_ui.node_ids[index]);
}

void
ui_activate() {
  if (g_ui.focused_id) ui_activate(g_ui.focused_id);
}

void
ui_set_focus_to(UiTree::Id id) {
    if (id == g_ui.focused_id) {
        log("redundant ui_set_focus_to\n");
        return;
    }
    log("changing focus from %#llx to %#llx\n", g_ui.focused_id, id);
    g_ui.focused_id = id;
    if (UiaClientsAreListening() && g_root_provider) {
        auto p = create_element_provider(g_ui.focused_id);
        IRawElementProviderSimple* sp;
        VERIFYHR(p->QueryInterface<IRawElementProviderSimple>(&sp));      
        VERIFYHR(UiaRaiseAutomationEvent(sp, UIA_AutomationFocusChangedEventId));
        p->Release();
        sp->Release();
    }
}

void
ui_activate(UiTree::Id id) {
  log("activating %#llx\n", id);
  auto action_pos = g_ui.actions.find(id);
  VERIFY(action_pos != g_ui.actions.end());

  auto fn = action_pos->second;
  fn();
}