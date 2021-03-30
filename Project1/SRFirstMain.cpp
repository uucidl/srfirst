// # SRFirst "Screen Reader First"
//
// An experiment in designing an app starting first from screen-reader support, before thinking about the GUI.
//

// First test it with the Accessibility Insight for Windows app, then test it with a screen-reader such as NVDA, Jaws or Narrator.

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

#include <cstdarg>
#include <cstdio>
#include <cstdlib>
#include <string>
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
log(char const* fmt, ...) {
  std::va_list args;
  va_start(args, fmt);
  std::vfprintf(g_output_log, fmt, args);
  va_end(args);
  std::fflush(g_output_log);
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
 
  ui_describe();

  for (;;) {
    MSG msg;
    switch (::GetMessageW(&msg, Window, 0, 0)) {
    case -1: VERIFY(0); break;
    case 0: goto end; // WM_QUIT was received.
    default: break;
    }
    ::DispatchMessageW(&msg);
  }
end:
  VERIFY(::DestroyWindow(g_hwnd));
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
    case WM_CLOSE: ::PostQuitMessage(0); return 0;
    //  2.1- Menu handling. URL(https://docs.microsoft.com/en-us/windows/win32/menurc/about-menus#messages-used-with-menus)
    case WM_COMMAND: {
      switch ((MenuId)LOWORD(wParam)) {
      case MenuId_File_Quit: ::PostQuitMessage(0); return 0;
      }
      
    } break;
    case WM_DESTROY: {
      // Microsoft recommends making this call from the WM_DESTROY message handler of the window that returns the UI Automation providers.
      // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcoreapi/nf-uiautomationcoreapi-uiareturnrawelementprovider)
      return ::UiaReturnRawElementProvider(hwnd, 0, 0, NULL);
    } break;
    case WM_GETOBJECT: {
      switch ((DWORD)lParam) {
      case UiaRootObjectId: {
        log("Hello, we just received a WM_GETOBJECT message for UiaAutomation with params: %u %u\n", wParam, lParam);
        if (!g_root_provider) {
          g_root_provider = new RootProvider;
        }
        return ::UiaReturnRawElementProvider(hwnd, wParam, lParam, g_root_provider);
      } break;

      default: break;
      }
    } 
    case WM_KEYDOWN: {
      switch (wParam) {
      case VK_TAB: {
        BYTE Keys[256];
        VERIFY(::GetKeyboardState(Keys));
        if (Keys[VK_SHIFT] & (1<<7)) {
          ui_focus_prev();
        } else {
          ui_focus_next();
        }
      } break;
      }

    } break;
    case WM_SETFOCUS: {
      log("WM_SETFOCUS received\n");
    } break;
  }
  return ::DefWindowProcW(hwnd, uMsg, wParam, lParam);
}

struct UiTree {
  using Id = std::uint64_t;
  // Id == -1 => invalid_id
  // Id == 0  => root
  // ...

  std::vector<Id> node_ids; // in presentation order.
  std::vector<std::wstring> node_names;

  Id focused_id = 0;
};

bool
valid_id(UiTree::Id id) {
  return 0 < id && id < UiTree::Id(-1);
}

static UiTree g_ui;

IRawElementProviderFragment*
create_element_provider(UiTree::Id element_id);

HRESULT
RootProvider::QueryInterface(REFIID riid, void** ppvObject) {
  log("%s %u\n", __func__, riid);
  if (!ppvObject) return E_POINTER;
  void* result = nullptr;
  if (riid == IID_IUnknown) { result = this; }
  if (riid == IID_IRawElementProviderSimple) { result = (IRawElementProviderSimple*)this; }
  if (riid == IID_IRawElementProviderFragmentRoot) { result = (IRawElementProviderFragmentRoot*)this; }
  if (riid == IID_IRawElementProviderFragment) { result = (IRawElementProviderFragment*)this; }

  if (!result) return E_NOINTERFACE;

  log("  supported_interface\n");
  AddRef();
  *ppvObject = result;
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
  log("%s\n", __func__);
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
    if (!g_ui.node_ids.empty()) { element_id = g_ui.node_ids.front(); }
  } break;
  case NavigateDirection_LastChild: {
    log("  last-child(Root)\n");
    if (!g_ui.node_ids.empty()) { element_id = g_ui.node_ids.back(); }
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
  return S_OK;
}

HRESULT
RootProvider::GetFocus(IRawElementProviderFragment** pRetVal) {
  log("%s\n", __func__);
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

  this->AddRef();
  *pRetVal = (IRawElementProviderFragment*)this;

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

  ULONG reference_count = 1;
  UiTree::Id id = -1;
};


HRESULT
AnyElementProvider::QueryInterface(REFIID riid, void** ppvObject) {
  log("%s %u\n", __func__, riid);
  if (!ppvObject) return E_POINTER;
  void* result = nullptr;
  if (riid == IID_IUnknown) { result = this; }
  if (riid == IID_IRawElementProviderSimple) { result = (IRawElementProviderSimple*)this; }
  if (riid == IID_IRawElementProviderFragment) { result = (IRawElementProviderFragment*)this; }

  if (!result) return E_NOINTERFACE;

  log("  supported_interface\n");
  AddRef();
  *ppvObject = result;
  return S_OK;
}

HRESULT
AnyElementProvider::get_HostRawElementProvider(IRawElementProviderSimple** pRetVal) {
  log("AnyElementProvider::%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementprovidersimple-get_hostrawelementprovider)
  if (!pRetVal) return E_INVALIDARG;
  *pRetVal = nullptr;
  return S_OK;
}

HRESULT
AnyElementProvider::get_ProviderOptions(ProviderOptions* pRetVal) {
  log("AnyElementProvider::%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementprovidersimple-get_provideroptions)
  return ProviderOptions_ServerSideProvider | ProviderOptions_UseComThreading;
}

HRESULT
AnyElementProvider::GetPatternProvider(PATTERNID patternId, IUnknown** pRetVal) {
  log("AnyElementProvider::%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementprovidersimple-getpatternprovider)
  if (!pRetVal) return E_POINTER;
  *pRetVal = nullptr;
  return S_OK;
}

HRESULT
AnyElementProvider::GetPropertyValue(PROPERTYID propertyId, VARIANT* pRetVal) {
  log("AnyElementProvider::%s(%d)\n", __func__, propertyId);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementprovidersimple-getpropertyvalue)
  if (!pRetVal) return E_POINTER;

  pRetVal->vt = VT_EMPTY;

  size_t index;
  for (index = 0; index < g_ui.node_ids.size(); index++) {
    if (g_ui.node_ids[index] == this->id) break;
  }
  VERIFY(index < g_ui.node_ids.size()); // is this a case that needs instead to be legitimately handled, like if we have elements that disappear?
  
  switch (propertyId) {
  case UIA_NamePropertyId: {
    pRetVal->vt = VT_BSTR;
    pRetVal->bstrVal = ::SysAllocString(g_ui.node_names[index].data());
  } break;

  case UIA_ControlTypePropertyId: {
    pRetVal->vt = VT_I4;
    pRetVal->iVal = UIA_TextControlTypeId; // For now we only have text elements. This should change very soon.
  } break;

  case UIA_IsEnabledPropertyId: {
    pRetVal->vt = VT_BOOL;
    pRetVal->boolVal = VARIANT_TRUE;
  } break;
  case UIA_IsKeyboardFocusablePropertyId: {
    pRetVal->vt = VT_BOOL;
    pRetVal->boolVal = VARIANT_TRUE;
  } break;
  }
  return S_OK;
}

HRESULT
AnyElementProvider::get_BoundingRectangle(UiaRect* pRetVal) {
  log("AnyElementProvider::%s\n", __func__);
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
  log("AnyElementProvider::%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-get_fragmentroot)
  if (!pRetVal) return E_INVALIDARG;
  g_root_provider->AddRef();
  *pRetVal = g_root_provider;
  return S_OK;
}

HRESULT
AnyElementProvider::GetEmbeddedFragmentRoots(SAFEARRAY** pRetVal) {
  log("AnyElementProvider::%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-getembeddedfragmentroots)
  if (!pRetVal) return E_INVALIDARG;
  *pRetVal = nullptr;
  return S_OK;
}

HRESULT
AnyElementProvider::GetRuntimeId(SAFEARRAY** pRetVal) {
  log("AnyElementProvider::%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-getruntimeid)
  if (!pRetVal) return E_INVALIDARG;

  std::int32_t ids[] = { UiaAppendRuntimeId, std::int32_t(bits(this->id, 0, 32)) };
  auto num_ids = sizeof ids / sizeof ids[0];

  SAFEARRAY* psa = ::SafeArrayCreateVector(VT_I4, 0, num_ids);
  if (psa == NULL) {
    return E_OUTOFMEMORY;
  }

  for (size_t i = 0; i < num_ids; i++) {
    LONG idx = i;
    ::SafeArrayPutElement(psa, &idx, &(ids[i]));
  }

  *pRetVal = psa;
  return S_OK;
}

HRESULT
AnyElementProvider::Navigate(NavigateDirection direction, IRawElementProviderFragment** pRetVal) {
  log("AnyElementProvider::%s %d\n", __func__, direction);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-navigate)
  if (!pRetVal) return E_INVALIDARG;
  *pRetVal = nullptr;

  UiTree::Id element_id = -1;

  size_t index;
  for (index = 0; index < g_ui.node_ids.size(); index++) {
    if (g_ui.node_ids[index] == this->id) break;
  }
  log("AnyElementProvider::Navigate: found id %#llx at %zu\n", this->id, index);
  VERIFY(index < g_ui.node_ids.size()); // is this a case that needs instead to be legitimately handled, like if we have elements that disappear?

  switch (direction) {
  case NavigateDirection_Parent: { element_id = 0; } break;
  case NavigateDirection_NextSibling: {
    if (index < g_ui.node_ids.size() - 1) { index++; }
    element_id = g_ui.node_ids[index];
  } break;
  case NavigateDirection_PreviousSibling: { 
    if (index > 0) { index--;  }
    element_id = g_ui.node_ids[index];
  } break;
  }

  log("  Navigating from element %#llx to %#llx\n", this->id, element_id);

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
  log("AnyElementProvider::%s\n", __func__);
  // URL(https://docs.microsoft.com/en-us/windows/win32/api/uiautomationcore/nf-uiautomationcore-irawelementproviderfragment-setfocus)
  g_ui.focused_id = this->id;
  return S_OK;
}



IRawElementProviderFragment*
create_element_provider(UiTree::Id element_id) {
  VERIFY(valid_id(element_id));

  auto p = new AnyElementProvider;
  p->id = element_id;
  VERIFY(g_ui.node_ids.end() != std::find(g_ui.node_ids.begin(), g_ui.node_ids.end(), element_id));

  return p;
}

UiTree::Id
ui_text_paragraph(wchar_t const* text) {
  auto num_bytes = wcslen(text) * sizeof text[0];
  auto id = hash(num_bytes, text);
  log("text-paragraph: %#llx (%ls)\n", id, text);
  VERIFY(valid_id(id));
  g_ui.node_ids.push_back(id);
  g_ui.node_names.push_back(text);
  return id;
}

void
ui_describe() {
  log("ui_describe: START\n");
  ui_text_paragraph(L"This is the first paragraph.");
  auto fid = ui_text_paragraph(L"Hello, Dreamer of dreams.");
  ui_text_paragraph(L"Yet another paragraph");
  log("ui_describe: END\n");

  log("g_ui.node_ids.size() = %zu\n", g_ui.node_ids.size());

  // Initialize focus
  if (g_ui.focused_id == 0 && !g_ui.node_ids.empty()) {
    g_ui.focused_id = fid;
  }
}

// TODO(nil): embed in single describe function.
void
ui_focus_next() {
  size_t index;
  for (index = 0; index < g_ui.node_ids.size(); index++) {
    if (g_ui.node_ids[index] == g_ui.focused_id) break;
  }
  VERIFY(index < g_ui.node_ids.size()); // is this a case that needs instead to be legitimately handled, like if we have elements that disappear?

  if (index < g_ui.node_ids.size() - 1) { index++; }
  g_ui.focused_id = g_ui.node_ids[index];
  if (UiaClientsAreListening() && g_root_provider)
  {
    auto p = create_element_provider(g_ui.focused_id);
    IRawElementProviderSimple *sp;
    VERIFYHR(p->QueryInterface<IRawElementProviderSimple>(&sp));
    VERIFYHR(UiaRaiseAutomationEvent(sp, UIA_AutomationFocusChangedEventId)); // TODO(nil): this doesn't look enough, it seems NVDA doesn't pick up the focus change when I just do that.
  }
}

void
ui_focus_prev() {
  size_t index;
  for (index = 0; index < g_ui.node_ids.size(); index++) {
    if (g_ui.node_ids[index] == g_ui.focused_id) break;
  }
  VERIFY(index < g_ui.node_ids.size()); // is this a case that needs instead to be legitimately handled, like if we have elements that disappear?
  if (index > 0) { index--; }
  g_ui.focused_id = g_ui.node_ids[index];
  if (UiaClientsAreListening() && g_root_provider)
  {
    auto p = create_element_provider(g_ui.focused_id);
    IRawElementProviderSimple* sp;
    VERIFYHR(p->QueryInterface<IRawElementProviderSimple>(&sp));
    VERIFYHR(UiaRaiseAutomationEvent(sp, UIA_AutomationFocusChangedEventId));
  }
}

