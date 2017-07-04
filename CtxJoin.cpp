// CtxJoin.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "resource.h"
#include <initguid.h>

#include "ContextMenu.h"


CShellModule _Module;

const CLSID CLSID_ContextMenu = {0x2B3256E4,0x49DF,0x11D3,{0x82,0x29,0x00,0x80,0xAE,0x50,0x90,0x59}};


BEGIN_OBJECT_MAP(ObjectMap)
OBJECT_ENTRY(CLSID_ContextMenu, CContextMenu)
END_OBJECT_MAP()


/////////////////////////////////////////////////////////////////////////////
// DLL Entry Point

extern "C"
BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID /*lpReserved*/)
{
   if (dwReason == DLL_PROCESS_ATTACH) {
      _Module.Init(ObjectMap, hInstance);
      ::DisableThreadLibraryCalls(hInstance);
   }
   else if (dwReason == DLL_PROCESS_DETACH) {
      _Module.Term();
   }
   return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// Used to determine whether the DLL can be unloaded by OLE

STDAPI DllCanUnloadNow(void)
{
   return (_Module.GetLockCount()==0) ? S_OK : S_FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// Returns a class factory to create an object of the requested type

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
   return _Module.GetClassObject(rclsid, riid, ppv);
}

/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
#ifdef _DEBUG
   // During DEBUG I have found it neat to kill the Explorer
   // during DLL registration under W2K.
   // It is much easier to debug this way, because once you
   // start tracing (running) you effectively debug the
   // first instance of Explorer.exe!
   ::PostMessage(::FindWindow(_T("Progman"), NULL), WM_QUIT, 0, 0L);
#endif
   // registers object, typelib and all interfaces in typelib
   return _Module.RegisterServer();
}

/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
   return _Module.UnregisterServer();
}



// IDropTarget

STDMETHODIMP CContextMenu::DragEnter(LPDATAOBJECT pDataObj, DWORD dwKeyState, POINTL ptl, LPDWORD pdwEffect)
{
   return S_OK;
}

STDMETHODIMP CContextMenu::DragOver(DWORD dwKeyState, POINTL ptl, LPDWORD pdwEffect)
{
   return S_OK;
}

STDMETHODIMP CContextMenu::DragLeave(VOID)
{
   return S_OK;
}

STDMETHODIMP CContextMenu::Drop(LPDATAOBJECT pDataObj, DWORD dwKeyState, POINTL ptl, LPDWORD pdwEffect)
{
   return S_OK;       // Did we just deny this drop?

}

// IPersist

STDMETHODIMP CContextMenu::GetClassID(CLSID* pClassID)
{
   return S_OK;
}

// IPersistFile

STDMETHODIMP CContextMenu::IsDirty()
{
   ATLTRACENOTIMPL(L"CDropTarget::IsDirty");
}

STDMETHODIMP CContextMenu::Load(LPCOLESTR pwstrFileName, DWORD dwMode)
{
	ATLTRACENOTIMPL(L"CDropTarget::Load");
}

STDMETHODIMP CContextMenu::Save(LPCOLESTR, BOOL fRemember)
{
   ATLTRACENOTIMPL(L"CDropTarget::Save");
}

STDMETHODIMP CContextMenu::SaveCompleted(LPCOLESTR pszFileName)
{
   ATLTRACENOTIMPL(L"CDropTarget::SaveCompleted");
}

STDMETHODIMP CContextMenu::GetCurFile(LPOLESTR* ppszFileName)
{
   ATLTRACENOTIMPL(L"CDropTarget::GetCurFile");
}   

