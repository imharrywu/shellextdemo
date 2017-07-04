#if !defined(AFX_CONTEXTMENU_H__20010124_FB58_1C60_A0E5_0080AD509054__INCLUDED_)
#define AFX_CONTEXTMENU_H__20010124_FB58_1C60_A0E5_0080AD509054__INCLUDED_

#pragma once

#include "shlwapi.h"
#pragma comment(lib, "shlwapi.lib")


//////////////////////////////////////////////////////////////////////////////
// CContextMenu

// The IContextMenu for file context menu
class ATL_NO_VTABLE CContextMenu : 
   public CComObjectRootEx<CComSingleThreadModel>,
   public CComCoClass<CContextMenu, &CLSID_ContextMenu>,
   public CShellContextMenu<CContextMenu>,
   public IDropTarget,
   public IPersistFile
{
public:

DECLARE_NOT_AGGREGATABLE(CContextMenu)
DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CContextMenu)
   COM_INTERFACE_ENTRY_IID(IID_IDropTarget, IDropTarget)
   COM_INTERFACE_ENTRY_IID(IID_IPersistFile, IPersistFile)
   COM_INTERFACE_ENTRY_IID(IID_IContextMenu,IContextMenu)
   COM_INTERFACE_ENTRY_IID(IID_IShellExtInit, IShellExtInit)
END_COM_MAP()

BEGIN_CONTEXTMENU_MAP(CContextMenu)
   CONTEXTMENU_HANDLER(IDS_JOIN, IDS_DESC_JOIN, &CContextMenu::OnJoin)
   CONTEXTMENU_HANDLER(IDS_SPLIT, IDS_DESC_SPLIT, &CContextMenu::OnSplit)
END_CONTEXTMENU_MAP()

public:
   // IDropTarget

   STDMETHOD(DragEnter)(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);       
   STDMETHOD(DragOver)(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);      
   STDMETHOD(DragLeave)(VOID);
   STDMETHOD(Drop)(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);

   // IPersist

   STDMETHOD(GetClassID)(CLSID* pClassID);

   // IPersistFile

   STDMETHOD(IsDirty)(VOID);
   STDMETHOD(Load)(LPCOLESTR pwstrFileName, DWORD dwMode);
   STDMETHOD(Save)(LPCOLESTR pwstrFileName, BOOL fRemember);
   STDMETHOD(SaveCompleted)(LPCOLESTR pszFileName);
   STDMETHOD(GetCurFile)(LPOLESTR* ppszFileName);


// CContextMenu
public:
   static HRESULT WINAPI UpdateRegistry(BOOL bRegister)
   {
      CComBSTR bstrDescription;
      CComBSTR bstrCLSID(CLSID_ContextMenu);
      bstrDescription.LoadString(IDS_DESCRIPTION);
      _ATL_REGMAP_ENTRY rm[] = { 
         { OLESTR("CLSID"), bstrCLSID },
         { OLESTR("DESCRIPTION"), bstrDescription },
         { NULL,NULL } };
      return _Module.UpdateRegistryFromResource(IDR_CONTEXTMENU, bRegister, rm);
   }

// CContextMenu
public:

#define RETURN_ERROR { ::MessageBeep(-1); return; }

   void __stdcall OnJoin(void)
   {
      USES_CONVERSION;

      _RemoveFolders(m_arrFiles);
      if( m_arrFiles.GetSize()<2 ) RETURN_ERROR;
      _SortFiles(m_arrFiles);
     
      TCHAR szFirstFile[MAX_PATH];
      _tcscpy( szFirstFile, W2CT(m_arrFiles[0]) );
      
      TCHAR szFileName[MAX_PATH];
      _tcscpy( szFileName, szFirstFile );
      ::PathStripPath(szFileName);

      // Check if the joining files have the numbered (.001) extensions.
      // If so, strip them to get to actual filename or make a new extension.
      LPCTSTR pstrExt = ::PathFindExtension(szFirstFile);
      int iRes;
      if( *pstrExt!=_T('\0') && ::StrToIntEx(pstrExt+1, STIF_DEFAULT, &iRes) ) {
         ::PathRemoveExtension(szFileName);
      }
      else {
         ::PathAddExtension( szFileName, _T(".joined") );
      }

      TCHAR szTempFileName[MAX_PATH+1];
      ZeroMemory(szTempFileName, sizeof(szTempFileName)/sizeof(TCHAR)); // Need this because of ::SHFileOperation()
      CTemporaryFile out;
      if( out.Create(szTempFileName, MAX_PATH)==FALSE ) RETURN_ERROR;
      for( int i=0; i<m_arrFiles.GetSize(); i++ ) {
         CFile f;
         if( f.Open(W2CT(m_arrFiles[i]))==FALSE ) RETURN_ERROR;
         BYTE buf[1024];
         DWORD dwRead;
         while( f.Read(buf, sizeof(buf), &dwRead), dwRead>0 ) {
            out.Write(buf, dwRead);
         }
         f.Close();
      }
      out.Close();

      TCHAR szNewFileName[MAX_PATH+1];
      ZeroMemory(szNewFileName, sizeof(szNewFileName)/sizeof(TCHAR)); // Need this because of ::SHFileOperation()
      _tcscpy( szNewFileName, szFirstFile );
      ::PathRemoveFileSpec(szNewFileName);
      ::PathAppend(szNewFileName, szFileName);

      if( ::PathIsDirectory(szNewFileName) ) RETURN_ERROR;

      TCHAR szCaption[64];
      ::LoadString(_Module.GetResourceInstance(), IDS_FILECAPTION, szCaption, sizeof(szCaption)/sizeof(TCHAR));

      // Use ::SHFileOperation() to move the temporary file.
      // This way we get the "Confirm Overwrite" dialog if nessecary.
      SHFILEOPSTRUCT sfo = { 0 };
      sfo.hwnd = NULL;
      sfo.wFunc = FO_MOVE;
      sfo.pFrom = szTempFileName;
      sfo.pTo = szNewFileName;
      sfo.fFlags = FOF_SIMPLEPROGRESS;
      sfo.lpszProgressTitle = szCaption;
      ::SHFileOperation(&sfo);

      // Notify shell
      ::PathRemoveFileSpec(szNewFileName);
      ::SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_PATH, (LPCVOID)szNewFileName, 0);
   }

   void __stdcall OnSplit(void)
   {
      USES_CONVERSION;
      const DWORD CHUNKSIZE = 32L*1024L;
      
      _RemoveFolders(m_arrFiles);

      for( int i=0; i<m_arrFiles.GetSize(); i++ ) {

         TCHAR szSourceFile[MAX_PATH];
         _tcscpy( szSourceFile, W2CT(m_arrFiles[i]));

         CFile f;
         if( f.Open(szSourceFile)==FALSE ) RETURN_ERROR;

         if( ::GetFileSize(f, NULL) < CHUNKSIZE ) continue;

         short nPart = 1;
         short nChunk = 9999; // Make sure we start by creating a new file
        
         CFile out;
         BYTE buf[1024];
         DWORD dwRead;
         while( f.Read(buf, sizeof(buf), &dwRead), dwRead>0 ) {
            
            if( nChunk>=CHUNKSIZE/1024 ) {
               out.Close();
               TCHAR szFileName[MAX_PATH];
               ::wsprintf(szFileName, _T("%s.%03d"), szSourceFile, nPart);
               nPart++;
               if( out.Create(szFileName)==FALSE ) RETURN_ERROR;
               nChunk = 0;
            }

            out.Write(buf, dwRead);
            nChunk++;
         }
         out.Close();
         f.Close();

         // Notify shell
         ::PathRemoveFileSpec(szSourceFile);
         ::SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_PATH, (LPCVOID)szSourceFile, 0);
      }
   }

private:
   void _RemoveFolders(CSimpleArray<CComBSTR> &arr)
   {
      USES_CONVERSION;
      for( int i=0; i<arr.GetSize(); i++ ) {
         if( ::PathIsDirectory( W2CT(arr[i]) ) ) {
            arr.RemoveAt(i);
            i=-1;
         }
      }
   }

   void _SortFiles(CSimpleArray<CComBSTR> &arr)
   {
      int cnt = arr.GetSize();
      bool bSorting;
      do {
         bSorting = false;
         for( int i=0; i<cnt-1; i++ ) {
            CComBSTR s1( arr[i] );
            CComBSTR s2( arr[i+1] );
            if( s2 < (BSTR)s1 ) {
               arr.SetAtIndex(i, s2);
               arr.SetAtIndex(i+1, s1);
               bSorting = true;
            }
         }
      } while( bSorting );
   }

};


#endif // !defined(AFX_CONTEXTMENU_H__20010124_FB58_1C60_A0E5_0080AD509054__INCLUDED_)

