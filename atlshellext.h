#ifndef __ATLSHELLEXT_H__
#define __ATLSHELLEXT_H__

#pragma once

///////////////////////////////////////////////////////////////////
// Shell Extension wrappers
//
// Written by Bjarke Viksoe (bjarke@viksoe.dk)
// Copyright (c) 2001-2011 Bjarke Viksoe.
//
// This code may be used in compiled form in any way you desire. This
// source file may be redistributed by any means PROVIDING it is 
// not sold for profit without the authors written consent, and 
// providing that this notice and the authors name is included. 
//
// This file is provided "as is" with no expressed or implied warranty.
// The author accepts no liability if it causes any damage to you or your
// computer whatsoever. It's free, so don't hassle me about it.
//
// Beware of bugs.
//

#ifndef __ATLSHELLEXTBASE_H__
   #error atlshellext.h requires atlshellextbase.h to be included first
#endif

#include <atlwin.h>
#include <prsht.h>

#include "resource.h"


//////////////////////////////////////////////////////////////////////////////
// CPidl

class CPidl
{
public:
   LPITEMIDLIST m_pidl;

public:
   CPidl() : m_pidl(NULL)
   {
   }

   virtual ~CPidl()
   {
      Delete();
   }

   BOOL IsEmpty() const
   {
      return PidlIsEmpty(m_pidl);
   }

   operator LPITEMIDLIST() const
   {
      return m_pidl;
   }

   LPITEMIDLIST* operator&()
   {
      ATLASSERT(m_pidl==NULL);
      return &m_pidl;
   }

   void Attach(LPITEMIDLIST pSrc)
   {
      Delete();
      m_pidl = pSrc;
   }

   LPITEMIDLIST Detach()
   {
      LPITEMIDLIST pidl = m_pidl;
      m_pidl = NULL;
      return pidl;
   }

   LPITEMIDLIST GetData() const
   {
      return m_pidl;
   }

   DWORD GetByteSize() const
   { 
      return PidlGetByteSize(m_pidl);
   }

   UINT GetCount() const
   { 
      return PidlGetCount(m_pidl);
   }

   LPCITEMIDLIST GetNextItem() const
   { 
      return PidlGetNextItem(m_pidl);
   }

   LPCITEMIDLIST GetLastItem() const
   { 
      return PidlGetLastItem(m_pidl);
   }

   LPITEMIDLIST CopyFirstItem() const
   { 
      return PidlCopyFirstItem(m_pidl);
   }
  
   void Delete()
   { 
      PidlDelete(m_pidl);
      m_pidl = NULL;
   }

   LPITEMIDLIST Copy() const
   { 
      return PidlCopy(m_pidl);
   }

   void Copy(LPCITEMIDLIST pidlSource) 
   { 
      Delete();
      m_pidl = PidlCopy(pidlSource);
   }  

   void Construct(LPCITEMIDLIST pidlRoot, LPCITEMIDLIST pidlPath) 
   { 
      Copy(pidlRoot);
      Concatenate(pidlPath);
   }  

   void Construct(LPCITEMIDLIST pidlRoot, LPCITEMIDLIST pidlPath, LPCITEMIDLIST pidlChild) 
   { 
      Copy(pidlRoot);
      Concatenate(pidlPath);
      ConcatenateChild(pidlChild);
   }  

   void Concatenate(LPCITEMIDLIST pidl2)
   {
      if( m_pidl == NULL ) {
         m_pidl = PidlCopy(pidl2);
         return;
      }
      if( pidl2 == NULL ) return;
      DWORD cb1 = GetByteSize() - sizeof(USHORT);
      DWORD cb2 = PidlGetByteSize(pidl2);
      LPITEMIDLIST pidlNew = (LPITEMIDLIST) _Module.m_Allocator.Alloc(cb1 + cb2);
      if( pidlNew != NULL ) {
         ::CopyMemory(pidlNew, m_pidl, cb1);
         ::CopyMemory( ((LPBYTE)pidlNew) + cb1, pidl2, cb2 );
      }
      Attach(pidlNew);
   }

   void ConcatenateChild(LPCITEMIDLIST pidl2)
   {
      if( m_pidl == NULL ) {
         m_pidl = PidlCopy(pidl2);
         return;
      }
      if( pidl2 == NULL ) return;
      DWORD cb1 = GetByteSize() - sizeof(USHORT);
      DWORD cb2 = pidl2->mkid.cb;
      LPITEMIDLIST pidlNew = (LPITEMIDLIST) _Module.m_Allocator.Alloc(cb1 + cb2 + sizeof(USHORT));
      if( pidlNew != NULL ) {
         ::CopyMemory(pidlNew, m_pidl, cb1);
         ::CopyMemory( ((LPBYTE)pidlNew) + cb1, pidl2, cb2 );
         *(WORD*) (((LPBYTE)pidlNew) + cb1 + cb2) = 0;  
      }
      Attach(pidlNew);
   }

   void RemoveLast()
   {
      LPITEMIDLIST pidlLast = const_cast<LPITEMIDLIST>(PidlGetLastItem(m_pidl));
      if( pidlLast != NULL ) pidlLast->mkid.cb = 0;
   }

   DWORD GetHash() const
   {
      DWORD dwHash = 0;
      UINT nSize = GetByteSize();
      LPBYTE pBytes = (LPBYTE) m_pidl;
      for( UINT i = 0; i < nSize; i++ ) dwHash += pBytes[i];
      return dwHash;
   }

   HRESULT GetFromShellFilename(HWND hWnd, LPCOLESTR pstrFileName)
   {
      Delete();
      HRESULT Hr;
      LPSHELLFOLDER pDesktopFolder = NULL;
      Hr = ::SHGetDesktopFolder(&pDesktopFolder);
      if( FAILED(Hr) ) return Hr;
      Hr = pDesktopFolder->ParseDisplayName(hWnd, NULL, const_cast<LPOLESTR>(pstrFileName), NULL, &m_pidl, NULL);
      pDesktopFolder->Release();
      return Hr; 
   }

   HRESULT WriteToStream(IStream* pStream) const
   {
      DWORD dwBytesWritten = 0;
      DWORD dwSize = GetByteSize();
      if( FAILED( pStream->Write(&dwSize, sizeof(DWORD), &dwBytesWritten) ) ) return E_FAIL;
      if( FAILED( pStream->Write(GetData(), dwSize, &dwBytesWritten) ) ) return E_FAIL;
      return S_OK;
   }

   HRESULT ReadFromStream(IStream* pStream)
   {
      DWORD dwSize = 0;
      DWORD dwBytesRead = 0;
      if( pStream->Read(&dwSize, sizeof(DWORD), &dwBytesRead) != S_OK ) return E_FAIL;
      LPITEMIDLIST p = (LPITEMIDLIST) _alloca(dwSize + 1);   // +1 because of possible NULL pidl; alloc fails otherwise
      if( pStream->Read(p, dwSize, &dwBytesRead) != S_OK ) return E_FAIL;
      Copy(dwBytesRead == 0 ? NULL : p);
      return dwSize == dwBytesRead ? S_OK : E_FAIL;
   }

   inline static BOOL PidlIsEmpty(LPCITEMIDLIST pidl)
   {
      return pidl == NULL || pidl->mkid.cb == 0;
   }

   static LPCITEMIDLIST PidlGetLastItem(LPCITEMIDLIST pidl) 
   { 
      // Get the PIDL of the last item in the list 
      LPCITEMIDLIST pidlLast = NULL;  
      if( pidl != NULL ) {    
         while( pidl->mkid.cb > 0 ) { 
            pidlLast = pidl;
            pidl = PidlGetNextItem(pidl);       
         }      
      }  
      return pidlLast; 
   } 

   inline static LPITEMIDLIST PidlGetNextItem(LPCITEMIDLIST pidl) 
   { 
      return pidl == NULL ? NULL : (LPITEMIDLIST)(LPBYTE)(((LPBYTE)pidl) + pidl->mkid.cb);
   }  

   static UINT PidlGetCount(LPCITEMIDLIST pidlSource)
   {
      UINT cbTotal = 0; 
      if( pidlSource != NULL ) {    
         while( pidlSource->mkid.cb > 0 ) {
            cbTotal++;
            pidlSource = PidlGetNextItem(pidlSource); 
         }
      }
      return cbTotal;
   }

   static DWORD PidlGetByteSize(LPCITEMIDLIST pidlSource)
   { 
      DWORD cbTotal = 0; 
      if( pidlSource != NULL ) {
         while( pidlSource->mkid.cb > 0 ) {
            cbTotal += pidlSource->mkid.cb;
            pidlSource = PidlGetNextItem(pidlSource);
         }
         // Add the size of the NULL terminating ITEMIDLIST    
         cbTotal += sizeof(USHORT);
      }  
      return cbTotal; 
   }

   static void PidlDelete(LPITEMIDLIST pidlSource)
   {
      if( pidlSource == NULL ) return;
      _Module.m_Allocator.Free(pidlSource);
   }

   static LPITEMIDLIST PidlCopy(LPCITEMIDLIST pidlSource) 
   { 
      // Allocate the new pidl
      if( pidlSource == NULL ) return NULL;
      DWORD cbSource = PidlGetByteSize(pidlSource);
      LPITEMIDLIST pidlTarget = (LPITEMIDLIST) _Module.m_Allocator.Alloc(cbSource); 
      if( pidlTarget == NULL ) return NULL;
      // Copy the source to the target 
      ::CopyMemory(pidlTarget, pidlSource, cbSource);
      return pidlTarget; 
   }

   static LPITEMIDLIST PidlCopyFirstItem(LPCITEMIDLIST pidlSource) 
   { 
      // Allocate the new pidl 
      if( pidlSource == NULL ) return NULL;  
      DWORD cbSource = pidlSource->mkid.cb + sizeof(USHORT); 
      LPITEMIDLIST pidlTarget = (LPITEMIDLIST) _Module.m_Allocator.Alloc(cbSource); 
      if( pidlTarget == NULL ) return NULL;
      // Copy the source to the target 
      ::CopyMemory(pidlTarget, pidlSource, cbSource);
      *(WORD*) (((LPBYTE)pidlTarget) + pidlTarget->mkid.cb) = 0;  
      return pidlTarget;
   }
};


/////////////////////////////////////////////////////////////////////////////
// CPidlList

class CPidlList
{
public:
   LPITEMIDLIST* m_pidls;
   UINT m_nCount;

   CPidlList() : m_pidls(NULL), m_nCount(0)
   {
   }

   CPidlList(LPITEMIDLIST pidl) : m_pidls(NULL), m_nCount(0)
   {
      SetList(pidl);
   }

   CPidlList(LPCITEMIDLIST* pidl, UINT nCount) : m_pidls(NULL), m_nCount(0)
   {
      SetList(pidl, nCount);
   }

   CPidlList(HWND hwndList, DWORD dwListViewMask, UINT uSortStateFirst = 0) : m_pidls(NULL), m_nCount(0)
   {
      SetList(hwndList, dwListViewMask, uSortStateFirst);
   }
   
   virtual ~CPidlList()
   {
      Delete();
   }

   operator LPCITEMIDLIST *() const 
   { 
      return const_cast<LPCITEMIDLIST*>(m_pidls); 
   }

   UINT GetCount() const 
   { 
      return m_nCount; 
   }

   LPITEMIDLIST* Detach()
   {
      LPITEMIDLIST* pidls = m_pidls;
      m_pidls = NULL;
      m_nCount = 0;
      return pidls;
   }
   
   HRESULT Attach(LPITEMIDLIST* pidlSource, UINT nCount)
   {
      Delete();
      m_pidls = pidlSource;
      m_nCount = nCount;
      return S_OK;
   }

   HRESULT SetList(LPITEMIDLIST pidl)
   {
      Delete();
      UINT nCount = CPidl::PidlGetCount(pidl);
      LPITEMIDLIST* pidls = (LPITEMIDLIST*) _Module.m_Allocator.Alloc(nCount * sizeof(LPITEMIDLIST));   
      if( pidls == NULL ) return E_OUTOFMEMORY;
      for( UINT i = 0; i < nCount; i++ ) {
         pidls[i] = CPidl::PidlCopyFirstItem(pidl);
         pidl = CPidl::PidlGetNextItem(pidl);
      }
      m_pidls = pidls;
      m_nCount = nCount;
      return S_OK;
   }

   HRESULT SetList(HWND hwndList, DWORD dwListViewMask, UINT uSortStateFirst = 0)
   {
      ATLASSERT(::IsWindow(hwndList));

      if( !::IsWindow(hwndList) ) return E_UNEXPECTED;

      Delete();

      UINT nCount = ( dwListViewMask == LVNI_SELECTED ? ListView_GetSelectedCount(hwndList) : ListView_GetItemCount(hwndList) );
      if( nCount == 0 ) return S_OK;

      LPITEMIDLIST* pidls = (LPITEMIDLIST*) _Module.m_Allocator.Alloc(nCount * sizeof(LPITEMIDLIST));   
      if( pidls == NULL ) return E_OUTOFMEMORY;

      UINT iIndex = 0;
      UINT iTopIndex = 0;
      int nItem = -1;
      while( (nItem = ListView_GetNextItem(hwndList, nItem, dwListViewMask)) != -1 ) {
         LVITEM lvi = { 0 };
         lvi.mask = LVIF_PARAM | LVIF_STATE;
         lvi.iItem = nItem;
         lvi.stateMask = uSortStateFirst;
         if( !ListView_GetItem(hwndList, &lvi) ) break;
         // Copy the PIDL to the list
         pidls[iIndex] = CPidl::PidlCopy( reinterpret_cast<LPITEMIDLIST>(lvi.lParam) );
         // Put selected items at top?
         if( uSortStateFirst != 0 && (lvi.state & uSortStateFirst) != 0 ) {
            LPITEMIDLIST pTemp = pidls[iTopIndex];
            pidls[iTopIndex] = pidls[iIndex];
            pidls[iIndex] = pTemp;
            iTopIndex++;
         }
         iIndex++;
      }
      ATLASSERT(iIndex==nCount);

      m_pidls = pidls;
      m_nCount = iIndex;
      return S_OK;
   }

   HRESULT SetList(LPCITEMIDLIST* pidlSource, UINT nCount)
   {
      Delete();
      if( pidlSource == NULL || nCount == 0 ) return S_OK;
      m_pidls = (LPITEMIDLIST*) _Module.m_Allocator.Alloc(nCount * sizeof(LPITEMIDLIST));
      if( m_pidls == NULL ) return E_OUTOFMEMORY;
      // Copy the items
      m_nCount = nCount;
      for( UINT i = 0; i < nCount; i++ ) m_pidls[i] = CPidl::PidlCopy(pidlSource[i]);
      return S_OK;
   }

   void Delete()
   {
      if( m_pidls == NULL ) return;
      for( UINT i = 0; i < m_nCount; i++ ) _Module.m_Allocator.Free( (LPVOID)m_pidls[i] );
      _Module.m_Allocator.Free(m_pidls);
      m_pidls = NULL;
      m_nCount = 0;
   }

   HRESULT ReadFromStream(IStream* pStream)
   {
      UINT nCount = 0;
      DWORD dwBytesRead = 0;
      if( pStream->Read(&nCount, sizeof(UINT), &dwBytesRead) != S_OK ) return E_FAIL;
      LPITEMIDLIST* pList = (LPITEMIDLIST*) _Module.m_Allocator.Alloc(nCount * sizeof(LPITEMIDLIST));   
      if( pList == NULL && nCount != 0 ) return E_OUTOFMEMORY;
      UINT i = 0;
      while( i < nCount ) {
         CPidl pidl;
         if( pidl.ReadFromStream(pStream) != S_OK ) break;
         pList[i++] = pidl.Detach();
      }
      return Attach(pList, i);
   }

   HRESULT WriteToStream(IStream* pStream)
   {
      DWORD dwBytesWritten = 0;
      if( FAILED( pStream->Write(&m_nCount, sizeof(UINT), &dwBytesWritten) ) ) return E_FAIL;
      for( UINT i = 0; i < m_nCount; i++ ) {
         CPidl pidl;
         pidl.Copy(m_pidls[i]);
         if( FAILED( pidl.WriteToStream(pStream) ) ) return E_FAIL;
      }
      return S_OK;
   }

   HRESULT FilterOnSFGAOF(IShellFolder* pFolder, DWORD dwAttribMask)
   {
      ATLASSERT(pFolder);
      ATLASSERT(dwAttribMask!=0);

      if( pFolder == NULL ) return E_POINTER;
      if( m_nCount == 0 ) return S_OK;

      LPITEMIDLIST* pidls = (LPITEMIDLIST*) _Module.m_Allocator.Alloc(m_nCount * sizeof(LPITEMIDLIST));
      if( pidls == NULL ) return E_OUTOFMEMORY;

      UINT nCount = 0;
      for( UINT i = 0; i < m_nCount; i++ ) {
         DWORD dwAttribs = dwAttribMask;
         LPCITEMIDLIST pidl = m_pidls[i];
         pFolder->GetAttributesOf(1, &pidl, &dwAttribs);
         if( (dwAttribs & dwAttribMask) == dwAttribMask ) {
            pidls[nCount] = CPidl::PidlCopy( m_pidls[i] );
            nCount++;
         }
      }

      // We've recreated a new PIDL list.
      // The allocated memory block for the list may actually be too large, but
      // it doesn't matter for the PIDL functions.
      return Attach(pidls, nCount);
   }
};


#ifdef __ATLCOM_H__

/////////////////////////////////////////////////////////////////////////////
// CPidlEnum

class ATL_NO_VTABLE CPidlEnum : 
   public CComObjectRootEx<CComSingleThreadModel>,
   public IEnumIDList
{
public:
   CPidlEnum() : m_lCount(0), m_lPos(0), m_pCur(NULL)
   {
   }

DECLARE_NO_REGISTRY()

BEGIN_COM_MAP(CPidlEnum)
   COM_INTERFACE_ENTRY_IID(IID_IEnumIDList, IEnumIDList)
END_COM_MAP()

public:
   HRESULT Init(LPCITEMIDLIST pidl)
   {      
      m_pidl.Copy(pidl);
      m_lCount = m_pidl.GetCount();
      m_lPos = 0;
      m_pCur = m_pidl;
      return S_OK;
   }

   STDMETHOD(Next)(ULONG celt, LPITEMIDLIST* rgelt, ULONG* pceltFetched)
   { 
      ATLTRACE2(atlTraceCOM, 0, _T("IEnumIDList::Next\n"));
      *rgelt = NULL; 
      if( pceltFetched != NULL ) *pceltFetched = 0;
      if( pceltFetched == NULL && celt != 1 ) return E_POINTER;
      ULONG nCount = 0;
      while( m_lPos < m_lCount && nCount < celt ) {
         rgelt[nCount++] = CPidl::PidlCopyFirstItem(m_pCur);
         m_pCur = CPidl::PidlGetNextItem(m_pCur);
         m_lPos++;
      }
      if( pceltFetched != NULL ) *pceltFetched = nCount;
      return celt == nCount ? S_OK : S_FALSE;
   }   

   STDMETHOD(Reset)(void)
   { 
      ATLTRACE2(atlTraceCOM, 0, _T("IEnumIDList::Reset\n"));
      m_lPos = 0;
      m_pCur = m_pidl;
      return S_OK; 
   }

   STDMETHOD(Skip)(ULONG celt)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IEnumIDList::Skip\n"));
      ULONG nCount = 0;
      while( m_lPos < m_lCount && nCount < celt ) {
         m_pCur = CPidl::PidlGetNextItem(m_pCur);
         m_lPos++; nCount++;
      }
      return nCount == celt ? S_OK : S_FALSE;
   }

   STDMETHOD(Clone)(IEnumIDList** /*ppEnum*/)
   {
      ATLTRACENOTIMPL(_T("IEnumIDList::Clone"));
   }

public:
   CPidl m_pidl;
   LPCITEMIDLIST m_pCur;
   ULONG m_lCount; 
   ULONG m_lPos;
};


//////////////////////////////////////////////////////////////////////////////
// IShellFolderImpl

#define GET_SHGDN_FOR(dwFlags)         ((DWORD)dwFlags & (DWORD)0x0000FF00)
#define GET_SHGDN_RELATION(dwFlags)    ((DWORD)dwFlags & (DWORD)0x000000FF)


template< class T >
class ATL_NO_VTABLE IShellFolderImpl : 
   public IShellFolder2,
   public IPersistIDList,
   public IPersistFolder3
{
public:
   CPidl m_pidlRoot;
   CPidl m_pidlPath;

   // IPersistFolder

   STDMETHOD(GetClassID)(CLSID* pClassID)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IPersistFolder::GetClassID\n"));
      ATLASSERT(pClassID);
      if( pClassID == NULL ) return E_POINTER;
      *pClassID = T::GetObjectCLSID();
      return S_OK;
   }

   STDMETHOD(Initialize)(LPCITEMIDLIST pidlRoot)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IPersistFolder::Initialize\n"));
      ATLASSERT(pidlRoot);
      if( pidlRoot == NULL ) return E_INVALIDARG;
      m_pidlRoot.Copy(pidlRoot);
      return S_OK;
   }

   // IPersistFolder2

   STDMETHOD(GetCurFolder)(LPITEMIDLIST* ppidl)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IPersistFolder2::GetCurFolder\n"));
      ATLASSERT(ppidl);
      if( ppidl == NULL ) return E_INVALIDARG;
      CPidl pidlFolder;
      pidlFolder.Copy(m_pidlRoot);
      pidlFolder.Concatenate(m_pidlPath);
      *ppidl = pidlFolder.Detach();
      return S_OK;
   }

   // IPersistFolder3

   STDMETHOD(InitializeEx)(IBindCtx* /*pbc*/, LPCITEMIDLIST pidlRoot, const PERSIST_FOLDER_TARGET_INFO* /*ppfti*/)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IPersistFolder3::InitializeEx\n"));
      return Initialize(pidlRoot);
   }

   STDMETHOD(GetFolderTargetInfo)(PERSIST_FOLDER_TARGET_INFO* ppfti)
   {
      ::ZeroMemory(ppfti, sizeof(PERSIST_FOLDER_TARGET_INFO));
      ATLTRACENOTIMPL(_T("IPersistFolder3::GetFolderTargetInfo"));
   }

   // IPersistIDList

   STDMETHOD(SetIDList)(LPCITEMIDLIST /*pidl*/)
   {
      ATLTRACENOTIMPL(_T("IPersistIDList::SetIDList"));
   }

   STDMETHOD(GetIDList)(LPITEMIDLIST* ppidl)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IPersistIDList::GetIDList\n"));
      return GetCurFolder(ppidl);
   }

   // IShellFolder

   STDMETHOD(EnumObjects)(HWND, DWORD, LPENUMIDLIST* ppRetVal)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellFolder::EnumObjects\n"));
      ATLASSERT(ppRetVal);
      // Return empty collection
      // TODO: Override this and populate the pidl list.
      HRESULT Hr;
      CPidl pidlEmpty;
      CComObject<CPidlEnum>* pEnum = NULL;
      Hr = CComObject<CPidlEnum>::CreateInstance(&pEnum);
      if( FAILED(Hr) ) return Hr;
      Hr = pEnum->Init(pidlEmpty);
      if( FAILED(Hr) ) return Hr;
      return pEnum->QueryInterface(IID_IEnumIDList, (LPVOID*) ppRetVal);
   }

   STDMETHOD(BindToObject)(LPCITEMIDLIST /*pidl*/, LPBC, REFIID /*riid*/, LPVOID* ppRetVal)
   {
      // TODO: Implement to support sub-folders
      *ppRetVal = NULL;
      ATLTRACENOTIMPL(_T("IShellFolder::BindToObject"));
   }

   STDMETHOD(BindToStorage)(LPCITEMIDLIST pidl, LPBC pBindCtx, REFIID riid, LPVOID* ppRetVal)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellFolder::BindToStorage\n"));
      T* pT = static_cast<T*>(this); pT;
      return pT->BindToObject(pidl, pBindCtx, riid, ppRetVal);
   }

   STDMETHOD(CreateViewObject)(HWND /*hwndOwner*/, REFIID /*riid*/, LPVOID* /*ppRetVal*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolder::CreateViewObject"));
   }

   STDMETHOD(GetUIObjectOf)(HWND, UINT, LPCITEMIDLIST*, REFIID, LPUINT, LPVOID*)
   {
      ATLTRACENOTIMPL(_T("IShellFolder::GetUIObjectOf"));
   }

   STDMETHOD(GetDisplayNameOf)(LPCITEMIDLIST, DWORD, LPSTRRET)
   {
      ATLTRACENOTIMPL(_T("IShellFolder::GetDisplayNameOf"));
   }

   STDMETHOD(GetAttributesOf)(UINT, LPCITEMIDLIST*, LPDWORD rgfInOut)
   {
      *rgfInOut = 0;
      ATLTRACENOTIMPL(_T("IShellFolder::GetAttributesOf"));
   }

   STDMETHOD(ParseDisplayName)(HWND, LPBC, LPOLESTR, LPDWORD, LPITEMIDLIST* ppList, LPDWORD)
   {
      *ppList = NULL;
      ATLTRACENOTIMPL(_T("IShellFolder::ParseDisplayName"));
   }

   STDMETHOD(SetNameOf)(HWND, LPCITEMIDLIST, LPCOLESTR, DWORD, LPITEMIDLIST*)
   {
      ATLTRACENOTIMPL(_T("IShellFolder::SetNameOf"));
   }

   STDMETHOD(CompareIDs)(LPARAM lParam, LPCITEMIDLIST pidl1, LPCITEMIDLIST pidl2)
   {
      // The lParam is used to determine whether to compare items or sub-items.
      // We must compare the complete ITEMIDLIST.
      T* pT = static_cast<T*>(this); pT;
      while( pidl1 != NULL || pidl2 != NULL ) {
         short iCmd = pT->_CompareItems( (short) (lParam & 0xFF), pidl1, pidl2 );
         if( iCmd < 0 ) return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(-1));
         if( iCmd > 0 ) return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(1));
         pidl1 = CPidl::PidlGetNextItem(pidl1);
         pidl2 = CPidl::PidlGetNextItem(pidl2);
         if( CPidl::PidlIsEmpty(pidl1) ) pidl1 = NULL;
         if( CPidl::PidlIsEmpty(pidl2) ) pidl2 = NULL;
      }
      return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
   }

   // IShellFolder2

   STDMETHOD(GetDefaultSearchGUID)(GUID* /*pguid*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolder2::GetDefaultSearchGUID"));
   }

   STDMETHOD(EnumSearches)(IEnumExtraSearch** /*ppenum*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolder2::EnumSearches"));
   }

   STDMETHOD(GetDefaultColumn)(DWORD /*dwRes*/, ULONG* /*pSort*/, ULONG* /*pDisplay*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolder2::GetDefaultColumn"));
   }

   STDMETHOD(GetDefaultColumnState)(UINT /*iColumn*/, SHCOLSTATEF* /*pcsFlags*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolder2::GetDefaultColumnState"));
   }

   STDMETHOD(GetDetailsEx)(LPCITEMIDLIST /*pidl*/, const SHCOLUMNID* /*pscid*/, VARIANT* /*pv*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolder2::GetDetailsEx"));
   }

   STDMETHOD(GetDetailsOf)(LPCITEMIDLIST /*pidl*/, UINT /*iColumn*/, SHELLDETAILS* /*psd*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolder2::GetDetailsOf"));
   }

   STDMETHOD(MapColumnToSCID)(UINT /*iColumn*/, SHCOLUMNID* /*pscid*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolder2::MapColumnToSCID"));
   }

   // IShellDetails

   STDMETHOD(ColumnClick)(UINT /*iColumn*/)
   {
      ATLTRACENOTIMPL(_T("IShellDetails::ColumnClick"));
   }

   // Implementation

   short _CompareItems(short /*iColumn*/, LPCITEMIDLIST /*pData1*/, LPCITEMIDLIST /*pData2*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolder::_CompareItems"));
   }
};

#ifndef SID_SShellFolderView
   #define SID_SShellFolderView __uuidof(IShellFolderView)
#endif // SID_SShellFolderView


//////////////////////////////////////////////////////////////////////////////
// IShellViewImpl

typedef enum 
{
   TBI_STD = 0,
   TBI_VIEW,
   TBI_LOCAL,
   TBI_LAST
} TOOLBARITEM;

typedef struct
{
   TOOLBARITEM nType;
   TBBUTTON tb;
} NS_TOOLBUTTONINFO, *PNS_TOOLBUTTONINFO;

#ifndef _WIN32_WINNT_WIN7
   #define FVM_CONTENT  8
#endif // _WIN32_WINNT_WIN7


template< class T >
class ATL_NO_VTABLE IShellViewImpl : 
#if _WIN32_WINNT >= 0x0600
   public IShellView3
#else
   public IShellView2
#endif // _WIN32_WINNT
{
public:
   enum { IDC_LISTVIEW = 123 };

   BEGIN_MSG_MAP( IShellViewImpl<T> )
      MESSAGE_HANDLER(WM_CREATE, OnCreate)
      MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
      MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
      MESSAGE_HANDLER(WM_SIZE, OnSize)
      MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
      MESSAGE_HANDLER(WM_INITMENUPOPUP, OnInitMenu)
      MESSAGE_HANDLER(WM_ERASEBKGND, OnEraseBackground)
      NOTIFY_CODE_HANDLER(NM_SETFOCUS, OnNotifySetFocus)
      NOTIFY_CODE_HANDLER(NM_KILLFOCUS, OnNotifyKillFocus)
   END_MSG_MAP()

   IShellViewImpl() : m_uViewState(SVUIA_DEACTIVATE), m_hWnd(NULL), m_hwndList(NULL), m_hChangeNotify(NULL), m_dwProfferCookie(0UL)
   {
      HrOleInit = ::OleInitialize(NULL);
      ::ZeroMemory(&m_ShellFlags, sizeof(m_ShellFlags));
      ::ZeroMemory(&m_FolderSettings, sizeof(m_FolderSettings));
      m_iIconSize = ::GetSystemMetrics(SM_CXICON);
      m_FolderSettings.ViewMode = FVM_DETAILS,
      m_FolderSettings.fFlags = FWF_AUTOARRANGE | FWF_SHOWSELALWAYS;
      m_dwListViewStyle = WS_TABSTOP | WS_VISIBLE | WS_CHILD | LVS_REPORT | LVS_NOSORTHEADER | LVS_SHAREIMAGELISTS;
   }

   ~IShellViewImpl()
   {
      if( ::IsWindow(m_hWnd) ) ::DestroyWindow(m_hWnd);
      if( SUCCEEDED(HrOleInit) ) ::OleUninitialize();
   }

public:
   HWND m_hWnd;
   HWND m_hwndList;
   DWORD m_dwListViewStyle;
#ifdef _DEBUG
   const MSG* m_pCurrentMsg;
#endif // _DEBUG
   SHELLFLAGSTATE m_ShellFlags;
   FOLDERSETTINGS m_FolderSettings;
   CComPtr<IDropTarget> m_spDropTarget;
   CComQIPtr<IShellBrowser, &IID_IShellBrowser> m_spShellBrowser;
   CComQIPtr<ICommDlgBrowser, &IID_ICommDlgBrowser> m_spCommDlg;
   ULONG m_dwProfferCookie;
   ULONG m_hChangeNotify;
   HRESULT HrOleInit;
   UINT m_uViewState;
   int m_iIconSize;

public:
   // IOleWindow

   STDMETHOD(GetWindow)(HWND* phWnd)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IOleWindow::GetWindow\n"));
      ATLASSERT(phWnd);
      *phWnd = m_hWnd;
      return S_OK;
   }

   STDMETHOD(ContextSensitiveHelp)(BOOL)
   {
      ATLTRACENOTIMPL(_T("IOleWindow::ContextSesitiveHelp"));
   }

   // IShellView

   STDMETHOD(TranslateAccelerator)(LPMSG /*lpmsg*/)
   {
      return E_NOTIMPL;
   }

   STDMETHOD(Refresh)(void)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::Refresh\n"));
      ATLASSERT(::IsWindow(m_hwndList));
      // Refill the list
      T* pT = static_cast<T*>(this); pT;
      pT->_FillListView();
      return S_OK;
   }

   STDMETHOD(AddPropertySheetPages)(DWORD /*dwReserved*/, LPFNADDPROPSHEETPAGE /*lpfn*/, LPARAM /*lParam*/)
   {
      ATLTRACENOTIMPL(_T("IShellView::AddPropertySheetPages"));
   }

   STDMETHOD(SelectItem)(LPCITEMIDLIST pidlItem, UINT uFlags)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::SelectItem\n"));
      T* pT = static_cast<T*>(this); pT;
      return pT->SelectAndPositionItem(pidlItem, uFlags, NULL);
   }

   STDMETHOD(GetItemObject)(UINT /*uItem*/, REFIID /*riid*/, LPVOID* ppRetVal)
   {
      ATLASSERT(ppRetVal); 
      if( ppRetVal != NULL ) *ppRetVal = NULL;
      ATLTRACENOTIMPL(_T("IShellView::GetItemObject"));
   }

   STDMETHOD(EnableModeless)(BOOL /*fEnable*/)
   {
      ATLTRACENOTIMPL(_T("IShellView::EnableModeless"));
   }

   STDMETHOD(GetCurrentInfo)(LPFOLDERSETTINGS lpFS)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::GetCurrentInfo\n"));
      ATLASSERT(lpFS);
      *lpFS = m_FolderSettings;
      return S_OK;
   }

   STDMETHOD(UIActivate)(UINT uState)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::UIActivate (%d)\n"), uState);
      // Only do this if we are active
      if( m_uViewState == uState ) return S_OK;      
      // _ViewActivate() handles merging of menus etc
      T* pT = static_cast<T*>(this); pT;
      if( SVUIA_ACTIVATE_FOCUS == uState ) ::SetFocus(m_hwndList);
      pT->_ViewActivate(uState);
      if( uState != SVUIA_DEACTIVATE && m_spShellBrowser != NULL ) {
         // Update the status bar: set 'parts' and change text
         LRESULT lResult = 0;
         int nPartArray[1] = { -1 };
         m_spShellBrowser->SendControlMsg(FCW_STATUS, SB_SETPARTS, 1, (LPARAM) nPartArray, &lResult);
         // Set the statusbar text to the default description.
         // The string resource IDS_DESCRIPTION must be defined!
         TCHAR szName[128] = { 0 };
         ::LoadString(_Module.GetResourceInstance(), IDS_DESCRIPTION, szName, (sizeof(szName) / sizeof(szName[0])) - 1);
         m_spShellBrowser->SendControlMsg(FCW_STATUS, SB_SETTEXT, 0, (LPARAM) szName, &lResult);
      }
      return S_OK;
   }

   STDMETHOD(CreateViewWindow)(
      IShellView* /*lpPrevView*/,
      LPCFOLDERSETTINGS pFS, 
      IShellBrowser* pSB,
      RECT* prcView, 
      HWND* phWnd)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::CreateViewWindow\n"));
      ATLASSERT(prcView);
      ATLASSERT(pSB);
      ATLASSERT(pFS);
      ATLASSERT(phWnd);

      T* pT = static_cast<T*>(this); pT;

      *phWnd = NULL;

      // Register the ClassName.
      // The ClassName comes from the string resource IDS_CLASSNAME!
      TCHAR szClassName[64] = { 0 };
      ::LoadString(_Module.GetResourceInstance(), IDS_CLASSNAME, szClassName, (sizeof(szClassName) / sizeof(szClassName[0])) - 1);
      // If our window class has not been registered, then do so now...
      WNDCLASS wc = { 0 };
      if( ::GetClassInfo(_Module.GetModuleInstance(), szClassName, &wc) == FALSE ) {
         wc.style          = 0;
         wc.lpfnWndProc    = (WNDPROC)WndProc;
         wc.cbClsExtra     = 0;
         wc.cbWndExtra     = 0;
         wc.hInstance      = _Module.GetModuleInstance();
         wc.hIcon          = NULL;
         wc.hCursor        = ::LoadCursor(NULL, IDC_ARROW);
         wc.hbrBackground  = (HBRUSH)(COLOR_WINDOW + 1);
         wc.lpszMenuName   = NULL;
         wc.lpszClassName  = szClassName; 
         if( !::RegisterClass(&wc) ) return HRESULT_FROM_WIN32(::GetLastError());
      }

      // Set up the member variables
      m_spCommDlg = pSB;
      m_spShellBrowser = pSB;
      m_FolderSettings = *pFS;      
      m_ShellFlags.fWin95Classic = TRUE;
      m_ShellFlags.fShowAttribCol = TRUE;
      m_ShellFlags.fShowAllObjects = TRUE;

      // Get our parent window and create host window
      HWND hwndShell = NULL;
      m_spShellBrowser->GetWindow(&hwndShell);
      *phWnd = ::CreateWindowEx(WS_EX_CONTROLPARENT,
         szClassName,
         NULL,
         WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_TABSTOP,
         prcView->left, prcView->top,
         prcView->right - prcView->left, prcView->bottom - prcView->top,
         hwndShell,
         NULL,
         _Module.GetModuleInstance(),
         (LPVOID) pT);
      if( *phWnd == NULL ) return HRESULT_FROM_WIN32(::GetLastError());

      pT->_MergeToolbar(SVUIA_ACTIVATE_FOCUS);

      return S_OK;
   }

   STDMETHOD(DestroyViewWindow)(void)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::DestroyViewWindow\n"));

      // Make absolutely sure all our UI is cleaned up.
      UIActivate(SVUIA_DEACTIVATE);

      // Kill the window
      if( ::IsWindow(m_hWnd) ) ::DestroyWindow(m_hWnd);
      m_hwndList = NULL;
      m_hWnd = NULL;

      // Release the shell browser objects
      m_spShellBrowser.Release();
      m_spCommDlg.Release();

      return S_OK;
   }

   STDMETHOD(SaveViewState)(void)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::SaveViewState\n"));
      if( m_spShellBrowser == NULL ) return S_OK;
      CComPtr<IStream> spStream;
      m_spShellBrowser->GetViewStateStream(STGM_WRITE, &spStream);
      if( spStream == NULL ) return S_OK;
      DWORD dwWritten = 0;
      spStream->Write(&m_FolderSettings, sizeof(m_FolderSettings), &dwWritten);
      return S_OK;
   }

   // IShellView2

   STDMETHOD(CreateViewWindow2)(LPSV2CVW2_PARAMS lpParams) 
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::CreateViewWindow2\n"));
      // The pvid takes precedence over pfs->ViewMode 
      _ViewModeFromSVID(lpParams->pvid, (FOLDERVIEWMODE*) &lpParams->pfs->ViewMode);
      // Create the view...
      return CreateViewWindow(lpParams->psvPrev, lpParams->pfs, lpParams->psbOwner, lpParams->prcView, &lpParams->hwndView);
   }

   STDMETHOD(GetView)(SHELLVIEWID* pvid, ULONG uView) 
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::GetView (%ld)\n"), uView);
      if( pvid == NULL ) return E_POINTER;
      switch( uView ) { 
      case 0:                 return _SVIDFromViewMode(FVM_ICON, pvid); 
      case 1:                 return _SVIDFromViewMode(FVM_SMALLICON, pvid); 
      case 2:                 return _SVIDFromViewMode(FVM_LIST, pvid); 
      case 3:                 return _SVIDFromViewMode(FVM_DETAILS, pvid); 
      case 4:                 return _SVIDFromViewMode(FVM_TILE, pvid); 
      case SV2GV_CURRENTVIEW: return _SVIDFromViewMode((FOLDERVIEWMODE)m_FolderSettings.ViewMode, pvid); 
      case SV2GV_DEFAULTVIEW: return _SVIDFromViewMode(FVM_ICON, pvid); 
      } 
      return E_NOTIMPL;
   }

   STDMETHOD(HandleRename)(LPCITEMIDLIST pidl) 
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::HandleRename\n"));
      T* pT = static_cast<T*>(this); pT;
      POINT ptDummy = { 0 };
      return pT->SelectAndPositionItem(pidl, SVSI_EDIT, &ptDummy);
   }

   STDMETHOD(SelectAndPositionItem)(LPCITEMIDLIST /*pidlItem*/, UINT /*uFlags*/, POINT* /*point*/) 
   {
      ATLTRACENOTIMPL(_T("IShellView2::SelectAndPositionItem"));
   }

#ifdef __IShellView3_INTERFACE_DEFINED__
   // IShellView3

   STDMETHOD(CreateViewWindow3)(
      IShellBrowser* psbOwner, 
      IShellView* psvPrev, 
      SV3CVW3_FLAGS dwViewFlags,
      FOLDERFLAGS dwMask,
      FOLDERFLAGS dwFlags,
      FOLDERVIEWMODE fvMode,
      const SHELLVIEWID* /*pvid*/,
      const RECT *prcView,
      HWND* phwndView) 
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::CreateViewWindow3\n"));
      ATLASSERT(psbOwner);
      FOLDERSETTINGS fs = { 0 };
      CComPtr<IStream> spStream;
      psbOwner->GetViewStateStream(STGM_READ, &spStream);
      DWORD dwRead = 0;
      if( spStream != NULL ) spStream->Read(&fs, sizeof(fs), &dwRead);
      if( (dwViewFlags & SV3CVW3_FORCEVIEWMODE) != 0 ) fs.ViewMode = fvMode;
      if( (dwViewFlags & SV3CVW3_FORCEFOLDERFLAGS) != 0 ) fs.fFlags = (dwFlags & dwMask);
      RECT rcView = *prcView;
      // Create the view...
      return CreateViewWindow(psvPrev, &fs, psbOwner, &rcView, phwndView);
   }
#endif // __IShellView3_INTERFACE_DEFINED__

   // View handlers

   LRESULT _ViewActivate(UINT uState)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::_ViewActivate %d\n"), uState);
      // Don't do anything if the state isn't really changing
      if( m_uViewState == uState ) return S_OK;
      if( !::IsWindow(m_hwndList) ) return S_OK;
      T* pT = static_cast<T*>(this); pT;
      // Deactivate old view/menus first
      pT->_ViewDeactivate();
      // Only do this if we are now active...
      if( uState != SVUIA_DEACTIVATE ) {
         pT->_MergeMenus(uState);
         pT->_UpdateToolbar();
      }
      m_uViewState = uState;
      return 0;
   }

   LRESULT _ViewDeactivate(void)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellView::_ViewDeactivate\n"));
      if( !::IsWindow(m_hwndList) ) return S_OK;
      T* pT = static_cast<T*>(this);
      pT->_MergeMenus(SVUIA_DEACTIVATE);
      pT->_MergeToolbar(SVUIA_DEACTIVATE);
      m_uViewState = SVUIA_DEACTIVATE;
      return 0;
   }

   // Since ::SHGetSettings() is not implemented in all versions of the shell, get the 
   // function address manually at run time. This allows the extension to run on all 
   // platforms.
   void _GetShellSettings(SHELLFLAGSTATE& sfs, DWORD dwMask)
   {
      ATLASSERT(::GetModuleHandle(_T("SHELL32.DLL"))!=NULL);
      typedef void (WINAPI *PFNSHGETSETTINGSPROC)(LPSHELLFLAGSTATE,DWORD);
      static PFNSHGETSETTINGSPROC fnSHGetSettings = (PFNSHGETSETTINGSPROC) ::GetProcAddress(::GetModuleHandle(_T("SHELL32.DLL")), "SHGetSettings");
      if( fnSHGetSettings != NULL ) fnSHGetSettings(&sfs, dwMask);
   }

   // Windows Vista assigns its own theme to its Shell controls (ie. ListView). We'll allow 
   // it too.
   void _SetShellControlTheme(HWND hWnd)
   {
      typedef void (WINAPI *PFNSETWINDOWTHEME)(HWND,LPCWSTR,LPVOID);
      static HMODULE hInstUtx = ::LoadLibrary(_T("UxTheme.dll"));
      if( hInstUtx == NULL ) return;
      static PFNSETWINDOWTHEME fnSetWindowTheme = (PFNSETWINDOWTHEME) ::GetProcAddress(hInstUtx, "SetWindowTheme");
      if( fnSetWindowTheme != NULL ) fnSetWindowTheme(hWnd, L"explorer", NULL);
   }

   // Create an IDropTarget (through the IShellFolder::CreateViewObject) and register it as drag'n'drop
   // for the list window.
   HRESULT _RegisterDropTarget()
   {
      ATLASSERT(::IsWindow(m_hwndList));
      m_spDropTarget.Release();
      T* pT = static_cast<T*>(this); pT;
      if( FAILED( pT->GetItemObject(SVGIO_BACKGROUND, IID_IDropTarget, (LPVOID*) &m_spDropTarget) ) ) return E_FAIL;
      return ::RegisterDragDrop(m_hwndList, m_spDropTarget);
   }

   void _RevokeDropTarget()
   {
      if( ::IsWindow(m_hwndList) ) ::RevokeDragDrop(m_hwndList);
      m_spDropTarget.Release();
   }

   // Register for view changes.
   // This API was previously undocumented and thus only exported by ordinal. Since then,
   // Microsoft was forced to document it (reluctantly) in a DoJ anti-trust settlement.
   HRESULT _RegisterChangeNotify(LPCITEMIDLIST pidl, UINT uMsg, LONG fEvents = SHCNE_UPDATEDIR)
   {
#define SHCNF_ACCEPT_INTERRUPTS     0x0001 
#define SHCNF_ACCEPT_NON_INTERRUPTS 0x0002 
      typedef ULONG (WINAPI *PFNSHCHANGENOTIFYREGISTER)(HWND,int,LONG,UINT,int,SHChangeNotifyEntry*);
      PFNSHCHANGENOTIFYREGISTER fnSHChangeNotifyRegister = NULL; 
      if( fnSHChangeNotifyRegister == NULL ) 
         fnSHChangeNotifyRegister = (PFNSHCHANGENOTIFYREGISTER) ::GetProcAddress(::GetModuleHandle(_T("SHELL32.DLL")), "SHChangeNotifyRegister");
      if( fnSHChangeNotifyRegister == NULL ) 
         fnSHChangeNotifyRegister = (PFNSHCHANGENOTIFYREGISTER) ::GetProcAddress(::GetModuleHandle(_T("SHELL32.DLL")), MAKEINTRESOURCEA(2)); 
      if( fnSHChangeNotifyRegister == NULL ) return E_NOTIMPL;
      SHChangeNotifyEntry Nr = { pidl, TRUE };
      m_hChangeNotify = fnSHChangeNotifyRegister(m_hWnd, 
         SHCNF_ACCEPT_INTERRUPTS | SHCNF_ACCEPT_NON_INTERRUPTS,
         fEvents,
         uMsg,
         1, &Nr);
      return m_hChangeNotify != NULL ? S_OK : E_FAIL;
   }

   void _RevokeChangeNotify()
   {
      if( m_hChangeNotify == NULL ) return;
      typedef BOOL (WINAPI *PFNSHCHANGENOTIFYDEREGISTER)(ULONG);
      PFNSHCHANGENOTIFYDEREGISTER fnSHChangeNotifyDeregister = NULL; 
      if( fnSHChangeNotifyDeregister == NULL ) 
         fnSHChangeNotifyDeregister = (PFNSHCHANGENOTIFYDEREGISTER) ::GetProcAddress(::GetModuleHandle(_T("SHELL32.DLL")), "SHChangeNotifyDeregister");
      if( fnSHChangeNotifyDeregister == NULL ) 
         fnSHChangeNotifyDeregister = (PFNSHCHANGENOTIFYDEREGISTER) ::GetProcAddress(::GetModuleHandle(_T("SHELL32.DLL")), MAKEINTRESOURCEA(4)); 
      if( fnSHChangeNotifyDeregister == NULL ) return;
      fnSHChangeNotifyDeregister(m_hChangeNotify);
      m_hChangeNotify = NULL;
   }

   HRESULT _RegisterProffer(REFGUID guidService)
   {
      m_dwProfferCookie = 0;
      CComQIPtr<IProfferService> spProffer = m_spShellBrowser;
      CComQIPtr<IServiceProvider> spProvider = this;
      if( spProffer == NULL || spProvider == NULL ) return E_NOINTERFACE;
      return spProffer->ProfferService(guidService, spProvider, &m_dwProfferCookie);
   }
   
   void _RevokeProffer()
   {
      if( m_dwProfferCookie == 0 ) return;
      CComQIPtr<IProfferService> spProffer = m_spShellBrowser;
      if( spProffer == NULL ) return;
      spProffer->RevokeService(m_dwProfferCookie);
      m_dwProfferCookie = 0;
   }

   // A helper function which will take care of some of
   // the fancy new Win98 settings...
   void _UpdateShellSettings(void)
   {
      // Get the m_ShellFlags state
      _GetShellSettings(m_ShellFlags, 
         SSF_DESKTOPHTML | 
         SSF_NOCONFIRMRECYCLE | 
         SSF_SHOWALLOBJECTS | 
         SSF_SHOWATTRIBCOL | 
         SSF_DOUBLECLICKINWEBVIEW | 
         SSF_SHOWCOMPCOLOR |
         SSF_WIN95CLASSIC);

#ifndef LVS_EX_DOUBLEBUFFER
      const DWORD LVS_EX_DOUBLEBUFFER = 0x00010000;
#endif // LVS_EX_DOUBLEBUFFER
#ifndef LVS_EX_HEADERINALLVIEWS
      const DWORD LVS_EX_HEADERINALLVIEWS = 0x02000000;
#endif // LVS_EX_HEADERINALLVIEWS
#ifndef FWF_FULLROWSELECT
      const DWORD FWF_FULLROWSELECT =  0x200000;
#endif // FWF_FULLROWSELECT

      // Update the ListView control accordingly
      DWORD dwExStyles = LVS_EX_HEADERDRAGDROP |      // Allow but no auto-persist
                         LVS_EX_DOUBLEBUFFER |        // Causes blue marquee on WinXP
                         LVS_EX_HEADERINALLVIEWS;     // Causes headers to always display on Vista
      if( !m_ShellFlags.fWin95Classic && !m_ShellFlags.fDoubleClickInWebView ) {
         dwExStyles |= LVS_EX_ONECLICKACTIVATE | LVS_EX_TRACKSELECT | LVS_EX_UNDERLINEHOT;
      }
      if( _Module.m_dwWinVer >= MAKEWINVER(6,0) || (FWF_FULLROWSELECT & m_FolderSettings.fFlags) != 0 ) {
         dwExStyles |= LVS_EX_FULLROWSELECT;
      }
      if( (FWF_SINGLECLICKACTIVATE & m_FolderSettings.fFlags) != 0 ) {
         dwExStyles |= LVS_EX_ONECLICKACTIVATE;
      }
      ListView_SetExtendedListViewStyle(m_hwndList, dwExStyles);
   }

   LRESULT DefWindowProc(UINT uMsg, WPARAM wParam, LPARAM lParam)
   {
#ifdef STRICT
      return ::DefWindowProc(m_hWnd, uMsg, wParam, lParam);
#else
      return ::DefWindowProc(m_hWnd, uMsg, wParam, lParam);
#endif // STRICT
   }

   static LRESULT CALLBACK WndProc(HWND hWnd, UINT uMessage, WPARAM wParam, LPARAM lParam)
   {
      T* pT = reinterpret_cast<T*>(::GetWindowLongPtr(hWnd, GWLP_USERDATA)); pT;
      if( uMessage == WM_NCCREATE ) {
         LPCREATESTRUCT lpcs = (LPCREATESTRUCT) lParam;
         pT = (T*) lpcs->lpCreateParams;
         ::SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR) pT);
         // Set the window handle
         pT->m_hWnd = hWnd;
         return 1;
      }
      ATLASSERT(pT);
#ifdef _DEBUG
      MSG msg = { pT->m_hWnd, uMessage, wParam, lParam, 0, { 0, 0 } };
      const MSG* pOldMsg = pT->m_pCurrentMsg;
      pT->m_pCurrentMsg = &msg;
#endif // _DEBUG
      // pass to the message map to process
      LRESULT lRes = 0;
      BOOL bRet = pT->ProcessWindowMessage(pT->m_hWnd, uMessage, wParam, lParam, lRes, 0);
      // Restore saved value for the current message
#ifdef _DEBUG
      ATLASSERT(pT->m_pCurrentMsg==&msg);
      pT->m_pCurrentMsg = pOldMsg;
#endif // _DEBUG
      if( !bRet ) lRes = pT->DefWindowProc(uMessage, wParam, lParam);
      return lRes;
   }

   // Message handlers

   LRESULT OnSetFocus(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
   {
      ATLTRACE2(atlTraceWindowing, 0, _T("IShellView::OnSetFocus\n"));
      if( ::IsWindow(m_hwndList) ) ::SetFocus(m_hwndList);
      return 0;
   }

   LRESULT OnKillFocus(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
   {
      ATLTRACE2(atlTraceWindowing, 0, _T("IShellView::OnKillFocus\n"));
      return 0;
   }

   LRESULT OnSettingChange(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
   {
      T* pT = static_cast<T*>(this); pT;
      pT->_UpdateShellSettings();
      return 0;
   }

   LRESULT OnNotifySetFocus(UINT /*CtlID*/, LPNMHDR /*lpnmh*/, BOOL &/*bHandled*/)
   {
      ATLTRACE2(atlTraceWindowing, 0, _T("IShellView::OnNotifySetFocus\n"));
      if( !::IsWindow(m_hwndList) ) return 0;
      if( m_uViewState == SVUIA_DEACTIVATE ) return 0;
      // Tell the browser one of our windows has received the focus. This should always 
      // be done before merging menus (_ViewActivate() merges the menus) if one of our 
      // windows has the focus.
      T* pT = static_cast<T*>(this); pT;
      if( m_spShellBrowser != NULL ) m_spShellBrowser->OnViewWindowActive(pT);
      pT->_ViewActivate(SVUIA_ACTIVATE_FOCUS);
      if( m_spCommDlg != NULL ) m_spCommDlg->OnStateChange(this, CDBOSC_SETFOCUS);
      return 0;
   }

   LRESULT OnNotifyKillFocus(UINT /*CtlID*/, LPNMHDR /*lpnmh*/, BOOL &/*bHandled*/)
   {
      ATLTRACE2(atlTraceWindowing, 0, _T("IShellView::OnNotifyKillFocus\n"));
      if( !::IsWindow(m_hwndList) ) return 0;
      if( m_uViewState == SVUIA_DEACTIVATE ) return 0;
      T* pT = static_cast<T*>(this); pT;
      pT->_ViewActivate(SVUIA_ACTIVATE_NOFOCUS);
      if( m_spCommDlg != NULL ) m_spCommDlg->OnStateChange(this, CDBOSC_KILLFOCUS);
      return 0;
   }

   LRESULT OnInitMenu(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/)
   {
      T* pT = static_cast<T*>(this);
      pT->_UpdateMenu((HMENU)wParam, NULL);
      return 0;
   }

   LRESULT OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*bHandled*/)
   {
      // Resize the ListView to fit our window
      if( ::IsWindow(m_hwndList) ) ::MoveWindow(m_hwndList, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
      return 0;
   }

   LRESULT OnEraseBackground(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
   {
      return 1; // avoid flicker
   }

   LRESULT OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
   {
      // Create the ListView
      T* pT = static_cast<T*>(this); pT;
      if( !pT->_CreateListView() ) return 0;
      if( !pT->_InitListView() ) return 0;
      if( !pT->_FillListView() ) return 0;
      return 0;
   }

   // Operations

   BOOL _AppendToolbarItems(
      const PNS_TOOLBUTTONINFO pButtons, 
      int nCount, 
      LPARAM lOffsetFile, 
      LPARAM lOffsetView, 
      LPARAM lOffsetCustom)
   {
      ATLASSERT(nCount>0);
      if( m_spShellBrowser == NULL ) return TRUE;
      LPTBBUTTON ptbb = (LPTBBUTTON) ::GlobalAlloc(GPTR, sizeof(TBBUTTON) * nCount);
      if( ptbb == NULL ) return FALSE;
      int i;
      for( i = 0; i < nCount; i++ ) {
         switch( pButtons[i].nType ) {
         case TBI_STD:   ptbb[i].iBitmap = (int) lOffsetFile + pButtons[i].tb.iBitmap; break;
         case TBI_VIEW:  ptbb[i].iBitmap = (int) lOffsetView + pButtons[i].tb.iBitmap; break;
         case TBI_LOCAL: ptbb[i].iBitmap = (int) lOffsetCustom + pButtons[i].tb.iBitmap; break;
         }
         if( pButtons[i].nType == TBI_LAST ) break;
         ptbb[i].idCommand = pButtons[i].tb.idCommand;
         ptbb[i].fsState = pButtons[i].tb.fsState;
         ptbb[i].fsStyle = pButtons[i].tb.fsStyle;
         ptbb[i].dwData = pButtons[i].tb.dwData;
         ptbb[i].iString = pButtons[i].tb.iString;
      }
      m_spShellBrowser->SetToolbarItems(ptbb, i, FCT_MERGE);
      ::GlobalFree((HGLOBAL)ptbb);
      return TRUE;
   }

   UINT _GetMenuPosFromID(HMENU hMenu, UINT ID) const
   {
      UINT nCount = ::GetMenuItemCount(hMenu);
      for( UINT i = 0; i < nCount; i++ ) {
         if( ::GetMenuItemID(hMenu, i) == ID ) return i;
      }
      return (UINT) -1;
   }

   BOOL _AppendMenu(HMENU hMenu, HMENU hMenuSource, UINT nPosition)
   {
      ATLASSERT(::IsMenu(hMenu));
      ATLASSERT(::IsMenu(hMenuSource));
      // Get the HMENU of the popup
      if( hMenu == NULL ) return FALSE;
      if( hMenuSource == NULL ) return FALSE;
      // Make sure that we start with only one separator menu-item
      int iStartPos = 0;
      if( (::GetMenuState(hMenuSource, 0, MF_BYPOSITION) & MF_SEPARATOR) != 0 ) {
         if( (nPosition == 0) || (::GetMenuState(hMenu, nPosition - 1, MF_BYPOSITION) & MF_SEPARATOR) != 0 ) {
            iStartPos++;
         }
      }
      // Go...
      int nMenuItems = ::GetMenuItemCount(hMenuSource);
      for( int i = iStartPos; i < nMenuItems; i++ ) {
         // Get state information
         UINT state = ::GetMenuState(hMenuSource, i, MF_BYPOSITION);
         TCHAR szItemText[256] = { 0 };
         int nLen = ::GetMenuString(hMenuSource, i, szItemText, (sizeof(szItemText) / sizeof(szItemText[0])) - 1, MF_BYPOSITION);
         // Is this a separator?
         if( (state & MF_SEPARATOR) != 0 ) {
            ::InsertMenu(hMenu, nPosition++, state | MF_STRING | MF_BYPOSITION, 0, _T(""));
         }
         else if( (state & MF_POPUP) != 0 ) {
            // Strip the HIBYTE because it contains a count of items
            state = LOBYTE(state) | MF_POPUP;
            // Then create the new submenu by using recursive call
            HMENU hSubMenu = ::CreateMenu();
            _AppendMenu(hSubMenu, ::GetSubMenu(hMenuSource, i), 0);
            ATLASSERT(::GetMenuItemCount(hSubMenu)>0);
            // Non-empty popup -- add it to the shared menu bar
            ::InsertMenu(hMenu, nPosition++, state | MF_BYPOSITION, (UINT_PTR) hSubMenu, szItemText);
         }
         else if( nLen > 0 ) {
            // Only non-empty items should be added
            ATLASSERT(szItemText[0] != _T('\0'));
            ATLASSERT(::GetMenuItemID(hMenuSource, i)>FCIDM_SHVIEWFIRST && ::GetMenuItemID(hMenuSource, i)<FCIDM_SHVIEWLAST);
            // Here the state does not contain a count in the HIBYTE
            ::InsertMenu(hMenu, nPosition++, state | MF_BYPOSITION, ::GetMenuItemID(hMenuSource, i), szItemText);
         }
      }
      return TRUE;
   }

   HRESULT _ViewModeFromSVID(const SHELLVIEWID* pvid, FOLDERVIEWMODE* pViewMode) const
   {
      if( pViewMode != NULL ) *pViewMode = FVM_ICON;
      if( pvid == NULL || pViewMode == NULL ) return E_INVALIDARG;
      if( *pvid == VID_LargeIcons )      *pViewMode = FVM_ICON; 
      else if( *pvid == VID_SmallIcons ) *pViewMode = FVM_SMALLICON; 
      else if( *pvid == VID_Thumbnails ) *pViewMode = FVM_THUMBNAIL; 
      else if( *pvid == VID_ThumbStrip ) *pViewMode = FVM_THUMBSTRIP; 
      else if( *pvid == VID_List )       *pViewMode = FVM_LIST; 
      else if( *pvid == VID_Tile )       *pViewMode = FVM_TILE; 
      else if( *pvid == VID_Details )    *pViewMode = FVM_DETAILS; 
      else return E_FAIL; 
      return S_OK; 
   }

   HRESULT _SVIDFromViewMode(FOLDERVIEWMODE mode, SHELLVIEWID* svid) const
   { 
      ATLASSERT(svid);
      switch( mode ) { 
      case FVM_SMALLICON:  *svid = VID_SmallIcons; break;
      case FVM_LIST:       *svid = VID_List;       break;
      case FVM_DETAILS:    *svid = VID_Details;    break;
      case FVM_THUMBNAIL:  *svid = VID_Thumbnails; break;
      case FVM_TILE:       *svid = VID_Tile;       break;
      case FVM_THUMBSTRIP: *svid = VID_ThumbStrip; break;
      case FVM_ICON:       *svid = VID_LargeIcons; break;
      default:             *svid = VID_LargeIcons; break;
      }
      return S_OK;
   } 

   BOOL _IsExplorerMode() const
   {
      if( m_spShellBrowser == NULL ) return FALSE;
      // MSDN actually documents that we can determine if we're in explorer mode
      // by asking for the tree control.
      HWND hwndTree = NULL;
      return SUCCEEDED(m_spShellBrowser->GetControlWindow(FCW_TREE, &hwndTree)) && hwndTree != NULL;
   }

   BOOL _IsCommDlgMode() const
   {
      return m_spCommDlg != NULL;
   }

   // Overridables

   BOOL _CreateListView(void)
   {
      ATLASSERT((m_dwListViewStyle & (WS_VISIBLE|WS_CHILD))==(WS_VISIBLE|WS_CHILD));

      // Initialize and create the actual List View control
      m_dwListViewStyle &= ~LVS_TYPEMASK;
      m_dwListViewStyle |= LVS_ICON;
      if( (FWF_ALIGNLEFT & m_FolderSettings.fFlags) != 0 ) m_dwListViewStyle |= LVS_ALIGNLEFT;
      if( (FWF_AUTOARRANGE & m_FolderSettings.fFlags) != 0 ) m_dwListViewStyle |= LVS_AUTOARRANGE;
#if (_WIN32_IE >= 0x0500)
      if( (FWF_SHOWSELALWAYS & m_FolderSettings.fFlags) != 0 ) m_dwListViewStyle |= LVS_SHOWSELALWAYS;
#endif // _WIN32_IE
      if( _Module.m_dwWinVer >= MAKEWINVER(6,0) ) m_FolderSettings.fFlags |= FWF_NOCLIENTEDGE;

      // Go on, create the ListView control
      m_hwndList = ::CreateWindowEx( (FWF_NOCLIENTEDGE & m_FolderSettings.fFlags) != 0 ? 0 : WS_EX_CLIENTEDGE,
         WC_LISTVIEW,
         NULL,
         m_dwListViewStyle,
         0,0,0,0,
         m_hWnd,
         (HMENU) IDC_LISTVIEW,
         _Module.GetModuleInstance(),
         NULL);
      if( m_hwndList == NULL ) return FALSE;

      T* pT = static_cast<T*>(this); pT;
      pT->_SetShellControlTheme(m_hwndList);
      pT->_UpdateShellSettings();

      return TRUE;
   }

   BOOL _InitListView(void) { return TRUE; };
   BOOL _FillListView(void) { return TRUE; };
   BOOL _MergeToolbar(UINT /*uState*/) { return TRUE; };
   BOOL _MergeMenus(UINT /*uState*/) { return TRUE; };
   BOOL _UpdateToolbar() { return TRUE; };
   BOOL _UpdateMenu(HMENU /*hMenu*/, LPCITEMIDLIST /*pidl*/) { return TRUE; };
};


//////////////////////////////////////////////////////////////////////////////
// IShellFolderViewCBImpl

template< class T >
class ATL_NO_VTABLE IShellFolderViewCBImpl : public IShellFolderViewCB
{
public:
   STDMETHOD(MessageSFVCB)(UINT uMsg, WPARAM wParam, LPARAM lParam)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellFolderViewCB::MessageSFVCB (%lu,%ld,%ld)\n"), uMsg, wParam, lParam);
      LONG lResult = 0;
      T* pT = static_cast<T*>(this); pT;
      BOOL bResult = pT->ProcessWindowMessage(NULL, uMsg, wParam, lParam, lResult, 0);
      return bResult ? lResult : E_NOTIMPL;
   }
};


//////////////////////////////////////////////////////////////////////////////
// IFolderView for ListView control

template< class T >
class ATL_NO_VTABLE IFolderView2Impl : public IFolderView2
{
public:
   // IFolderView

   STDMETHOD(GetCurrentViewMode)(UINT* pViewMode)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::GetCurrentViewMode\n"));
      T* pT = static_cast<T*>(this); pT;
      *pViewMode = pT->m_FolderSettings.ViewMode;
      return S_OK;
   }

   STDMETHOD(SetCurrentViewMode)(UINT ViewMode)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::SetCurrentViewMode\n"));
      T* pT = static_cast<T*>(this); pT;
      pT->m_FolderSettings.ViewMode = ViewMode;
      if( !pT->_InitListView() ) return E_FAIL;
      if( !pT->_FillListView() ) return E_FAIL;
      return S_OK;
   }

   STDMETHOD(GetFolder)(REFIID riid, LPVOID* ppv)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::GetFolder\n"));
      T* pT = static_cast<T*>(this); pT;
      return pT->m_pFolder->QueryInterface(riid, ppv);
   }

   STDMETHOD(Item)(int iItemIndex, LPITEMIDLIST* ppidl)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::Item\n"));
      T* pT = static_cast<T*>(this); pT;
      LPCITEMIDLIST pidl = _GetPidlFromListIndex(iItemIndex);
      if( pidl == NULL ) return E_FAIL;
      *ppidl = CPidl::PidlCopy(pidl);
      return S_OK;
   }

   STDMETHOD(ItemCount)(UINT uFlags, int* pcItems)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::ItemCount\n"));
      T* pT = static_cast<T*>(this); pT;
      *pcItems = 0;
      int nCount = 0;
      if( (uFlags & SVGIO_TYPE_MASK) == SVGIO_CHECKED ) return S_OK;
      UINT uState = ((uFlags & SVGIO_TYPE_MASK) == SVGIO_SELECTION) ? LVNI_SELECTED : LVNI_ALL;
      for( int iItem = -1; (iItem = ListView_GetNextItem(pT->m_hwndList, iItem, uState)) != -1; ) nCount++;
      *pcItems = nCount;
      return S_OK;
   }

   STDMETHOD(Items)(UINT uFlags, REFIID riid, LPVOID* ppv)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::Items\n"));
      T* pT = static_cast<T*>(this); pT;
      if( (uFlags & SVGIO_SELECTION) != 0 && ListView_GetSelectedCount(pT->m_hwndList) == 0 ) return pT->m_pFolder->CreateViewObject(pT->m_hWnd, riid, ppv);
      return pT->GetItemObject(uFlags, riid, ppv);
   }

   STDMETHOD(GetSelectionMarkedItem)(int* /*piItem*/)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::GetSelectionMarkedItem"));
   }

   STDMETHOD(GetFocusedItem)(int* piItem)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView::GetFocusedItem\n"));
      T* pT = static_cast<T*>(this); pT;
      *piItem = ListView_GetNextItem(pT->m_hwndList, -1, LVNI_FOCUSED);
      return *piItem >= 0 ? S_OK : E_FAIL;
   }

   STDMETHOD(GetItemPosition)(LPCITEMIDLIST pidl, POINT* ppt)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::GetItemPosition"));
      T* pT = static_cast<T*>(this); pT;
      int iItem = _GetListIndexFromPidl(pidl);
      if( iItem == -1 ) return E_FAIL;
      ListView_GetItemPosition(pT->m_hwndList, iItem, ppt);
      return S_OK;
   }

   STDMETHOD(GetSpacing)(POINT* ppt)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::GetSpacing\n"));
      T* pT = static_cast<T*>(this); pT;
      DWORD dwSpacing = ListView_GetItemSpacing(pT->m_hwndList, FALSE);
      ppt->x = LOWORD(dwSpacing);
      ppt->y = HIWORD(dwSpacing);
      return S_OK;
   }

   STDMETHOD(GetDefaultSpacing)(POINT* ppt)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::GetDefaultSpacing\n"));
      ppt->x = ::GetSystemMetrics(SM_CXICONSPACING);
      ppt->y = ::GetSystemMetrics(SM_CYICONSPACING);
      return 0;
   }

   STDMETHOD(GetAutoArrange)(void)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::GetAutoArrange\n"));
      T* pT = static_cast<T*>(this); pT;
      DWORD dwStyle = ::GetWindowLong(pT->m_hwndList, GWL_STYLE);
      return (dwStyle & LVS_AUTOARRANGE) != 0 ? S_OK : S_FALSE;
   }

   STDMETHOD(SelectItem)(int iItemIndex, DWORD uFlags)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::SelectItem\n"));
      T* pT = static_cast<T*>(this); pT;
      LPCITEMIDLIST pidl = _GetPidlFromListIndex(iItemIndex);
      if( pidl == NULL ) return E_FAIL;
      return pT->SelectAndPositionItem(pidl, uFlags, NULL);
   }

   STDMETHOD(SelectAndPositionItems)(UINT cidl, LPCITEMIDLIST* apidl, POINT* apt, DWORD uFlags)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::SelectAndPositionItems\n"));
      T* pT = static_cast<T*>(this); pT;
      for( UINT i = 0; i < cidl; i++ ) {
         HRESULT Hr = pT->SelectAndPositionItem(apidl[i], uFlags, apt);
         if( FAILED(Hr) ) return Hr;
         uFlags &= ~(SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE | SVSI_FOCUSED | SVSI_EDIT);
      }
      return S_OK;
   }

   // IFolderView2

   STDMETHOD(SetGroupBy)(REFPROPERTYKEY, BOOL)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::SetGroupBy"));
   }

   STDMETHOD(GetGroupBy)(PROPERTYKEY*, BOOL*)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::GetGroupBy"));
   }

#if defined(_MSC_VER) && (_MSC_VER < 1300)
   STDMETHOD(SetViewProperty)(LPCITEMIDLIST, REFPROPERTYKEY, VARIANT&)
#else
   STDMETHOD(SetViewProperty)(LPCITEMIDLIST, REFPROPERTYKEY, REFPROPVARIANT)
#endif // _MSC_VER
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::SetViewProperty"));
   }

#if defined(_MSC_VER) && (_MSC_VER < 1300)
   STDMETHOD(GetViewProperty)(LPCITEMIDLIST, REFPROPERTYKEY, VARIANT*)
#else
   STDMETHOD(GetViewProperty)(LPCITEMIDLIST, REFPROPERTYKEY, PROPVARIANT*)
#endif // _MSC_VER
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::GetViewProperty"));
   }

   STDMETHOD(SetTileViewProperties)(LPCITEMIDLIST, LPCWSTR)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::SetTileViewProperties"));
   }

   STDMETHOD(SetExtendedTileViewProperties)(LPCITEMIDLIST, LPCWSTR)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::SetExtendedTileViewProperties"));
   }

   STDMETHOD(SetText)(FVTEXTTYPE, LPCWSTR)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::SetText"));
   }

   STDMETHOD(SetCurrentFolderFlags)(DWORD dwMask, DWORD dwValue)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::SetCurrentFolderFlags\n"));
      T* pT = static_cast<T*>(this); pT;
      pT->m_FolderSettings.fFlags = (pT->m_FolderSettings.fFlags & ~dwMask) | (dwValue & dwMask);
      if( ::IsWindow(pT->m_hwndList) ) pT->_UpdateShellSettings();
      return S_OK;
   }

   STDMETHOD(GetCurrentFolderFlags)(DWORD* pFlags)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::GetCurrentFolderFlags\n"));
      T* pT = static_cast<T*>(this); pT;
      *pFlags = pT->m_FolderSettings.fFlags;
      return S_OK;
   }

   STDMETHOD(GetSortColumnCount)(int*)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::GetSelectionMarkedItem"));
   }

   STDMETHOD(SetSortColumns)(const SORTCOLUMN*, int)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::SetSortColumns"));
   }

   STDMETHOD(GetSortColumns)(SORTCOLUMN*, int)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::GetSortColumns"));
   }

   STDMETHOD(GetItem)(int, REFIID, void**)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::GetItem"));
   }

   STDMETHOD(GetVisibleItem)(int iStartIndex, BOOL fPrevious, int* piIndex)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::GetVisibleItem\n"));
      T* pT = static_cast<T*>(this); pT;
#ifndef LVNI_VISIBLEONLY
      const UINT LVNI_PREVIOUS = 0x0020;
      const UINT LVNI_VISIBLEONLY = 0x0040;
#endif // LVNI_VISIBLEONLY
      *piIndex = ListView_GetNextItem(pT->m_hwndList, iStartIndex, LVNI_VISIBLEONLY | (fPrevious ? LVNI_PREVIOUS : 0));
      return *piIndex == -1 ? S_FALSE : S_OK;
   }

   STDMETHOD(GetSelectedItem)(int iStartIndex, int* piIndex)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::GetSelectedItem\n"));
      T* pT = static_cast<T*>(this); pT;
      *piIndex = ListView_GetNextItem(pT->m_hwndList, iStartIndex, LVNI_SELECTED);
      return *piIndex == -1 ? S_FALSE : S_OK;
   }

   STDMETHOD(GetSelection)(BOOL fNoneImpliesFolder, IShellItemArray** ppsia)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::GetSelection\n"));
      T* pT = static_cast<T*>(this); pT;
      if( fNoneImpliesFolder ) return pT->Items(SVGIO_SELECTION, IID_IShellItemArray, (LPVOID*) ppsia);
      return pT->GetItemObject(SVGIO_SELECTION, IID_IShellItemArray, (LPVOID*) ppsia);
   }

   STDMETHOD(GetSelectionState)(LPCITEMIDLIST pidl, DWORD* puState)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::GetSelectionState\n"));
      T* pT = static_cast<T*>(this); pT;
      *puState = 0;
      int iItem = _GetListIndexFromPidl(pidl);
      if( iItem == -1 ) return E_FAIL;
      DWORD dwState = ListView_GetItemState(pT->m_hwndList, iItem, LVIS_SELECTED|LVIS_FOCUSED);
      if( (dwState & LVIS_FOCUSED) != 0 ) *puState = SVSI_FOCUSED;
      if( (dwState & LVIS_SELECTED) != 0 ) *puState = SVSI_SELECT;
      return S_OK;
   }

   STDMETHOD(InvokeVerbOnSelection)(LPCSTR)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::InvokeVerbOnSelection"));
   }

   STDMETHOD(SetViewModeAndIconSize)(FOLDERVIEWMODE uViewMode, int iIconSize)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::SetViewModeAndIconSize\n"));
      T* pT = static_cast<T*>(this); pT;
      pT->m_FolderSettings.ViewMode = (int) uViewMode;
      pT->m_iIconSize = iIconSize;
      pT->_InitListView();
      pT->_FillListView();
      return S_OK;
   }

   STDMETHOD(GetViewModeAndIconSize)(FOLDERVIEWMODE* puViewMode, int* piIconSize)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::GetViewModeAndIconSize\n"));
      T* pT = static_cast<T*>(this); pT;
      *puViewMode = (FOLDERVIEWMODE) pT->m_FolderSettings.ViewMode;
      *piIconSize = (::GetWindowLong(pT->m_hwndList, GWL_STYLE) & LVS_ICON) == 0 ? ::GetSystemMetrics(SM_CXSMICON) : pT->m_iIconSize;
      return S_OK;
   }

   STDMETHOD(SetGroupSubsetCount)(UINT)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::SetGroupSubsetCount"));
   }

   STDMETHOD(GetGroupSubsetCount)(UINT*)
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::GetGroupSubsetCount"));
   }

   STDMETHOD(SetRedraw)(BOOL bRedraw)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::SetRedraw\n"));
      T* pT = static_cast<T*>(this); pT;
      ::SendMessage(pT->m_hwndList, WM_SETREDRAW, (WPARAM) bRedraw, 0L);
      return S_OK;
   }

   STDMETHOD(IsMoveInSameFolder)()
   {
      ATLTRACENOTIMPL(_T("IFolderView2Impl::IsMoveInSameFolder"));
   }

   STDMETHOD(DoRename)()
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IFolderView2Impl::DoRename\n"));
      T* pT = static_cast<T*>(this); pT;
      ::SetFocus(pT->m_hwndList);
      return ListView_EditLabel(pT->m_hwndList, ListView_GetNextItem(pT->m_hwndList, -1, LVNI_SELECTED)) == NULL ? E_FAIL : S_OK;
   }

   // Implementation

   LPCITEMIDLIST _GetPidlFromListIndex(int iItemIndex)
   {
      T* pT = static_cast<T*>(this); pT;
      LVITEM lvItem = { 0 };
      lvItem.mask = LVIF_PARAM;
      lvItem.iItem = iItemIndex;
      if( !ListView_GetItem(pT->m_hwndList, &lvItem) ) return NULL;
      return reinterpret_cast<LPCITEMIDLIST>(lvItem.lParam);
   }

   int _GetListIndexFromPidl(LPCITEMIDLIST pidl)
   {
      T* pT = static_cast<T*>(this); pT;
      for( int iItem = -1; (iItem = ListView_GetNextItem(pT->m_hwndList, iItem, LVNI_ALL)) != -1; ) {
         LPCITEMIDLIST pidlList = _GetPidlFromListIndex(iItem);
         if( pidlList == NULL ) continue;
         if( pT->m_pFolder->_CompareItems(COL_NAME, pidlList, pidl) == 0 ) return iItem;
      }
      return -1;
   }
};


//////////////////////////////////////////////////////////////////////////////
// CShellConnectionPointContainer

template< typename T >
class ATL_NO_VTABLE CShellConnectionPointContainer : 
   public IConnectionPointContainerImpl<T>,
   public IConnectionPointImpl<T, &DIID_DShellFolderViewEvents>
{
public:

BEGIN_CONNECTION_POINT_MAP(T)
   CONNECTION_POINT_ENTRY(DIID_DShellFolderViewEvents)
END_CONNECTION_POINT_MAP()

   STDMETHOD(Advise)(IUnknown* pUnkSink, DWORD* pdwCookie)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("CShellConnectionPointContainer::Advise\n"));
      if( pUnkSink == NULL || pdwCookie == NULL ) return E_POINTER;
      IUnknown* p = NULL;
      HRESULT Hr = pUnkSink->QueryInterface(IID_IDispatch, (LPVOID*) &p);
      if( SUCCEEDED(Hr) ) {
         *pdwCookie = m_vec.Add(p);
         Hr = (*pdwCookie != NULL) ? S_OK : CONNECT_E_ADVISELIMIT;
         if( Hr != S_OK ) p->Release();
      }
      if( Hr == E_NOINTERFACE ) Hr = CONNECT_E_CANNOTCONNECT;
      if( FAILED(Hr) ) *pdwCookie = 0;
      return Hr;
   }

   HRESULT FireAutomationEvent(DISPID dispid)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("CShellConnectionPointContainer::FireAutomationEvent\n"));
      for( int nIndex = 0; nIndex < m_vec.GetSize(); nIndex++ ) {
         CComQIPtr<IDispatch> spDisp = m_vec.GetAt(nIndex);
         if( spDisp != NULL ) {
            CComVariant vResult;
            DISPPARAMS disp = { NULL, NULL, 0, 0 };
            spDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, &vResult, NULL, NULL);
         }
      }
      return S_OK;
   }
};


//////////////////////////////////////////////////////////////////////////////
// CShellFolderViewDual

class ATL_NO_VTABLE CShellFolderViewDual : 
   public CComObjectRootEx<CComSingleThreadModel>,
   public CShellConnectionPointContainer<CShellFolderViewDual>,
   public IDispatchImpl<IShellFolderViewDual, &__uuidof(IShellFolderViewDual), &LIBID_Shell32>
{
public:
   BEGIN_COM_MAP(CShellFolderViewDual)
      COM_INTERFACE_ENTRY(IDispatch)
      COM_INTERFACE_ENTRY_IID(IID_IConnectionPointContainer, IConnectionPointContainer)
   END_COM_MAP()

   STDMETHOD(get_Application)(IDispatch** /*ppid*/)
   {
      ATLTRACENOTIMPL(_T("CShellFolderViewDual::get_Application"));
   }
           
   STDMETHOD(get_Parent)(IDispatch** /*ppid*/)
   {
      ATLTRACENOTIMPL(_T("CShellFolderViewDual::get_Parent"));
   }
           
   STDMETHOD(get_Folder)(Folder** /*ppid*/)
   {
      ATLTRACENOTIMPL(_T("CShellFolderViewDual::get_Folder"));
   }
           
   STDMETHOD(SelectedItems)(FolderItems** /*ppid*/)
   {
      ATLTRACENOTIMPL(_T("CShellFolderViewDual::SelectedItems"));
   }

   STDMETHOD(get_FocusedItem)(FolderItem** /*ppid*/)
   {
      ATLTRACENOTIMPL(_T("CShellFolderViewDual::get_FocusedItem"));
   }
           
   STDMETHOD(SelectItem)(VARIANT* /*pvfi*/, int /*dwFlags*/)
   {
      ATLTRACENOTIMPL(_T("CShellFolderViewDual::SelectItem"));
   }
           
   STDMETHOD(PopupItemMenu)(FolderItem* /*pfi*/, VARIANT /*vx*/, VARIANT /*vy*/, BSTR* /*pbs*/)
   {
      ATLTRACENOTIMPL(_T("CShellFolderViewDual::PopupItemMenu"));
   }

   STDMETHOD(get_Script)(IDispatch** /*ppDisp*/)
   {
      ATLTRACENOTIMPL(_T("CShellFolderViewDual::get_Script"));
   }

   STDMETHOD(get_ViewOptions)(long* /*plViewOptions*/)
   {
      ATLTRACENOTIMPL(_T("CShellFolderViewDual::get_ViewOptions"));
   }
};


//////////////////////////////////////////////////////////////////////////////
// IShellSearchTargetStub

MIDL_INTERFACE("9A7A94F5-FBF1-49A8-B0D9-44667635FE97")
IShellSearchTargetStub : public IUnknown
{
public:
   STDMETHOD(ShellSearchTarget01)() PURE;
   STDMETHOD(ShellSearchTarget02)(LPCWSTR) PURE;
   STDMETHOD(ShellSearchTarget03)() PURE;
   STDMETHOD(ShellSearchTarget04)() PURE;
   STDMETHOD(ShellSearchTarget05)() PURE;
   STDMETHOD(ShellSearchTarget06)() PURE;
   STDMETHOD(ShellSearchTarget07)() PURE;
   STDMETHOD(ShellSearchTarget08)() PURE;
   STDMETHOD(ShellSearchTarget09)() PURE;
   STDMETHOD(ShellSearchTarget10)() PURE;
   STDMETHOD(ShellSearchTarget11)() PURE;
   STDMETHOD(ShellSearchTarget12)() PURE;
   STDMETHOD(ShellSearchTarget13)() PURE;
};

template< typename T >
class IShellSearchTargetStubImpl : public IShellSearchTargetStub
{
public:
   STDMETHOD(ShellSearchTarget01)() { return E_NOTIMPL; }
   STDMETHOD(ShellSearchTarget02)(LPCWSTR) { return E_NOTIMPL; }
   STDMETHOD(ShellSearchTarget03)() { return E_NOTIMPL; }
   STDMETHOD(ShellSearchTarget04)() { return E_NOTIMPL; }
   STDMETHOD(ShellSearchTarget05)() { return E_NOTIMPL; }
   STDMETHOD(ShellSearchTarget06)() { return E_NOTIMPL; }
   STDMETHOD(ShellSearchTarget07)() { return E_NOTIMPL; }
   STDMETHOD(ShellSearchTarget08)() { return E_NOTIMPL; }
   STDMETHOD(ShellSearchTarget09)() { return E_NOTIMPL; }
   STDMETHOD(ShellSearchTarget10)() { return E_NOTIMPL; }
   STDMETHOD(ShellSearchTarget11)() { return E_NOTIMPL; }
   STDMETHOD(ShellSearchTarget12)() { return E_NOTIMPL; }
   STDMETHOD(ShellSearchTarget13)() { return E_NOTIMPL; }
};


//////////////////////////////////////////////////////////////////////////////
// IShellFolderView for ListView control

template< class T >
class IShellFolderViewImpl : public IShellFolderView
{
public:
   CComPtr<IDispatch> m_spAutomation;

   STDMETHOD(Rearrange)(LPARAM /*lParamSort*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::Rearrange"));
   }

   STDMETHOD(GetArrangeParam)(LPARAM* /*plParamSort*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::GetArrangeParam"));
   }

   STDMETHOD(ArrangeGrid)(void)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::ArrangeGrid"));
   }

   STDMETHOD(AutoArrange)(void)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::AutoArrange"));
   }

   STDMETHOD(GetAutoArrange)(void)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::GetAutoArrange"));
   }

   STDMETHOD(AddObject)(LPITEMIDLIST /*pidl*/, UINT* /*puItem*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::AddObject"));
   }

#if defined(_MSC_VER) && (_MSC_VER < 1300)
   STDMETHOD(GetObject)(LPITEMIDLIST /*pidl*/, UINT /*puItem*/)
#else
   STDMETHOD(GetObject)(LPITEMIDLIST* /*pidl*/, UINT /*puItem*/)
#endif // _MSC_VER
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::GetObject"));
   }

   STDMETHOD(RemoveObject)(LPITEMIDLIST /*pidl*/, UINT* /*puItem*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::RemoveObject"));
   }

   STDMETHOD(GetObjectCount)(UINT* puCount)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellFolderView::GetObjectCount\n"));
      T* pT = static_cast<T*>(this); pT;
      *puCount = (UINT) ListView_GetItemCount(pT->m_hwndList);
      return S_OK;
   }

   STDMETHOD(SetObjectCount)(UINT /*uCount*/, UINT /*dwFlags*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::SetObjectCount"));
   }

   STDMETHOD(UpdateObject)(LPITEMIDLIST /*pidlOld*/, LPITEMIDLIST /*pidlNew*/, UINT* /*puItem*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::UpdateObject"));
   }

   STDMETHOD(RefreshObject)(LPITEMIDLIST /*pidl*/, UINT* /*puItem*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::RefreshObject"));
   }

   STDMETHOD(SetRedraw)(BOOL bRedraw)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellFolderView::SetRedraw\n"));
      T* pT = static_cast<T*>(this); pT;
      ::SendMessage(pT->m_hwndList, WM_SETREDRAW, (WPARAM) bRedraw, 0L);
      return S_OK;
   }

   STDMETHOD(GetSelectedCount)(UINT* puSelected)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellFolderView::GetSelectedCount\n"));
      T* pT = static_cast<T*>(this); pT;
      *puSelected = ListView_GetSelectedCount(pT->m_hwndList);
      return S_OK;
   }

#if defined(_MSC_VER) && (_MSC_VER < 1300)
   STDMETHOD(GetSelectedObjects)(LPITEMIDLIST** /*ppidl*/, UINT* /*puItems*/)
#else
   STDMETHOD(GetSelectedObjects)(LPCITEMIDLIST** /*ppidl*/, UINT* /*puItems*/)
#endif // _MSC_VER
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::GetSelectedObjects"));
   }

   STDMETHOD(IsDropOnSource)(IDropTarget* /*DropTarget*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::IsDropOnSource"));
   }

   STDMETHOD(GetDragPoint)(POINT* /*ppt*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::GetDragPoint"));
   }

   STDMETHOD(GetDropPoint)(POINT* /*ppt*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::GetDropPoint"));
   }

   STDMETHOD(MoveIcons)(IDataObject* /*pDataObject*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::MoveIcons"));
   }

#if defined(_MSC_VER) && (_MSC_VER < 1300)
   STDMETHOD(SetItemPos)(LPITEMIDLIST /*pidl*/, POINT* /*ppt*/)
#else
   STDMETHOD(SetItemPos)(LPCITEMIDLIST /*pidl*/, POINT* /*ppt*/)
#endif // _MSC_VER
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::SetItemPos"));
   }

   STDMETHOD(IsBkDropTarget)(IDropTarget* /*pDropTarget*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::IsBkDropTarget"));
   }

   STDMETHOD(SetClipboard)(BOOL /*bMove*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::SetClipboard"));
   }

   STDMETHOD(SetPoints)(IDataObject* /*pDataObject*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::SetPoints"));
   }

   STDMETHOD(GetItemSpacing)(ITEMSPACING* /*spacing*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::GetItemSpacing"));
   }

   STDMETHOD(SetCallback)(IShellFolderViewCB* /*pNewCB*/, IShellFolderViewCB** /*ppOldCB*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::SetCallback"));
   }

   STDMETHOD(Select)(UINT dwFlags)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellFolderView::Select\n"));
      T* pT = static_cast<T*>(this); pT;
      ::SetFocus(pT->m_hwndList);
      switch( dwFlags ) {
      case 0: ListView_SetItemState(pT->m_hwndList, -1, 0, LVIS_SELECTED); return S_OK;
      case 1: ListView_SetItemState(pT->m_hwndList, -1, LVIS_SELECTED, LVIS_SELECTED); return S_OK;
      }
      return E_NOTIMPL;
   }

   STDMETHOD(QuerySupport)(UINT* /*pdwSupport*/)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::QuerySupport"));
   }

   STDMETHOD(SetAutomationObject)(IDispatch* pDisp)
   {
      ATLTRACENOTIMPL(_T("IShellFolderViewImpl::SetAutomationObject"));
      m_spAutomation = pDisp;
      return S_OK;
   }
};


//////////////////////////////////////////////////////////////////////////////
// Windows XP Task band

MIDL_INTERFACE("EC6FE84F-DC14-4FBB-889F-EA50FE27FE0F")
IUIElement : public IUnknown
{
   STDMETHOD(get_Name)(IShellItemArray* pItemArray, LPWSTR* pwstrName) PURE;
   STDMETHOD(get_Icon)(IShellItemArray* pItemArray, LPWSTR* pwstrName) PURE;
   STDMETHOD(get_Tooltip)(IShellItemArray* pItemArray, LPWSTR* pwstrName) PURE;
};

MIDL_INTERFACE("4026DFB9-7691-4142-B71C-DCF08EA4DD9C")
IUICommand : public IUIElement
{
    STDMETHOD(get_CanonicalName)(GUID* pGuid) PURE;
    STDMETHOD(get_State)(IShellItemArray* pItemArray, int nRequested, enum UISTATE* pState) PURE;
    STDMETHOD(Invoke)(IShellItemArray* pItemArray, IBindCtx* pCtx) PURE;
};

MIDL_INTERFACE("869447DA-9F84-4E2A-B92D-00642DC8A911")
IEnumUICommand : public IUnknown
{
    STDMETHOD(Next)(ULONG celt, IUICommand* rgelt, ULONG* pceltFetched) PURE;
    STDMETHOD(Skip)(ULONG celt) PURE;
    STDMETHOD(Reset)(VOID) PURE;
    STDMETHOD(Clone)(IEnumUICommand**ppenum) PURE;
};

#define SFVM_GET_WEBVIEW_CONTENT  83
#define SFVM_GET_WEBVIEW_TASKS    84

struct SFVM_WEBVIEW_CONTENT_DATA
{
   long l1;
   long l2;
   IUIElement* pUIElement1;
   IUIElement* pUIElement2;
   IEnumIDList* pEnum;
};

struct SFVM_WEBVIEW_TASKSECTION_DATA
{
   IEnumUICommand* pEnum1;
   IEnumUICommand* pEnum2;
   BOOL fUnknown;
}; 


//////////////////////////////////////////////////////////////////////////////
// IShellFolderViewType

MIDL_INTERFACE("49422C1E-1C03-11d2-8DAB-0000F87A556C")
IShellFolderViewType : public IUnknown
{
   STDMETHOD(EnumViews)(ULONG grfFlags, IEnumIDList** ppenum) PURE;
   STDMETHOD(GetDefaultViewName)(DWORD  uFlags, LPWSTR* ppwszName) PURE;
   STDMETHOD(GetViewTypeProperties)(LPCITEMIDLIST pidl, DWORD* pdwFlags)  PURE;
   STDMETHOD(TranslateViewPidl)(LPCITEMIDLIST pidl, LPCITEMIDLIST pidlView, LPCITEMIDLIST* ppidlOut) PURE;
};

#define SFVTFLAG_NOTIFY_CREATE  0x00000001
#define SFVTFLAG_NOTIFY_RESORT  0x00000002


//////////////////////////////////////////////////////////////////////////////
// IShellFolderSearchable

MIDL_INTERFACE("4E1AE66C-204B-11d2-8DB3-0000F87A556C")
IShellFolderSearchable : public IUnknown
{
   STDMETHOD(FindString)(LPCWSTR pwszTarget, DWORD* pdwFlags, IUnknown* punkOnAsyncSearch, LPITEMIDLIST* ppidlOut) PURE;
   STDMETHOD(CancelAsyncSearch)(LPCITEMIDLIST pidlSearch, DWORD* pdwFlags) PURE;
   STDMETHOD(InvalidateSearch)(LPCITEMIDLIST pidlSearch, DWORD* pdwFlags) PURE;
};

MIDL_INTERFACE("F98D8294-2BBC-11d2-8DBD-0000F87A556C")
IShellFolderSearchableCallback : public IUnknown
{
   STDMETHOD(RunBegin)(DWORD dwReserved) PURE;
   STDMETHOD(RunEnd)(DWORD dwReserved) PURE;
};


//////////////////////////////////////////////////////////////////////////////
// CShellPropertyPage

template< class T >
class ATL_NO_VTABLE CShellPropertyPage :
   public IShellPropSheetExt,
   public IShellExtInit,
   public CWindow
{
public:
   TCHAR m_szCaption[80];
   TCHAR m_szFileName[MAX_PATH];
   const MSG* m_pCurrentMsg;
   bool m_bSeenAddRef;

   CShellPropertyPage() : m_bSeenAddRef(false)
   {
      ::ZeroMemory(m_szCaption, sizeof(m_szCaption));
      ::ZeroMemory(m_szFileName, sizeof(m_szFileName));      
   }

   // IShellPropSheetExt

   STDMETHOD(AddPages)(LPFNADDPROPSHEETPAGE pfnAddPage, LPARAM lParam)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellPropSheetExt::AddPages\n"));

      T* pT = static_cast<T*>(this); pT;

      ::LoadString(_Module.GetResourceInstance(), T::IDS_TABCAPTION, m_szCaption, (sizeof(m_szCaption) / sizeof(m_szCaption[0])) - 1);

      PROPSHEETPAGE psp = { 0 };
      psp.dwSize        = sizeof(psp);
      psp.dwFlags       = PSP_USETITLE | PSP_USECALLBACK;
      psp.hInstance     = _Module.GetResourceInstance();
      psp.pszTemplate   = MAKEINTRESOURCE(T::IDD);
      psp.hIcon         = 0;
      psp.pszTitle      = m_szCaption;
      psp.pfnDlgProc    = (DLGPROC) T::PageDlgProc;
      psp.pcRefParent   = NULL;
      psp.pfnCallback   = T::PropSheetPageProc;
      psp.lParam        = (LPARAM) pT;
      HPROPSHEETPAGE hPage = ::CreatePropertySheetPage(&psp);            
      if( hPage == NULL ) return E_OUTOFMEMORY;

      if( pfnAddPage(hPage, lParam) == FALSE ) {
         ::DestroyPropertySheetPage(hPage);
         return E_FAIL;
      }

      return MAKE_HRESULT(SEVERITY_SUCCESS, 0, T::ID_TAB_INDEX); // COMCTRL Ver 4.71 allows us to set the initial page index
   }

   STDMETHOD(ReplacePage)(UINT /*uPageID*/, LPFNADDPROPSHEETPAGE /*lpfnReplaceWith*/, LPARAM /*lParam*/) 
   {
      // The Shell doesn't call this for file class Property Sheets
      ATLTRACENOTIMPL(_T("IShellPropSheetExt::ReplacePage"));
   }

   // IShellExtInit

   STDMETHOD(Initialize)(LPCITEMIDLIST /*pidlFolder*/, IDataObject* pDataObject, HKEY /*hkeyProgID*/) 
   {
      ATLTRACE2(atlTraceCOM, 0, _T("IShellExtInit::Initialize\n"));
      ATLASSERT(pDataObject);
      STGMEDIUM medium = { 0 };
      FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
      if( pDataObject != NULL && SUCCEEDED( pDataObject->GetData(&fe, &medium) ) ) {
         // Get the file name from the CF_HDROP.
         HDROP hDrop = (HDROP) ::GlobalLock(medium.hGlobal);
         if( hDrop == NULL ) return E_POINTER;
         UINT uCount = ::DragQueryFile(hDrop, (UINT) -1, NULL, 0);
         if( uCount > 0 ) ::DragQueryFile(hDrop, 0, m_szFileName, MAX_PATH);
         ::GlobalUnlock(medium.hGlobal);
         ::ReleaseStgMedium(&medium);
      }
      return S_OK;
   }

   // Callbacks

   static UINT CALLBACK PropSheetPageProc(HWND /*hwnd*/, UINT uMsg, LPPROPSHEETPAGE ppsp)
   {
      ATLTRACE2(atlTraceCOM, 0, _T("CShellPropertyPage::PropSheetPageProc %ld\n"), uMsg);
      T* pT = reinterpret_cast<T*>(ppsp->lParam); pT;
      ATLASSERT(pT);
      if( pT == NULL ) return 0;
      switch( uMsg ) {
      case PSPCB_CREATE:
         if( !pT->m_bSeenAddRef ) pT->AddRef();
         return 1; // Allow dialog creation
#if (_WIN32_IE >= 0x0500)
      case PSPCB_ADDREF:
         pT->AddRef();
         pT->m_bSeenAddRef = true;
         break;
#endif // _WIN32_IE
      case PSPCB_RELEASE:
         pT->Release();
         break;
      }
      return 0;
   }

   void SetModified(BOOL bChanged = TRUE)
   {
      ATLASSERT(::IsWindow(m_hWnd));
      ATLASSERT(GetParent()!=NULL);
      ::SendMessage(GetParent(), bChanged ? PSM_CHANGED : PSM_UNCHANGED, (WPARAM) m_hWnd, 0L);
   }

   static int CALLBACK PageDlgProc(HWND hWnd, UINT uMessage, WPARAM wParam, LPARAM lParam)
   {    
      T* pT = reinterpret_cast<T*>(::GetWindowLongPtr(hWnd, GWLP_USERDATA));
      LRESULT lRes = 0;
      if( uMessage == WM_INITDIALOG ) {
         pT = reinterpret_cast<T*>(reinterpret_cast<LPPROPSHEETPAGE>(lParam)->lParam);
         ::SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR) pT);
         // Set the window handle
         pT->m_hWnd = hWnd;
         lRes = 1;
      }
      if( pT == NULL ) {
         // The first message might be WM_SETFONT not WM_INITDIALOG as expected.
         // Don't invoke default WindowProc since it will recurse!
         return FALSE;
      }
      MSG msg = { pT->m_hWnd, uMessage, wParam, lParam, 0, { 0, 0 } };
      const MSG* pOldMsg = pT->m_pCurrentMsg;
      pT->m_pCurrentMsg = &msg;
      // Pass to the message map to process
      BOOL bRet = pT->ProcessWindowMessage(pT->m_hWnd, uMessage, wParam, lParam, lRes, 0);
      // Restore saved value for the current message
      ATLASSERT(pT->m_pCurrentMsg==&msg);
      pT->m_pCurrentMsg = pOldMsg;
      // Set result if message was handled
      if( bRet ) {
         switch( uMessage ) {
         case WM_COMPAREITEM:
         case WM_VKEYTOITEM:
         case WM_CHARTOITEM:
         case WM_INITDIALOG:
         case WM_QUERYDRAGICON:
         case WM_CTLCOLORMSGBOX:
         case WM_CTLCOLOREDIT:
         case WM_CTLCOLORLISTBOX:
         case WM_CTLCOLORBTN:
         case WM_CTLCOLORDLG:
         case WM_CTLCOLORSCROLLBAR:
         case WM_CTLCOLORSTATIC:
            return (int) lRes;
         }
         ::SetWindowLongPtr(pT->m_hWnd, DWLP_MSGRESULT, lRes);
         return TRUE;
      }
      if( uMessage == WM_NCDESTROY ) {
         // Clear out window handle
         ::SetWindowLongPtr(hWnd, GWLP_USERDATA, 0L);
         pT->m_hWnd = NULL;
      }
      return FALSE;
   }
};


//////////////////////////////////////////////////////////////////////////////
// CExtractFileIcon

template< class T >
class ATL_NO_VTABLE CExtractFileIcon : 
   public CComObjectRootEx<CComSingleThreadModel>,
   public IExtractIconA,
   public IExtractIconW
{
public:
   CPidl m_pidl;

BEGIN_COM_MAP(CExtractFileIcon)
   COM_INTERFACE_ENTRY_IID(IID_IExtractIconA, IExtractIconA)
   COM_INTERFACE_ENTRY_IID(IID_IExtractIconW, IExtractIconW)
END_COM_MAP()

public:
   void Init(LPCITEMIDLIST pidl)
   {
      m_pidl.Copy(pidl);
   }

   // IExtractIconA

   STDMETHOD(GetIconLocation)(UINT uFlags, LPSTR /*szIconFile*/, UINT /*cchMax*/, LPINT piIndex, LPUINT puFlags)
   {
      ATLASSERT(piIndex);
      ATLASSERT(puFlags);
      if( m_pidl.IsEmpty() ) return E_FAIL;
      LPCITEMIDLIST pidlLast = CPidl::PidlGetLastItem(m_pidl);
      T* pData = reinterpret_cast<T*>(pidlLast);
      if( pData != NULL ) {
         switch( pData->type ) {
         case 0:
            *piIndex = 0;
            *puFlags = GIL_NOTFILENAME;
            break;
         case 1:
            *piIndex = 1;
            *puFlags = GIL_NOTFILENAME;
            break;
         default:
            *piIndex = 0;
            *puFlags = GIL_SIMULATEDOC;
            break;
         }
      }
      return S_OK;
   }

   STDMETHOD(Extract)(LPCSTR /*pszFile*/, UINT nIconIndex, HICON* phiconLarge, HICON* phiconSmall, UINT nIconSize)
   {
      if( m_pidl.IsEmpty() ) return E_FAIL;
      LPCITEMIDLIST pidlLast = CPidl::PidlGetLastItem(m_pidl);
      T* pData = reinterpret_cast<T*>(pidlLast);
      if( pData != NULL ) {
         switch( pData->type ) {
         case 0:
            {
               if( phiconLarge != NULL ) *phiconLarge = ImageList_GetIcon(_Module.m_ImageLists.m_hImageListLarge, 0, ILD_TRANSPARENT);
               if( phiconSmall != NULL ) *phiconSmall = ImageList_GetIcon(_Module.m_ImageLists.m_hImageListSmall, 0, ILD_TRANSPARENT);
            }
            return S_OK;
         case 1:
            {
               if( phiconLarge != NULL ) *phiconLarge = ImageList_GetIcon(_Module.m_ImageLists.m_hImageListLarge, 1, ILD_TRANSPARENT);
               if( phiconSmall != NULL ) *phiconSmall = ImageList_GetIcon(_Module.m_ImageLists.m_hImageListSmall, 1, ILD_TRANSPARENT);
            }
            return S_OK;
         case 2:
            {
               TCHAR szPath[MAX_PATH] = { 0 };
               ::lstrcpy(szPath, pData->ffd.cFileName);
               LPTSTR psz = szPath;
               TCHAR szExt[MAX_PATH] = { 0 };
               for( LPTSTR psz = szPath; *psz != '\0'; psz++ ) if( *psz == '.' ) ::lstrcpy(szExt, psz);
               SHFILEINFO sfi = { 0 };
               HIMAGELIST hImageListLarge = (HIMAGELIST) ::SHGetFileInfo(szExt, FILE_ATTRIBUTE_NORMAL, &sfi, sizeof(sfi), SHGFI_USEFILEATTRIBUTES|SHGFI_ICON|SHGFI_SYSICONINDEX);
               if( phiconLarge != NULL ) *phiconLarge = ImageList_GetIcon(hImageListLarge, sfi.iIcon, ILD_TRANSPARENT);
               HIMAGELIST hImageListSmall = (HIMAGELIST) ::SHGetFileInfo(szExt, FILE_ATTRIBUTE_NORMAL, &sfi, sizeof(sfi), SHGFI_USEFILEATTRIBUTES|SHGFI_ICON|SHGFI_SMALLICON|SHGFI_SYSICONINDEX);
               if( phiconSmall != NULL ) *phiconSmall = ImageList_GetIcon(hImageListSmall, sfi.iIcon, ILD_TRANSPARENT);
            }
            return S_OK;
         }
      }
      return E_FAIL;
   }

   // IExtractIconW

   STDMETHOD(GetIconLocation)(UINT uFlags, LPWSTR /*szIconFile*/, UINT /*cchMax*/, LPINT piIndex, LPUINT puFlags)
   {
      ATLASSERT(piIndex);
      ATLASSERT(puFlags);
      if( m_pidl.IsEmpty() ) return E_FAIL;
      LPCITEMIDLIST pidlLast = CPidl::PidlGetLastItem(m_pidl);
      T* pData = reinterpret_cast<T*>(pidlLast);
      if( pData != NULL ) {
         switch( pData->type ) {
         case 0:
            *piIndex = 0;
            *puFlags = GIL_NOTFILENAME;
            break;
         case 1:
            *piIndex = 1;
            *puFlags = GIL_NOTFILENAME;
            break;
         default:
            *piIndex = 0;
            *puFlags = GIL_SIMULATEDOC;
            break;
         }
      }
      return S_OK;
   }

   STDMETHOD(Extract)(LPCWSTR /*pszFile*/, UINT nIconIndex, HICON* phiconLarge, HICON* phiconSmall, UINT nIconSize)
   {
      if( m_pidl.IsEmpty() ) return E_FAIL;
      LPCITEMIDLIST pidlLast = CPidl::PidlGetLastItem(m_pidl);
      T* pData = reinterpret_cast<T*>(pidlLast);
      if( pData != NULL ) {
         switch( pData->type ) {
         case 0:
            {
               if( phiconLarge != NULL ) *phiconLarge = ImageList_GetIcon(_Module.m_ImageLists.m_hImageListLarge, 0, ILD_TRANSPARENT);
               if( phiconSmall != NULL ) *phiconSmall = ImageList_GetIcon(_Module.m_ImageLists.m_hImageListSmall, 0, ILD_TRANSPARENT);
            }
            return S_OK;
         case 1:
            {
               if( phiconLarge != NULL ) *phiconLarge = ImageList_GetIcon(_Module.m_ImageLists.m_hImageListLarge, 1, ILD_TRANSPARENT);
               if( phiconSmall != NULL ) *phiconSmall = ImageList_GetIcon(_Module.m_ImageLists.m_hImageListSmall, 1, ILD_TRANSPARENT);
            }
            return S_OK;
         case 2:
            {
               TCHAR szPath[MAX_PATH] = { 0 };
               ::lstrcpy(szPath, pData->ffd.cFileName);
               LPTSTR psz = szPath;
               TCHAR szExt[MAX_PATH] = { 0 };
               for( LPTSTR psz = szPath; *psz != '\0'; psz++ ) if( *psz == '.' ) ::lstrcpy(szExt, psz);
               SHFILEINFO sfi = { 0 };
               HIMAGELIST hImageListLarge = (HIMAGELIST) ::SHGetFileInfo(szExt, FILE_ATTRIBUTE_NORMAL, &sfi, sizeof(sfi), SHGFI_USEFILEATTRIBUTES|SHGFI_ICON|SHGFI_SYSICONINDEX);
               if( phiconLarge != NULL ) *phiconLarge = ImageList_GetIcon(hImageListLarge, sfi.iIcon, ILD_TRANSPARENT);
               HIMAGELIST hImageListSmall = (HIMAGELIST) ::SHGetFileInfo(szExt, FILE_ATTRIBUTE_NORMAL, &sfi, sizeof(sfi), SHGFI_USEFILEATTRIBUTES|SHGFI_ICON|SHGFI_SMALLICON|SHGFI_SYSICONINDEX);
               if( phiconSmall != NULL ) *phiconSmall = ImageList_GetIcon(hImageListSmall, sfi.iIcon, ILD_TRANSPARENT);
            }
            return S_OK;
         }
      }
      return E_FAIL;
   }
};


//////////////////////////////////////////////////////////////////////////////
// CShellContextMenu

template< class T >
struct _ATL_CTXMENU_ENTRY
{
  UINT id;
  UINT desc;
  void (__stdcall T::*pfn)(); // method to invoke
};

#define BEGIN_CONTEXTMENU_MAP(theClass) \
   static const _ATL_CTXMENU_ENTRY<theClass>* _GetCtxMap()\
   {\
      typedef theClass _atl_event_classtype;\
      static const _ATL_CTXMENU_ENTRY<_atl_event_classtype> _ctxmap[] = {

#define CONTEXTMENU_HANDLER(id, desc, func) \
    {id, desc, (void (__stdcall _atl_event_classtype::*)())func},\

#define END_CONTEXTMENU_MAP() {0,0,NULL}}; return _ctxmap;}


template< class T >
class ATL_NO_VTABLE CShellContextMenu : 
   public IContextMenu,
   public IShellExtInit
{
public:
   CRegKey m_regClass;
   CSimpleArray<CComBSTR> m_arrFiles;

   // IShellExtInit

   STDMETHOD(Initialize)(LPCITEMIDLIST pidlFolder, IDataObject* pDataObj, HKEY hkeyProgID)
   {
      ATLTRACE(_T("CShellContextMenu::Initialize\n"));
      // Get file class
      m_regClass.Open(hkeyProgID, NULL, KEY_READ);
      // Get files
      if( pDataObj != NULL ) {
         STGMEDIUM medium = { 0 };
         FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
         if( SUCCEEDED(pDataObj->GetData(&fe, &medium)) ) {
            // Get the filenames from the CF_HDROP.
            HDROP hDrop = (HDROP) ::GlobalLock(medium.hGlobal);
            UINT uCount = ::DragQueryFile(hDrop, (UINT) -1, NULL, 0);
            for( UINT i = 0; i < uCount; i++ ) {
               TCHAR szFileName[MAX_PATH] = { 0 };
               ::DragQueryFile(hDrop, i, szFileName, (sizeof(szFileName) / sizeof(szFileName[0])) - 1);
               CComBSTR bstrFilename = szFileName;
               m_arrFiles.Add(bstrFilename);
            }
            ::GlobalUnlock(medium.hGlobal);
            ::ReleaseStgMedium(&medium);
         }
      }    
      return S_OK;
    }

   // IContextMenu

   STDMETHOD(QueryContextMenu)(HMENU hMenu, UINT iIndexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
   {
      ATLTRACE(_T("CShellContextMenu::QueryContextMenu\n"));
      const _ATL_CTXMENU_ENTRY<T>* pMap = T::_GetCtxMap();
      UINT i = 0;
      while( pMap->pfn != NULL ) {  
        if( pMap->id == 0 ) {
          ::InsertMenu(hMenu, iIndexMenu++, MF_SEPARATOR | MF_STRING | MF_BYPOSITION, 0, _T(""));
        }
        else {
          TCHAR szText[128] = { 0 };
          ::LoadString(_Module.GetResourceInstance(), pMap->id, szText, (sizeof(szText) / sizeof(szText[0])) - 1);
          ::InsertMenu(hMenu, iIndexMenu++, MF_STRING | MF_BYPOSITION, idCmdFirst + i, szText);
        }
        i++;
        pMap++;
      }
      T* pT = static_cast<T*>(this);
      pT->UpdateMenu(hMenu, idCmdFirst);
      return MAKE_HRESULT(SEVERITY_SUCCESS, 0, i + 1);
   }

   STDMETHOD(InvokeCommand)(LPCMINVOKECOMMANDINFO lpcmi)
   {
      ATLTRACE(_T("CShellContextMenu::InvokeCommand\n"));
      if( HIWORD(lpcmi->lpVerb) ) return NOERROR; // The command is being sent via a verb
      UINT idxCmd = LOWORD(lpcmi->lpVerb);
      const _ATL_CTXMENU_ENTRY<T>* pMap = T::_GetCtxMap();
      UINT i = 0;
      while( pMap->pfn != NULL ) 
      {
         if( i == idxCmd ) 
         {
            T* pT = static_cast<T*>(this);
            CComStdCallThunk<T> thunk;
            thunk.Init(pMap->pfn, pT);
            VARIANT vRes;
            ::VariantInit(&vRes);
            return DispCallFunc(
               &thunk,
               0,
               CC_STDCALL,
               VT_EMPTY, // this is how DispCallFunc() represents VOID
               0,
               NULL,
               NULL,
               &vRes);
         }
         pMap++;
         i++;
      }
      return E_INVALIDARG;
   }

   STDMETHOD(GetCommandString)(UINT_PTR idCmd, UINT uFlags, LPUINT, LPSTR pszName, UINT cchMax)
   {
      ATLTRACE(_T("CShellContextMenu::GetCommandString\n"));
      switch( uFlags ) {
      case GCS_HELPTEXTA:
      case GCS_HELPTEXTW:
         {
            const _ATL_CTXMENU_ENTRY<T>* pMap = T::_GetCtxMap();
            UINT i = 0;
            while( pMap->pfn != NULL ) {
               if( i == idCmd ) {
                  // BUG: LoadStringW() is not supported on Win95
                  if( uFlags == GCS_HELPTEXTA ) ::LoadStringA(_Module.GetResourceInstance(), pMap->desc, pszName, cchMax);                 
                  else ::LoadStringW(_Module.GetResourceInstance(), pMap->desc, (LPWSTR) pszName, cchMax);
                  return S_OK;
               }
               pMap++;
               i++;
            }
         }
         return E_FAIL;
      case GCS_VERBA:
      case GCS_VERBW:
         return E_FAIL;
      case GCS_VALIDATE:
         return NOERROR;
      default:
         return E_NOTIMPL;
      }
   }

   void UpdateMenu(HMENU /*hMenu*/, UINT /*iCmdFirst*/) { };
};

#endif // __ATLCOM_H__


//////////////////////////////////////////////////////////////////////////////
// ::SHGetPathFromIDList() wrapper

class CShellPidlPath
{
public:
   TCHAR m_szPath[MAX_PATH];

   CShellPidlPath(LPCITEMIDLIST pidl)
   {
      ATLASSERT(pidl);     
      if( ::SHGetPathFromIDList(pidl, m_szPath) == FALSE ) m_szPath[0] = _T('\0');
   }

   operator LPCTSTR() const 
   { 
      return m_szPath; 
   }
};



#endif // __ATLSHELLEXT_H__

