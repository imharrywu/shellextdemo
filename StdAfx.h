// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently

#if !defined(AFX_STDAFX_H__A4D5B0F0_E999_4251_B2A6_ECF57087C162__INCLUDED_)
#define AFX_STDAFX_H__A4D5B0F0_E999_4251_B2A6_ECF57087C162__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0601
#endif
#define _ATL_APARTMENT_THREADED

#include <atlbase.h>
#include "atlshellextbase.h"
extern CShellModule _Module;
#include <atlcom.h>
#include "atlshellext.h"
#include "atlwfile.h"

extern const CLSID CLSID_ContextMenu;


//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__A4D5B0F0_E999_4251_B2A6_ECF57087C162__INCLUDED)
