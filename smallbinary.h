// A small header with MSVC linker options to include in your project.
//
// Use it to create smaller binaries.
// Don't use it if you're afraid it will break something.
//
// Copyright (c) 2001 Bjarke Viksoe

#ifndef _MSVC_LINKEROPTIONS_H_
#define _MSVC_LINKEROPTIONS_H_

#ifndef _DEBUG // Only in RELEASE builds

#ifdef _MSC_VER // Only Microsoft Visual C++

#pragma once

#if _MSC_VER < 1300
   // Setting this linker switch causes segment size to be set to 512 bytes
   #pragma comment(linker, "/OPT:NOWIN98")
#endif // _MSC_VER

#ifdef USE_UNSAFE_OPT
   // Merges the resource segment with the data segment
   // Resources will be read-only because of this.
   #pragma comment(linker, "/merge:.rdata=.text")
   #pragma comment(linker, "/IGNORE:4078")
   // Removes functions that are never referenced
   #pragma comment(linker, "/OPT:REF")
   // Removes redundant COMDATs
   #pragma comment(linker, "/OPT:ICF")
#endif // USE_UNSAFE_OPT

#endif // _MSC_VER

#endif // _DEBUG

#endif // _MSVC_LINKEROPTIONS_H_
