//==================================================
// CoClassSyms - Matt Pietrek 1999
// Microsoft Systems Journal, March 1999
// FILE: CoClassSyms.CPP
//==================================================
#include <windows.h>
#include <ole2.h>
#include <ocidl.h>
#include <imagehlp.h>
#include <stdio.h>
#include <tchar.h>
#include "CoClassSymsCallouts.h"

#pragma comment(lib, "Dbghelp.lib")
#pragma comment(lib, "USER32.lib")
#pragma comment(lib, "OLE32.lib")
#pragma comment(lib, "OLEAUT32.lib")

#ifndef UNICODE
#error "This file must be compiled as unicode"
#endif

//============================================================================

LPCTSTR g_szHelpText = 	_T( "CoClassSyms - Matt Pietrek 1999 for MSJ\n" )
					 	_T( "  Syntax: CoClassSyms <filename>\n" )
						_T( "with patches from Sergey Kovalev\n");

LPCTSTR g_pszFileName= 0;

PFNCCSCALLOUTBEGIN g_pfnCoClassSymsBeginSymbolCallouts = 0;
PFNCCSADDSYMBOL g_pfnCoClassSymsAddSymbol = 0;
PFNCCSFINISHED g_pfnCoClassSymsSymbolsFinished = 0;
					 
//============================================================================

void ProcessTypeLib( LPCTSTR pszFileName );
void EnumTypeLib( LPTYPELIB pITypeLib );
void ProcessTypeInfo( LPTYPEINFO pITypeInfo );
void ProcessReferencedTypeInfo( LPTYPEINFO pITypeInfo, LPTYPEATTR pTypeAttr,
								HREFTYPE hRefType );
void EnumTypeInfoMembers( LPTYPEINFO pITypeInfo, LPTYPEATTR pTypeAttr,
							LPUNKNOWN lpUnknown );
void GetTypeInfoName( LPTYPEINFO pITypeInfo, LPTSTR pszName,
						MEMBERID memid = MEMBERID_NIL );

BOOL HookupToCallouts( void );
BOOL VAToSectionOffset( PVOID address, WORD &section, DWORD &offset );

//============================================================================

extern "C" int _tmain( int argc, LPCTSTR * argv )
{
	CoInitialize( 0 );		// Initialize COM subsystem
	
	if ( 2 != argc )		// Command line processing.  No arguments causes
	{						// syntax and help to be displayed.
		_tprintf( g_szHelpText );
		return 0;
	}

	// Load the DLL that we'll write results to, and get the addresses of
	// the three callout APIs.
	if ( !HookupToCallouts() )
	{
		_tprintf( _T( "No output DLL found") );
		return 0;
	}
	
	g_pszFileName = argv[1];	// argv[1] == filename passed on command line
		
	ProcessTypeLib( g_pszFileName );
	
	CoUninitialize();
	
	return 0;
}

//============================================================================
// Given a filename for a typelib, attempt to get an ITypeLib for it.  Send
// the resultant ITypeLib instance to EnumTypeLib.
//============================================================================
void ProcessTypeLib( LPCTSTR pszFileName )
{
	LPTYPELIB pITypeLib;

	HRESULT hr = LoadTypeLib( pszFileName, &pITypeLib );
	if ( S_OK != hr )
	{
		_tprintf( _T("LoadTypeLib failed on file %s\n"), pszFileName );
		return;
	}

	EnumTypeLib( pITypeLib );
	
	pITypeLib->Release();
}

//============================================================================
// Enumerate through all the ITypeInfo instances in an ITypeLib.  Pass each
// instance to ProcessTypeInfo.
//============================================================================
void EnumTypeLib( LPTYPELIB pITypeLib )
{
	UINT tiCount = pITypeLib->GetTypeInfoCount();
	
	for ( UINT i = 0; i < tiCount; i++ )
	{
		LPTYPEINFO pITypeInfo;

		HRESULT hr = pITypeLib->GetTypeInfo( i, &pITypeInfo );

		if ( S_OK == hr )
		{
			ProcessTypeInfo( pITypeInfo );
							
			pITypeInfo->Release();
		}
	}

	g_pfnCoClassSymsSymbolsFinished();

}

//============================================================================
// Top level handling code for a single ITypeInfo extracted from a typelib
//============================================================================
void ProcessTypeInfo( LPTYPEINFO pITypeInfo )
{
	HRESULT hr;
		
	LPTYPEATTR pTypeAttr;
	hr = pITypeInfo->GetTypeAttr( &pTypeAttr );
	if ( S_OK != hr )
		return;
	
	if ( TKIND_COCLASS == pTypeAttr->typekind )
	{
		for ( unsigned short i = 0; i < pTypeAttr->cImplTypes; i++ )
		{
			HREFTYPE hRefType;
			
			hr = pITypeInfo->GetRefTypeOfImplType( i, &hRefType );
			
			if ( S_OK == hr )
				ProcessReferencedTypeInfo( pITypeInfo, pTypeAttr, hRefType );
 		}
	}
	
	pITypeInfo->ReleaseTypeAttr( pTypeAttr );
}

//============================================================================
// Given a TKIND_COCLASS ITypeInfo, get the ITypeInfo that describes the
// referenced (HREFTYPE) TKIND_DISPATCH or TKIND_INTERFACE.  Pass that
// ITypeInfo to EnumTypeInfoMembers.
//============================================================================
void ProcessReferencedTypeInfo( LPTYPEINFO pITypeInfo_CoClass,
								LPTYPEATTR pTypeAttr,
								HREFTYPE hRefType )
{
	LPTYPEINFO pIRefTypeInfo;
	
	HRESULT hr = pITypeInfo_CoClass->GetRefTypeInfo(hRefType, &pIRefTypeInfo);
	if ( S_OK != hr )
		return;

	LPTYPEATTR pRefTypeAttr;
	pIRefTypeInfo->GetTypeAttr( &pRefTypeAttr );

	LPUNKNOWN pIUnknown = 0;

	hr = CoCreateInstance( 	pTypeAttr->guid,
							0,					// pUnkOuter
							CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER,
							pRefTypeAttr->guid,
							(LPVOID *)&pIUnknown );

	if ( (S_OK == hr) && pIUnknown )
	{
		EnumTypeInfoMembers( pIRefTypeInfo, pRefTypeAttr, pIUnknown );

		pIUnknown->Release();
	}
						
	pIRefTypeInfo->ReleaseTypeAttr( pRefTypeAttr );
	pIRefTypeInfo->Release();
}

//============================================================================
// Enumerate through each member of an ITypeInfo.  Send the method name and
// address to the CoClassSymsAddSymbol function.
//=============================================================================
void EnumTypeInfoMembers( 	LPTYPEINFO pITypeInfo,	// The ITypeInfo to enum.
							LPTYPEATTR pTypeAttr,	// The associated TYPEATTR.
							LPUNKNOWN lpUnknown		// From CoCreateInstance.
						)
{
	// Only call CoClassSymsBeginSymbolCallout once	
	static BOOL fCalledBeginCallout = FALSE;
	if ( FALSE == fCalledBeginCallout )
	{
		char szFileName[MAX_PATH];
		wcstombs( szFileName, g_pszFileName, MAX_PATH );
		
		fCalledBeginCallout = g_pfnCoClassSymsBeginSymbolCallouts(szFileName);
	}

	// Make a pointer to the vtable.	
	PBYTE pVTable = (PBYTE)*(PDWORD)(lpUnknown);

	if ( 0 == pTypeAttr->cFuncs )	// Make sure at least one method!
		return;

	// Get the name of the ITypeInfo, to use as the interface name in the
	// symbol names we'll be constructing.
	TCHAR pszInterfaceName[256];
	GetTypeInfoName( pITypeInfo, pszInterfaceName );

	// Enumerate through each method, obtain it's name, address, and ship the
	// info off to CoClassSymsAddSymbol()
	for ( unsigned i = 0; i < pTypeAttr->cFuncs; i++ )
	{
		FUNCDESC * pFuncDesc;
		
		pITypeInfo->GetFuncDesc( i, &pFuncDesc );
		
		TCHAR pszMemberName[256];		
		GetTypeInfoName( pITypeInfo, pszMemberName, pFuncDesc->memid );

		// Index into the vtable to retrieve the method's virtual address
		DWORD pFunction = *(PDWORD)(pVTable + pFuncDesc->oVft);

		// Created the basic form of the symbol name in interface::method
		// form using ANSI characters
		char pszMungedName[512];
		wsprintfA(	pszMungedName,"%ls::%ls",
					pszInterfaceName,pszMemberName );

		INVOKEKIND invkind = pFuncDesc-> invkind;

		// If it's a property "get" or "put", append a meaningful ending.
		// The "put" and "get" will have identical names, so we want to
		// make them into unique names
		if ( INVOKE_PROPERTYGET == invkind )
			strcat( pszMungedName, "_get" );
		else if ( INVOKE_PROPERTYPUT == invkind )
			strcat( pszMungedName, "_put" );
		else if ( INVOKE_PROPERTYPUTREF == invkind )
			strcat( pszMungedName, "_putref" );
					

		// Convert the virtual address to a logical address
		unsigned short section;
		unsigned long offset;
		
		if ( VAToSectionOffset((PVOID)pFunction, section, offset) )
			g_pfnCoClassSymsAddSymbol( section, offset, pszMungedName );
		
		pITypeInfo->ReleaseFuncDesc( pFuncDesc );						
	}
}

//============================================================================
// Given an ITypeInfo instance, retrieve the name.
//=============================================================================
void GetTypeInfoName( LPTYPEINFO pITypeInfo, LPTSTR pszName, MEMBERID memid )
{
	BSTR pszTypeInfoName;
	HRESULT hr;
	
	hr = pITypeInfo->GetDocumentation( memid, &pszTypeInfoName, 0, 0, 0 );

	if ( S_OK != hr )
	{
		lstrcpy( pszName, _T("<unknown>") );
		return;
	}

	// Make a copy so that we can free the BSTR	allocated by ::GetDocumentation
	lstrcpyW( pszName, pszTypeInfoName );

	// Free the BSTR allocated by ::GetDocumentation
	SysFreeString( pszTypeInfoName );
}

//=============================================================================
// Hook up to one of the CoClassSymCallout DLLs via LoadLibrary/GetProcAddress
//=============================================================================
BOOL HookupToCallouts( void )
{
	// First try to load the DLL that will make a .DBG file
	HMODULE hModCalloutDLL = LoadLibrary( _T("CoClassSymsDBGFile.DLL") );

	// If the first DLL wasn't found, try the DLL that generates a .MAP file
	if ( !hModCalloutDLL )
		hModCalloutDLL = LoadLibrary( _T("CoClassSymsMapFile.DLL") );
		
	if ( !hModCalloutDLL )
		return FALSE;
		

	g_pfnCoClassSymsBeginSymbolCallouts = (PFNCCSCALLOUTBEGIN)
			GetProcAddress( hModCalloutDLL,	"CoClassSymsBeginSymbolCallouts" );
		
	g_pfnCoClassSymsAddSymbol = (PFNCCSADDSYMBOL)
			GetProcAddress( hModCalloutDLL, "CoClassSymsAddSymbol" );
			
	g_pfnCoClassSymsSymbolsFinished = (PFNCCSFINISHED)
			GetProcAddress( hModCalloutDLL, "CoClassSymsSymbolsFinished" );

	// Make sure all three functions were found.  Otherwise, return failure.
	if (!g_pfnCoClassSymsBeginSymbolCallouts ||
		!g_pfnCoClassSymsAddSymbol ||
		!g_pfnCoClassSymsAddSymbol )
		return FALSE;
		
	return TRUE;
}

//=============================================================================
// Convert a linear (virtual) address into a logical (section:offset) address
//=============================================================================
BOOL VAToSectionOffset( PVOID address, WORD &section, DWORD &offset )
{
	MEMORY_BASIC_INFORMATION mbi;

	// Tricky way to get the containing module from a linear address	
	VirtualQuery( address, &mbi, sizeof(mbi) );

	// "AllocationBase" is the same as an HMODULE
	LPVOID hModule = (LPVOID)mbi.AllocationBase;

	// Use IMAGEHLP API to get a pointer to the PE header.
	PIMAGE_NT_HEADERS pNtHeaders = ImageNtHeader(hModule);
	if ( !pNtHeaders )
		return FALSE;
		
	// Calculate relative virtual address (RVA)
	DWORD rva = (DWORD)address - (DWORD)hModule;

	PIMAGE_SECTION_HEADER pSectHdr;

	// Use another IMAGEHLP API to find the section containing the RVA
	pSectHdr = ImageRvaToSection( pNtHeaders, hModule, rva );
	if ( !pSectHdr )
		return FALSE;
		
	// Figure out the section number.  Warning: pointer math below!!!
	section = (WORD)(pSectHdr - IMAGE_FIRST_SECTION(pNtHeaders)) + 1;

	// Calculate offset within containing section
	offset = rva - pSectHdr->VirtualAddress;
	
	return TRUE;
}
