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
#include <stdlib.h>
#include <tchar.h>

#pragma comment(lib, "Dbghelp.lib")
#pragma comment(lib, "USER32.lib")
#pragma comment(lib, "OLE32.lib")
#pragma comment(lib, "OLEAUT32.lib")
#pragma comment(lib, "IMAGEHLP.lib")

#ifndef UNICODE
#error "This file must be compiled as unicode"
#endif

//============================================================================

LPCTSTR g_szHelpText = 	_T( "CoClassSyms - Matt Pietrek 1999 for MSJ\n" )
					 	_T( "  Syntax: CoClassSyms <filename>\n" )
						_T( "with patches from Sergey Kovalev\n");

LPCTSTR g_pszFileName= 0;

LOADED_IMAGE g_loadedImage;
FILE * g_pMapFile;
SIZE_T g_iSymbolsCount = 0;

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

BOOL VAToSectionOffset( PVOID address, SIZE_T &rva, SIZE_T &section, SIZE_T &offset );

BOOL CoClassSymsBeginSymbolCallouts( LPTYPELIB pITypeLib, LPCSTR pszExecutable );
BOOL CoClassSymsAddSymbol(
		SIZE_T pFunction,
		SIZE_T rva,
		SIZE_T section,
		SIZE_T offset,
		PSTR pszSymbolName );
BOOL CoClassSymsSymbolsFinished( void );

//============================================================================

extern "C" int _tmain( int argc, LPCTSTR * argv )
{
	CoInitialize( 0 );		// Initialize COM subsystem
	
	if ( 2 != argc )		// Command line processing.  No arguments causes
	{						// syntax and help to be displayed.
		_tprintf( g_szHelpText );
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
	char szFileName[MAX_PATH];
	wcstombs( szFileName, g_pszFileName, MAX_PATH );
	
	CoClassSymsBeginSymbolCallouts(pITypeLib, szFileName);

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

	CoClassSymsSymbolsFinished();

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
			else
				printf("WARNING Failed to GetRefTypeOfImplType!\n");
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
	{
		printf("WARNING Failed to GetRefTypeInfo!\n");
		return;
	}

	LPTYPEATTR pRefTypeAttr;
	pIRefTypeInfo->GetTypeAttr( &pRefTypeAttr );

	LPUNKNOWN pIUnknown = 0;
/*
	LPOLESTR pTypeAttrGUID;
	StringFromCLSID(pTypeAttr->guid, &pTypeAttrGUID);
	LPOLESTR pRefTypeAttrGUID;
	StringFromCLSID(pRefTypeAttr->guid, &pRefTypeAttrGUID);
	TCHAR pszInterfaceName[256];
	GetTypeInfoName( pITypeInfo_CoClass, pszInterfaceName );
	_tprintf( _T("\t* %-30s: TYPEKIND %d, CLSID %s, RefTypeGUID %s\n\n"), pszInterfaceName, pTypeAttr->typekind, pTypeAttrGUID, pRefTypeAttrGUID);
*/
	hr = CoCreateInstance( 	pTypeAttr->guid,
							0,					// pUnkOuter
							CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER,
							pRefTypeAttr->guid,
							(LPVOID *)&pIUnknown );
	if ( (S_OK == hr) && pIUnknown )
	{
		/**********************
		 *        DEBUG       *
		 **********************/
		{
			printf("\n\n=============== STR ===============\n\n\n");
			ITypeLib *pITypeLib = nullptr;
			UINT pIndex = 0;
			hr = pITypeInfo_CoClass->GetContainingTypeLib(&pITypeLib, &pIndex);
			UINT tiCount = pITypeLib->GetTypeInfoCount();
			
			for ( UINT i = 0; i < tiCount; i++ )
			{
				LPTYPEINFO pITypeInfo_dbg;

				HRESULT hr = pITypeLib->GetTypeInfo( i, &pITypeInfo_dbg );

				if ( S_OK == hr )
				{
					LPTYPEATTR pTypeAttr_dbg;
					hr = pITypeInfo_dbg->GetTypeAttr( &pTypeAttr_dbg );
					if ( S_OK != hr )
						return;
					
					if ( pTypeAttr->guid != pTypeAttr_dbg->guid &&
					     ( TKIND_COCLASS == pTypeAttr_dbg->typekind ||
						   TKIND_INTERFACE == pTypeAttr_dbg->typekind ||
						   TKIND_DISPATCH == pTypeAttr_dbg->typekind) )
					{
						for ( unsigned short i = 0; i < pTypeAttr_dbg->cImplTypes; i++ )
						{
							HREFTYPE hRefType_dbg;

							hr = pITypeInfo_dbg->GetRefTypeOfImplType( i, &hRefType_dbg );
							
							if ( S_OK == hr )
							{
								// ProcessReferencedTypeInfo( pITypeInfo_dbg, pTypeAttr_dbg, hRefType_dbg );
								LPTYPEINFO pIRefTypeInfo_dbg;
								
								HRESULT hr = pITypeInfo_dbg->GetRefTypeInfo(hRefType_dbg, &pIRefTypeInfo_dbg);
								if ( S_OK != hr )
								{
									printf("WARNING Failed to GetRefTypeInfo!\n");
									return;
								}

								LPTYPEATTR pRefTypeAttr_dbg;
								pIRefTypeInfo_dbg->GetTypeAttr( &pRefTypeAttr_dbg );

								EnumTypeInfoMembers( pIRefTypeInfo_dbg, pRefTypeAttr_dbg, pIUnknown );
							}
							else
								printf("WARNING Failed to GetRefTypeOfImplType!\n");
						}
					}
					
					pITypeInfo_dbg->ReleaseTypeAttr( pTypeAttr_dbg );
									
					pITypeInfo_dbg->Release();
				}
			}
			printf("\n\n=============== END ===============\n\n\n");
		} /* DEBUG */

		EnumTypeInfoMembers( pIRefTypeInfo, pRefTypeAttr, pIUnknown );

		pIUnknown->Release();
	}
	else
		printf("\tWARNING Failed to CoCreateInstance!\n");

						
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
	// Make a pointer to the vtable.	
	PBYTE pVTable = (PBYTE)*(PSIZE_T)(lpUnknown);

	if ( 0 == pTypeAttr->cFuncs )	// Make sure at least one method!
	{
		printf("WARNING pTypeAttr->cFuncs == 0!\n");
		return;
	}

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
		SIZE_T pFunction = *(PSIZE_T)(pVTable + pFuncDesc->oVft);

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
		SIZE_T rva;
		SIZE_T section;
		SIZE_T offset;
		
		if ( VAToSectionOffset((PVOID)pFunction, rva, section, offset) )
			CoClassSymsAddSymbol( pFunction, rva, section, offset, pszMungedName );
		
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

BOOL CoClassSymsBeginSymbolCallouts( LPTYPELIB pITypeLib, LPCSTR pszExecutable )
{
	if ( !MapAndLoad( (LPSTR)pszExecutable, 0, &g_loadedImage, FALSE, TRUE ) )
	{
		printf( "Unable to access or load executable\n" );
		return 0;
	}

	char szExeBaseName[MAX_PATH];
	char szMapFileName[MAX_PATH];
	_splitpath( pszExecutable, 0, 0, szExeBaseName, 0 );
	sprintf( szMapFileName, "%s.MAP", szExeBaseName );
	
	g_pMapFile = fopen( szMapFileName, "wt" );
	if ( !g_pMapFile )
		return FALSE;

	printf(
			" Start         Length     Name                   Class\n" );

	PIMAGE_SECTION_HEADER pSectHdr = g_loadedImage.Sections;
			
	for ( 	unsigned i=1;
			i <= g_loadedImage.NumberOfSections;
			i++, pSectHdr++ )
	{
		printf(
					" %04X:00000000 %08XH %-23.8hs %s\n",
					i, pSectHdr->Misc.VirtualSize, pSectHdr->Name,
					pSectHdr->Characteristics & IMAGE_SCN_CNT_CODE
						? "CODE" : "DATA" );
	}

	printf(
		"\n  pFunction        RVA                Address         Publics by Value              Rva+Base\n\n");	

	fprintf( g_pMapFile, "{\n");
	TLIBATTR *pTypeLibAttrs = nullptr;
	BSTR pTypeLibName = nullptr;
	BSTR pTypeLibDoc = nullptr;
	if (S_OK != pITypeLib->GetLibAttr(&pTypeLibAttrs) ||
		S_OK != pITypeLib->GetDocumentation(-1, &pTypeLibName, &pTypeLibDoc, nullptr, nullptr))
	{
		_tprintf( _T("GetLibAttr failed or GetDocumentation\n"));
	}
	else
	{
		LPOLESTR guidString;
		StringFromCLSID(pTypeLibAttrs->guid, &guidString);

		char szLibGUID[MAX_PATH];
		wcstombs( szLibGUID, guidString, MAX_PATH );
		char szLibName[MAX_PATH];
		wcstombs( szLibName, pTypeLibName, MAX_PATH );


		_tprintf( _T("TypeLib \"%s\"\n\"%s\"\nCLSID: %s\n\n"), pTypeLibName, pTypeLibDoc, guidString);
		fprintf( g_pMapFile,
				 "    \"metadata\": {\n"
				 "        \"name\": \"%s\",\n"
				 "        \"guid\": \"%s\"",
				 szLibName, szLibGUID);
		if (pTypeLibDoc)
		{
			char szLibDoc[MAX_PATH];
			wcstombs( szLibDoc, pTypeLibDoc, MAX_PATH );
			fprintf ( g_pMapFile, ",\n        \"doc\": \"%s\"\n", szLibDoc);
		}
		else
			fprintf( g_pMapFile, "\n");
		fprintf( g_pMapFile, "},\n");

		// ensure memory is freed
		SysFreeString(pTypeLibName);
		SysFreeString(pTypeLibDoc);
		CoTaskMemFree(guidString);
		pITypeLib->ReleaseTLibAttr(pTypeLibAttrs);
	}
	fprintf( g_pMapFile, "    \"symbols\": {\n");

	return TRUE;
}

BOOL CoClassSymsAddSymbol(
		SIZE_T pFunction,
		SIZE_T rva,
		SIZE_T section,
		SIZE_T offset,
		PSTR pszSymbolName )
{
	if ( !g_pMapFile )
		return FALSE;

	printf( "%016IX %016IX %08IX:%08IX       %-32s\n",
			 pFunction, rva, section, offset, pszSymbolName );
	if (g_iSymbolsCount++)
		fprintf( g_pMapFile, ",\n");
	fprintf( g_pMapFile, "        \"%s\": { \"address\": %Iu }", pszSymbolName, rva);
				
	return true;
}
		
BOOL CoClassSymsSymbolsFinished( void )
{
	if ( !g_pMapFile )
		return FALSE;
	
	DWORD entryRVA
		= g_loadedImage.FileHeader->OptionalHeader.AddressOfEntryPoint;
	
	PIMAGE_SECTION_HEADER pSectHdr;
	
	pSectHdr = ImageRvaToSection( 	g_loadedImage.FileHeader,
									g_loadedImage.MappedAddress,
									entryRVA );
	if ( pSectHdr )
	{
		// Pointer math below!!!
		WORD section = (WORD)(pSectHdr - g_loadedImage.Sections) +1;
		DWORD offset = entryRVA - pSectHdr->VirtualAddress;
		
		printf( "\n entry point at        %04X:%08X\n",
			 	 section, offset );
	}
	
	fprintf( g_pMapFile, "\n    }\n}\n");
	fclose( g_pMapFile );

	UnMapAndLoad( &g_loadedImage );		// Undo the MapAndLoad call
	
	return TRUE;
}

//=============================================================================
// Convert a linear (virtual) address into a logical (section:offset) address
//=============================================================================
BOOL VAToSectionOffset( PVOID address, SIZE_T &rva, SIZE_T &section, SIZE_T &offset )
{
	MEMORY_BASIC_INFORMATION mbi;

	// Tricky way to get the containing module from a linear address	
	VirtualQuery( address, &mbi, sizeof(mbi) );

	// "AllocationBase" is the same as an HMODULE
	LPVOID hModule = (LPVOID)mbi.AllocationBase;

	// Use IMAGEHLP API to get a pointer to the PE header.
	PIMAGE_NT_HEADERS pNtHeaders = ImageNtHeader(hModule);
	if ( !pNtHeaders )
	{
		printf("WARNING FAiled to ImageNtHeader!\n");
		return FALSE;
	}
		
	// Calculate relative virtual address (RVA)
	rva = (SIZE_T)address - (SIZE_T)hModule;

	PIMAGE_SECTION_HEADER pSectHdr;

	// Use another IMAGEHLP API to find the section containing the RVA
	pSectHdr = ImageRvaToSection( pNtHeaders, hModule, rva );
	if ( !pSectHdr )
	{
		printf("WARNING FAiled to ImageRvaToSection!\n");
		return FALSE;
	}
		
	// Figure out the section number.  Warning: pointer math below!!!
	section = (WORD)(pSectHdr - IMAGE_FIRST_SECTION(pNtHeaders)) + 1;

	// Calculate offset within containing section
	offset = rva - pSectHdr->VirtualAddress;
	
	return TRUE;
}
