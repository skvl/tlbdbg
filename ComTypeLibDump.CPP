//==================================================
// COMTypeLibDump - Matt Pietrek 1998
// Microsoft Systems Journal, December 1998
// FILE: COMTypeLibDump.CPP
//==================================================
#include <windows.h>
#include <ole2.h>
#include <tchar.h>

//============================================================================

LPCTSTR g_szHelpText = 	_T( "COMTypeLibDump - Matt Pietrek 1999 for MSJ\n" )
					 	_T( "  Syntax: COMTypeLibDump <filename>\n" );
					 
//============================================================================

void DisplayTypeLib( LPCTSTR pszFileName );
void EnumTypeLib( LPTYPELIB pITypeLib );
void DisplayTypeInfo( LPTYPEINFO pITypeInfo );
void EnumTypeInfoMembers( LPTYPEINFO pITypeInfo, LPTYPEATTR pTypeAttr);
LPCTSTR GetTypeKindName( TYPEKIND typekind );
LPCTSTR GetInvokeKindName( INVOKEKIND invkind );

//============================================================================

extern "C" int _tmain( int argc, LPCTSTR * argv )
{
	CoInitialize( 0 );
	
	if ( 2 != argc )
	{
		_tprintf( g_szHelpText );
		return 0;
	}
	
	DisplayTypeLib( argv[1] );
	
	CoUninitialize();
	
	return 0;
}

void DisplayTypeLib( LPCTSTR pszFileName )
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

void EnumTypeLib( LPTYPELIB pITypeLib )
{
	UINT tiCount = pITypeLib->GetTypeInfoCount();
	
	for ( UINT i = 0; i < tiCount; i++ )
	{
		LPTYPEINFO pITypeInfo;

		HRESULT hr = pITypeLib->GetTypeInfo( i, &pITypeInfo );

		if ( S_OK == hr )
		{
			DisplayTypeInfo( pITypeInfo );
							
			pITypeInfo->Release();
		}
	}
}

void DisplayTypeInfo( LPTYPEINFO pITypeInfo )
{
	HRESULT hr;
	
	BSTR pszTypeInfoName;
	hr = pITypeInfo->GetDocumentation(MEMBERID_NIL, &pszTypeInfoName, 0, 0, 0);
	if ( S_OK != hr )
		return;
		
	TYPEATTR * pTypeAttr;
	hr = pITypeInfo->GetTypeAttr( &pTypeAttr );
	if ( S_OK != hr )
	{
		SysFreeString( pszTypeInfoName );
		return;
	}
	
	_tprintf( _T("%ls - %s\n"), pszTypeInfoName,
				GetTypeKindName(pTypeAttr->typekind) );

	EnumTypeInfoMembers( pITypeInfo, pTypeAttr );

	_tprintf( _T("\n") );
		
	SysFreeString( pszTypeInfoName );
	
	pITypeInfo->ReleaseTypeAttr( pTypeAttr );
}

void EnumTypeInfoMembers( LPTYPEINFO pITypeInfo, LPTYPEATTR pTypeAttr  )
{
	if ( pTypeAttr->cFuncs )
	{
		_tprintf( _T("  Functions:\n") );
			
		for ( unsigned i = 0; i < pTypeAttr->cFuncs; i++ )
		{
			FUNCDESC * pFuncDesc;
			
			pITypeInfo->GetFuncDesc( i, &pFuncDesc );
			
			BSTR pszFuncName;		
			pITypeInfo->GetDocumentation(pFuncDesc->memid, &pszFuncName,0,0,0);

			_tprintf( _T("    %-32ls"), pszFuncName );

			_tprintf( _T(" (%ls)\n"), GetInvokeKindName(pFuncDesc->invkind) );
				
			pITypeInfo->ReleaseFuncDesc( pFuncDesc );						
			SysFreeString( pszFuncName );
		}
	}

	if ( pTypeAttr->cVars )
	{
		_tprintf( _T("  Variables:\n") );
			
		for ( unsigned i = 0; i < pTypeAttr->cVars; i++ )
		{
			VARDESC * pVarDesc;
			
			pITypeInfo->GetVarDesc( i, &pVarDesc );
			
			BSTR pszVarName;		
			pITypeInfo->GetDocumentation(pVarDesc->memid, &pszVarName,0,0,0);

			_tprintf( _T("    %ls\n"), pszVarName );

			pITypeInfo->ReleaseVarDesc( pVarDesc );						
			SysFreeString( pszVarName );
		}
	}

}

#define CASE_STRING( x ) case x: s = _T(#x); break;

LPCTSTR GetTypeKindName( TYPEKIND typekind )
{
	LPTSTR s = _T("<unknown>");
	
	switch( typekind )
	{
    	CASE_STRING( TKIND_ENUM )
    	CASE_STRING( TKIND_RECORD )
    	CASE_STRING( TKIND_MODULE )
    	CASE_STRING( TKIND_INTERFACE )
    	CASE_STRING( TKIND_DISPATCH )
    	CASE_STRING( TKIND_COCLASS )
    	CASE_STRING( TKIND_ALIAS )
    	CASE_STRING( TKIND_UNION )
	}
	
	return s;
}

LPCTSTR GetInvokeKindName( INVOKEKIND invkind )
{
	LPTSTR s = _T("<unknown>");
	
	switch( invkind )
	{
    	CASE_STRING( INVOKE_FUNC )
		CASE_STRING( INVOKE_PROPERTYGET )
    	CASE_STRING( INVOKE_PROPERTYPUT )
    	CASE_STRING( INVOKE_PROPERTYPUTREF )
	}	

	return s;
}

