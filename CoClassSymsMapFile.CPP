#include <windows.h>
#include <imagehlp.h>
#include <stdio.h>
#include <stdlib.h>

//=========================== Global Variables ===============================

LOADED_IMAGE g_loadedImage;
FILE * g_pMapFile;

//============================================================================

BOOL __stdcall CoClassSymsBeginSymbolCallouts( LPCSTR pszExecutable )
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

	fprintf( g_pMapFile,
			" Start         Length     Name                   Class\n" );

	PIMAGE_SECTION_HEADER pSectHdr = g_loadedImage.Sections;
			
	for ( 	unsigned i=1;
			i <= g_loadedImage.NumberOfSections;
			i++, pSectHdr++ )
	{
		fprintf( 	g_pMapFile,
					" %04X:00000000 %08XH %-23.8hs %s\n",
					i, pSectHdr->Misc.VirtualSize, pSectHdr->Name,
					pSectHdr->Characteristics & IMAGE_SCN_CNT_CODE
						? "CODE" : "DATA" );
	}

	fprintf( g_pMapFile, 
		"\n  Address         Publics by Value              Rva+Base\n\n");	

	return TRUE;
}

BOOL __stdcall CoClassSymsAddSymbol(
		unsigned short section,
		unsigned long offset,
		PSTR pszSymbolName )
{
	if ( !g_pMapFile )
		return FALSE;

	fprintf( g_pMapFile, " %04X:%08X       %-32s\n",
			 section, offset, pszSymbolName );
				
	return true;
}
		
BOOL __stdcall CoClassSymsSymbolsFinished( void )
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
		
		fprintf( g_pMapFile, "\n entry point at        %04X:%08X\n",
			 	 section, offset );
	}
	
	fclose( g_pMapFile );

	UnMapAndLoad( &g_loadedImage );		// Undo the MapAndLoad call
	
	return TRUE;
}


