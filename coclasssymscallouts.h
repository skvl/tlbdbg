#ifdef __cplusplus
extern "C" {
#endif

BOOL __stdcall CoClassSymsBeginSymbolCallouts( PSTR pszExecutable );

BOOL __stdcall  CoClassSymsAddSymbol(
		unsigned short section,
		unsigned long offset,
		PSTR pszSymbolName );

BOOL __stdcall CoClassSymsSymbolsFinished( void );


typedef BOOL (__stdcall * PFNCCSCALLOUTBEGIN)( PSTR pszExecutable );
typedef BOOL (__stdcall * PFNCCSADDSYMBOL)(unsigned short, unsigned long,PSTR);
typedef BOOL (__stdcall * PFNCCSFINISHED)(void);

#ifdef __cplusplus
}
#endif
