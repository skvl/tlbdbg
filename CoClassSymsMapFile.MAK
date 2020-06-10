PROJ = CoClassSymsMapFile
OBJS = $(PROJ).OBJ

CFLAGS = $(CFLAGS) /W3 /GF /DWIN32_LEAN_AND_MEAN 

LFLAGS = /DLL /SUBSYSTEM:console

!if "$(DEBUG)" == "1"
CFLAGS = $(CFLAGS) /YX /Zi /D_DEBUG /Fd"$(PROJ).PDB" /Fp"$(PROJ).PCH"
LFLAGS = $(LFLAGS) /DEBUG /DEBUGTYPE:CV
!else
CFLAGS = $(CFLAGS) /O1 /DNDEBUG
!endif

LIBS = IMAGEHLP.LIB

EXPORTS = \
	/EXPORT:CoClassSymsBeginSymbolCallouts \
	/EXPORT:CoClassSymsAddSymbol \
	/EXPORT:CoClassSymsSymbolsFinished \

$(PROJ).DLL: $(OBJS)
    echo >NUL @<<$(PROJ).CRF
$(LFLAGS) $(OBJS) -OUT:$(PROJ).DLL $(LIBS) $(EXPORTS)
<<
    link @$(PROJ).CRF

.cpp.obj::
    CL $(CFLAGS) /c $<
