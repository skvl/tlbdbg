PROJ = CoClassSyms
OBJS = $(PROJ).OBJ

CFLAGS = $(CFLAGS) /DUNICODE=1 /D_UNICODE=1 /W3 /GF /DWIN32_LEAN_AND_MEAN 

LFLAGS = /SUBSYSTEM:console /ENTRY:wmainCRTStartup

!if "$(DEBUG)" == "1"
CFLAGS = $(CFLAGS) /YX /Zi /D_DEBUG /Fd"$(PROJ).PDB" /Fp"$(PROJ).PCH" /MDd
LFLAGS = $(LFLAGS) /DEBUG /DEBUGTYPE:CV
!else
CFLAGS = $(CFLAGS) /O1 /DNDEBUG /MD
!endif

LIBS = USER32.LIB OLE32.LIB OLEAUT32.LIB IMAGEHLP.LIB

$(PROJ).EXE: $(OBJS)
    echo >NUL @<<$(PROJ).CRF
$(LFLAGS) $(OBJS) -OUT:$(PROJ).EXE $(LIBS)
<<
    link @$(PROJ).CRF

.cpp.obj::
    CL $(CFLAGS) /c $<
