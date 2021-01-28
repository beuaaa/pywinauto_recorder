!include "MUI2.nsh"
!define MUI_ICON "IconPyRec.ico"

!define ZIP2EXE_COMPRESSOR_SOLID
!define ZIP2EXE_COMPRESSOR_LZMA
!define ZIP2EXE_INSTALLDIR "C:\Program Files\Pywinauto recorder"
!define ZIP2EXE_NAME "Pywinauto recorder"
!define ZIP2EXE_OUTFILE "Pywinauto recorder installer.exe"

!include "${NSISDIR}\Contrib\zip2exe\Base.nsh"
!include "${NSISDIR}\Contrib\zip2exe\Modern.nsh"

!insertmacro SECTION_BEGIN
# specify file to go in output path
File /r pywinauto_recorder.dist\*.*
!insertmacro SECTION_END

# default section start
Section
	# define output path
	SetOutPath $INSTDIR
	CreateDirectory "$INSTDIR\Pywinauto recorder"
	CreateShortCut "$INSTDIR\Pywinauto recorder\Pywinauto recorder.lnk" "$INSTDIR\pywinauto_recorder.exe"


	# define uninstaller
	WriteUninstaller $INSTDIR\uninstall_Pywinauto_recorder.exe
	CreateShortCut "$INSTDIR\Pywinauto recorder\Uninstall.lnk" "$INSTDIR\uninstall_Pywinauto_recorder.exe"

	CreateShortcut "$SMPROGRAMS\Pywinauto recorder.lnk" "$INSTDIR\Pywinauto recorder"
	 
	#-------
	# default section end
SectionEnd


# create a section to define what the uninstaller does.
# the section will always be named "Uninstall"
Section "Uninstall"
	# Delete the directory
	MessageBox MB_OKCANCEL|MB_ICONEXCLAMATION "Are you sure you want to delete $INSTDIR" /SD IDCANCEL IDOK OK IDCANCEL CANCEL
		OK:
			RMDir /r $INSTDIR
			Delete "$DESKTOP\Pywinauto recorder.lnk"
			goto end_uninstall
		CANCEL:
			Abort
	end_uninstall:
SectionEnd


Section "Desktop Shortcut" SectionX
	# Create shortcut on desktop
	MessageBox MB_OKCANCEL "Do you want to add Pywinauto recorder shortcut on desktop?" /SD IDCANCEL IDOK OK IDCANCEL CANCEL
		OK:
			SetShellVarContext current
			CreateShortCut "$DESKTOP\Pywinauto recorder.lnk" "$INSTDIR\pywinauto_recorder.exe"
			goto end_create_shortcut_on_desktop
		CANCEL:
			Abort
	end_create_shortcut_on_desktop:
SectionEnd 


