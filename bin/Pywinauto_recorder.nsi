!include "MUI2.nsh"
!define MUI_ICON "..\pywinauto_recorder\Icons\IconPyRec.ico"

!define ZIP2EXE_COMPRESSOR_SOLID
!define ZIP2EXE_COMPRESSOR_LZMA
!define ZIP2EXE_INSTALLDIR "C:\Program Files\Pywinauto recorder"
!define ZIP2EXE_NAME "Pywinauto recorder"
!define ZIP2EXE_OUTFILE "Pywinauto_recorder_installer.exe"

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
	
	SetShellVarContext all
	
	# create Pywinauto recorder folder entry in Start menu
	CreateDirectory "$SMPROGRAMS\Pywinauto recorder"
	
	# create Pywinauto recorder\Pywinauto recorder.lnk entry in Start menu
	CreateShortCut "$SMPROGRAMS\Pywinauto recorder\Pywinauto recorder.lnk" "$INSTDIR\pywinauto_recorder.exe"
	
	# define uninstaller
	WriteUninstaller $INSTDIR\uninstall_Pywinauto_recorder.exe

	# create Pywinauto recorder\ywinauto recorder.lnk entry in Start menu
	CreateShortCut "$SMPROGRAMS\Pywinauto recorder\Uninstall Pywinauto recorder.lnk" "$INSTDIR\uninstall_Pywinauto_recorder.exe"

	#-------
	# default section end
SectionEnd


# create a section to define what the uninstaller does.
# the section will always be named "Uninstall"
Section "Uninstall"
	# Delete the directory
	MessageBox MB_OKCANCEL|MB_ICONEXCLAMATION "Are you sure you want to delete $INSTDIR" /SD IDCANCEL IDOK OK IDCANCEL CANCEL
		OK:
			SetShellVarContext all
			RMDir /r $INSTDIR
			Delete "$DESKTOP\Pywinauto recorder.lnk"
			RMDir /r "$SMPROGRAMS\Pywinauto recorder"
			goto end_uninstall
		CANCEL:
			Abort
	end_uninstall:
SectionEnd


Section "Desktop Shortcut" SectionX
	# Create shortcut on desktop
	MessageBox MB_OKCANCEL "Do you want to add Pywinauto recorder shortcut on desktop?" /SD IDCANCEL IDOK OK IDCANCEL CANCEL
		OK:
			SetShellVarContext all
			CreateShortCut "$DESKTOP\Pywinauto recorder.lnk" "$INSTDIR\pywinauto_recorder.exe"
			goto end_create_shortcut_on_desktop
		CANCEL:
			goto end_create_shortcut_on_desktop
	end_create_shortcut_on_desktop:
SectionEnd 


Section "Run clear_comtypes_cache.exe"
	Exec '"$INSTDIR\clear_comtypes_cache.exe" -y'
SectionEnd


Section "Run pywinauto_recorder.exe"
	# Run Pywinauto_recorde_exe
	MessageBox MB_OKCANCEL "Do you want to run Pywinauto recorder now?" /SD IDCANCEL IDOK OK IDCANCEL CANCEL
		OK:
			Exec '"$INSTDIR\pywinauto_recorder.exe"'
			goto end_run_pywinauto_recorder
		CANCEL:
			goto end_run_pywinauto_recorder
	end_run_pywinauto_recorder:
SectionEnd 
