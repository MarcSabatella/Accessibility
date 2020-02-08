; Install wizard for the MuseScore3 scripts
!define BASENAME "jfw_musescore3"
; Outfile should be constant for all versions, as it is used to name the uninstaller.
!define OUTFILE "${BASENAME}.exe"
!define LOGFILE "${BASENAME}_install.log"
; Name should be constant for all versions, so uninstallers can be found.
!define NAME "JAWS Scripts For MuseScore3"
; Release should indicate to the user which version is being installed.
!define RELEASE "${NAME}"
OutFile "${OUTFILE}"
NAME "${NAME}"
CAPTION "${RELEASE} Setup"
!define ATTRIBUTION ""
!define LICENSE "license_gpl2.txt"


; $R1 is a JAWS build like 18.0 or 2018.
; $R0 is a registry enumeration counter.
; Preserve any registers you modify in here!
!macro JAWSVerCheck
		${If} $R1 < 18.0
			${DetailPrint} "Skipping; incompatible with these scripts."
			IntOp $R0 $R0 + 1
			${Continue}
		${EndIf}
!macroEnd

; This script requires LogicLib and an NSIS version new enough to contain it.
; http://nsis.SourceForge.net
;
; Written by Victor Tsaran, updated by Scott McCormack, update by Jonathan Avila 4-2-07
; Restructured for multilanguage support (and other things) by Doug Lee, 02/2008, 03/2010.

SetCompressor /solid lzma

!include "MUI2.nsh"
!include "InstallOptions.nsh"
!include LogicLib.nsh
!include X64.nsh

;--------------------------------
;Modern UI Configuration

; This line keeps JAWS from saying "Nulsoft Installation System" repeatedly on various windows.
BrandingText /TrimCenter " "
XPStyle on
CRCCheck On
ShowInstDetails hide
ShowUninstDetails hide
SetOverwrite On
SetDateSave on
RequestExecutionLevel user

!define MUI_WELCOMEPAGE_TITLE "Welcome to the installation of ${NAME}."
!define MUI_WELCOMEPAGE_TEXT "${ATTRIBUTION}This wizard will guide you through the installation of ${NAME} \
for any compatible JAWS versions you have installed on your computer.$\n"
!define MUI_FINISHPAGE_TEXT_LARGE
; Dynamic text for the Finish page.
; TODO: This will complicate multilanguage installer creation.
var FinishText
!define MUI_FINISHPAGE_TEXT $FinishText
!define MUI_FINISHPAGE_NOREBOOTSUPPORT
!define MUI_WELCOMEFINISHPAGE_BITMAP_NOSTRETCH
!define MUI_ABORTWARNING

!define MUI_UNINSTALLER
; TODO: This is not the right define but I don't know the right one.  [DGL]
!define MUI_UNWELCOMEPAGE_TEXT "${ATTRIBUTION}This wizard will guide you through the uninstallation of ${NAME} \
for any JAWS versions containing these scripts.$\n"

; Custom page macros
!InsertMacro MUI_PAGE_WELCOME
!IfDef LICENSE
!InsertMacro MUI_PAGE_LICENSE "${LICENSE}"
!EndIf
Page custom CheckJAWSVersions
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!InsertMacro MUI_UNPAGE_WELCOME
UninstPage custom un.DisplayUninstallOptions
!InsertMacro MUI_UNPAGE_INSTFILES
!InsertMacro MUI_UNPAGE_FINISH

!define MUI_HEADERBITMAP "${NSISDIR}\Contrib\Icons\modern-header.bmp"

; Several user-defined constants
!define JAWSApp "jfw.exe"
!define UnInstaller "uninst_${OUTFILE}"
!define INSTINIFILE "install.ini"
!define UNINSTINIFILE "uninstall.ini"
!define INSTALLSINIFILE "installs.ini"
!define INSTALLSINISECT "Installs"

; Log text used when the Details list is not available for dumping at exit time.
var DetailLines
; Number of lines represented.
var DetailLineCount
; Set by .onInit to the current-user and all-user paths to JAWS files.
var userRoot
var AllUsersRoot
; The CheckJAWSVersions page leaves a |-separated list of JAWS folder indicators here,
; and the Install page uses it to do installations.
; Format: ver/lng|ver/lng..., e.g., 18.0/enu|18.0/esn|2018/enu.
var JAWSFolderList
; The JAWS language (e.g., enu).
var JAWSLang
; The JAWS version major number, e.g., "18 or 2018"
; Used when checking for installer minimum JAWS version requirement.
var JAWSVersion
; The JAWS build number used in paths, e.g., "18.0 or 2018 (same as above 2018 and later)"
var JAWSBuild
; The full build number.
var JAWSFullBuild
; List of JAWS install directories.
; This value is written into an ini file in the uninstaller's folder.
var JAWSInstallDirs
; Paths to parts of a JAWS version.
var UserSettingsPath
var AllUserSettingsPath
; For the Install section, counts of installs to be made and successfully made.
var InstCount
var InstSuccessCount

; InstallDir would normally be set here but it is managed by code.
; Same for InstallDirRegKey.
; The uninstaller goes where the [un.]getUninstallerPath function says.

;--------------------------------
;Languages
!insertmacro MUI_LANGUAGE "English"

;---
; Language strings
LangString TEXT_IO_PAGETITLE ${LANG_ENGLISH} "Checking installed versions of JAWS..."

;--------------------------------
;Reserve Files
ReserveFile '${NSISDIR}\Plugins\x86-unicode\InstallOptions.dll'
ReserveFile "${INSTINIFILE}"
ReserveFile "${UNINSTINIFILE}"

;--------------------------------
; Utility macros and defines.

; This makes it possible to say things like ${ErrMsg} "!Error: That didn't work"
; A leading "!" causes a message box; otherwise it's just a detailPrint.
!macro _ErrMsg _T
	push "${_T}"
	!ifdef __UNINSTALL__
	call un.errmsg
	!else
	call errmsg
	!endIf
!macroEnd
!define ErrMsg "!InsertMacro _ErrMsg"

; This makes ${DetailPrint} "blah..." do detailPrint "blah..." and save the text in DetailLines.
!macro _DetailPrint _T
	detailPrint "${_T}"
	StrCpy $DetailLines "$DetailLines${_T}$\n"
	IntOp $DetailLineCount $DetailLineCount + 1
!macroEnd
!define DetailPrint "!InsertMacro _DetailPrint"

;--------------------------------
; Initial code run before the install section.

Function .onInit
StrCpy $DetailLines ""
IntOp $DetailLineCount $DetailLineCount * 0
; Check if we are already installed and do other initial setup.
; Registers:
;	$1 = installed JAWS versions according to the uninstaller-folder ini file.
;		This would indicate a prior run of this installer.
;		Each version is represented by its full JAWS user folder/language path.
;		See $JAWSInstallDirs in the Install section, which creates this.
	${If} ${RunningX64}
		SetRegView 64
	${EndIf}
	StrCpy $FinishText "All requested installations were successful."
	!IfDef DBG
	SetDetailsView Show
	SetAutoClose False
	WriteUninstaller "$EXEDIR\${UnInstaller}"
	!EndIf
	call getUninstallerPath
	; Leave path on stack.
	push "${UnInstaller}"
	; TODO: Make this "/s" when the installer can handle it?
	; Would probably need to show a wait indicator of some sort too.
	push ""
	call uninstallIfNeeded
	; If we survived this far, there are no longer any scripts to remove.
	!insertmacro INSTALLOPTIONS_EXTRACT "${INSTINIFILE}"
	; TODO: Is $APPDATA enough when JAWS is not on C:?
	StrCpy $UserRoot "$APPDATA\Freedom Scientific\JAWS"
	SetShellVarContext all
	StrCpy $AllUsersRoot "$APPDATA\Freedom Scientific\JAWS"
	SetShellVarContext current
	pop $1
FunctionEnd

Function uninstallIfNeeded
var /GLOBAL uin_path
var /GLOBAL uin_exename
var /GLOBAL uin_parms
pop $uin_parms
	pop $uin_exename
	pop $uin_path
	${IfNot} ${FileExists} "$uin_path\$uin_exename"
		return
	${EndIf}
	${DetailPrint} "Found uninstaller $uin_path\$uin_exename"
	MessageBox MB_YESNO "A version of these scripts is already installed on your computer.$\n\
	The current scripts must be removed before this installation can continue.$\n\
	Would you like to run the uninstaller now?$\n" /SD IDYES IDYES +1 IDNO CancelInstall
	; Copy to temp, then run with the original path specified.
	; This allows us to wait for the uninstaller to exit.
	; Otherwise, it copies itself to temp, calls itself, then returns immediately.
	ClearErrors
	CopyFiles /SILENT "$uin_path\$uin_exename" "$TEMP"
	${If} ${Errors}
		${ErrMsg} "!The uninstaller could not be run.  This installation cannot continue."
		Quit
	${EndIf}
	push $1
	ExecWait '"$TEMP\$uin_exename" $uin_parms _?=$uin_path' $1
	${If} ${Errors}
		${ErrMsg} "!The uninstaller failed to run.  Installation cannot continue."
		Delete "$TEMP\$uin_exename"
		Quit
	${EndIf}
	Delete "$TEMP\$uin_exename"
	${If} $1 = 0
		; Make sure the uninstaller cleaned up everything.
		${If} ${FileExists} "$uin_path\$uin_exename"
			${ErrMsg} "!The uninstaller did not run or did not remove all copies of the scripts.$\n"
			Quit
		${EndIf}
		${ErrMsg} "!Uninstallation was successful$\n"
	${Else}
CancelInstall:
		${ErrMsg} "!The installation of ${NAME} could not continue...$\n"
		Quit
	${EndIf}
	pop $1
FunctionEnd

Function CheckJAWSVersions
; Page that displays before the Install section.
; Register usage (registers NOT preserved):
;	$R0: Enumeration counter.
;	$R1: JAWS version key from registry, and appname from registry.
;	$R2: JAWS path for a JAWS version, from registry.
;	$R5: List of JAWS paths.
;	$R8: Scratch register for flags etc.
	IntOp $R0 $R0 * 0
	StrCpy $R5 ""
	${DetailPrint} "Checking for JAWS versions"
	${Do}
		; Enumerate installed versions of JAWS
		EnumRegKey $R1 HKLM "Software\Freedom Scientific\JAWS" $R0
		${If} $R1 == ""
			${DetailPrint} "No more JAWS installation paths."
			${Break}
		${EndIf}
		${DetailPrint} "Key $R1"
		${If} $R1 < 1
			${DetailPrint} "Skipping; not a JAWS version number."
			IntOp $R0 $R0 + 1
			${Continue}
		${EndIf}
		!IfMacroDef JAWSVerCheck
		!InsertMacro JAWSVerCheck
		!EndIf
		; Find out paths for JAWS installations
		ReadRegStr $R2 HKLM "Software\Freedom Scientific\JAWS\$R1" "Target"
		${If} $R2 == ""
			${DetailPrint} "Skipping; no installation path (Target) found."
			IntOp $R0 $R0 + 1
			${Continue}
		${EndIf}

		; Make sure $R2 is a valid JAWS version path on disk.
		Push $R2
		Call IsValidJAWSPath
		Pop $R8
		${If} $R8 = 0
			${DetailPrint} "Skipping; JAWS installation path does not contain JAWS."
			IntOp $R0 $R0 + 1
			${Continue}
		${EndIf}
		StrCpy $R5 "$R5|$R2"
		IntOp $R0 $R0 + 1
	${Loop}  ; enumerating registry keys

	; If no valid JAWS registry keys found, check if it was registered in the list of applications
	${If} $R5 == ""
		; Check for existence of ${JAWSApp} among registered applications
		${DetailPrint} "Looking for JAWS among registered applications."
		IntOp $R0 $R0 * 0
		; Enumerate all registered applications and see if ${JAWSApp} is there
		${Do}
			EnumRegKey $R1 HKLM "Software\Microsoft\Windows\CurrentVersion\App Paths" $R0
			${If} $R1 == ""
				${DetailPrint} "${JAWSApp} not listed."
				goto Failed
			${EndIf}
			${If} $R1 == ${JAWSApp}
				${Break}
			${EndIf}
			IntOp $R0 $R0 + 1
		${Loop}
		; If we got this far, we found a JAWS app installation.
		${DetailPrint} "Found ${JAWSApp}."
		ReadRegStr $R2 HKLM "Software\Microsoft\Windows\CurrentVersion\App Paths\$R1" "Path"
		${If} $R2 == ""
			${DetailPrint} "but the path is empty or missing!"
			goto Failed
		${EndIf}

		; Otherwise, check if ${JAWSApp} is actually there.
		Push $R2
		Call IsValidJAWSPath
		Pop $R8
		${If} $R8 = 0
			${DetailPrint} "but the path ($R2) does not contain ${JAWSApp}!"
			goto Failed
		${EndIf}

		${DetailPrint} "JAWS found at $R2."
		StrCpy $R5 "$R5|$R2"
	${EndIf}

	; Remove leading "|" if present.
	StrCpy $R5 $R5 "" 1
	${If} $R5 == ""
		${DetailPrint} "No JAWS app paths found."
		goto failed
	${EndIf}

	; We now have a list of JAWS program paths,
	; but we want listItems for each language in each version.
	push $R5
	call JAWSProgramPathsToListItems
	pop $R5
	; and if there's only one language we needn't show the language codes.
	push $R5
	call tryToSetCommonLanguage
	pop $R5
	; Now either $JAWSLang is empty and $R5 looks like 18.0/enu|18.0/ita|2018/enu,
	; or $JAWSLang is a language code and $R5 looks like 17.0|18.0|2018.

	; For valid versions of JAWS fill up a value list in .ini configuration file
	!InsertMacro INSTALLOPTIONS_WRITE "${INSTINIFILE}" "Field 2" "ListItems" "$R5"
	; and auto-select them all while we're at it.
	!InsertMacro INSTALLOPTIONS_WRITE "${INSTINIFILE}" "Field 2" "state" "$R5"

	; Display the list of JAWS installations
	!InsertMacro INSTALLOPTIONS_DISPLAY_RETURN "${INSTINIFILE}"
		Pop $R8 ; Pop the return value from the stack
	StrCmp $R8 "cancel" CancelInstall +1
	!InsertMacro INSTALLOPTIONS_READ $R5 "${INSTINIFILE}" "Field 2" "State"
	!InsertMacro INSTALLOPTIONS_WRITE "${INSTINIFILE}" "Field 2" "ListItems" "" ; clean up the ListItems value

	; Save the list of JAWS versions/languages selected.
	StrCpy $JAWSFolderList $R5
	GoTo JAWS_OK

Failed:
	MessageBox MB_OK "No compatible JAWS version appears to be installed on this computer.$\n\
		Please install or reinstall JAWS and run this installation again. The installer will now quit.$\n" /SD IDOK
	call Quit

CancelInstall:
	MessageBox MB_OK "You have chosen to cancel this installation. The installer will now quit.$\n" /SD IDOK
	Quit

JAWS_OK:
FunctionEnd

Function JAWSProgramPathsToListItems
; Convert a list of JAWS program paths to their corresponding listItems.
; For multilanguage JAWS installations, more list items can come out of this than came in.
; JAWS version filtering is also implemented here.
; Usage: push progPathList, call, pop listItems.
; ListItem format: ver/lng, e.g., 18.0/enu or 2018/enu.
	push $R1
	exch
	exch $1
	push $2
	push $3
	push $4
	StrCpy $R1 ""
	${DoWhile} $1 != ""
		; Next JAWS program path to $2.
		push $1
		push "|"
		call SplitFirstStrPart
		pop $2
		pop $1
		; Set $JAWSVersion, $JAWSBuild, and $JAWSFullBuild for this version.
		push $2
		call getJAWSExeVersion
		${DetailPrint} "JAWS $JAWSBuild:"
		; Outlaw JAWS versions that are not fit for these scripts.
		${If} $JAWSVersion < 18.0
			${DetailPrint} "These scripts will not work with this version. Skipping this JAWS version."
			${Continue}
		${EndIf}
		; Set up My Settings and Shared Settings paths for this JAWS version,
		; except for the final language folder part.
		StrCpy $UserSettingsPath "$userRoot\$JAWSBuild\Settings"
		StrCpy $AllUserSettingsPath "$allUsersRoot\$JAWSBuild\Settings"
		; Get a list of the language folders found in this JAWS version.
		push $AllUserSettingsPath
		call GetJAWSLanguages
		pop $3
		${DetailPrint} "Languages found: $3."
		${DoWhile} $3 != ""
			; Get the next language folder name.
			push $3
			push "|"
			Call SplitFirstStrPart
			pop $4
			pop $3
			; Add a listItem.
			; TODO: Could consider including indications here of which folders
			; already contain these scripts or possibly conflicting ones.
			; E.g., "18.0/enu (already installed)|2018/enu (possible conflict)"
			StrCpy $R1 "$R1|$JAWSBuild/$4"
		${Loop}
	${Loop}
	StrCpy $R1 $R1 "" 1
	pop $4
	pop $3
	pop $2
	pop $1
	exch $R1
FunctionEnd

Function tryToSetCommonLanguage
; Given a |-delimited list of the form ver/lng|ver/lng...,
; either set $JAWSLang to the language common to all items in the list
; and remove them, leaving the list in the form ver|ver...,
; or set $JAWSLang to be empty and leave the list alone, if not all languages are the same.
var /GLOBAL scl_list
	pop $scl_list
	push $4
	push $1
	push $2
	push $3
	StrCpy $1 $scl_list
	StrCpy $JAWSLang ""
	StrCpy $4 ""
	${DoWhile} $1 != ""
		push $1
		push "|"
		call SplitFirstStrPart
		pop $2
		pop $1
		push $2
		push "/"
		call SplitFirstStrPart
		pop $2
		pop $3
		${If} $JAWSLang == ""
			; First time we got a language; just record it.
			StrCpy $JAWSLang $3
			; and start the $4 list of just JAWS versions.
			StrCpy $4 $2
		${ElseIf} $3 == $JAWSLang
			; Another instance of the language we already saw.
			; Build a list in $4 of just the JAWS versions without languages.
			StrCpy $4 "$4|$2"
			${Continue}
		${Else}
			; Different language found, so there is no common one.
			StrCpy $JAWSLang ""
			StrCpy $4 $scl_list
			${Break}
		${EndIf}
	${Loop}
	; $4 is now either the original list or the languageless one.
	; Note that it is not necessary to remove a leading "|"
	; because we avoided putting it there in the first place.
	; $JAWSLang is now the common language or blank if there isn't one.
	pop $3
	pop $2
	pop $1
	exch $4
FunctionEnd

Function GetJAWSEXEVersion
; Sets $JAWSBuild, $JAWSFullBuild, and $JAWSVersion from the JAWS path on the stack.
; The JAWS path should be the JAWS program folder path.
; Registers (preserved):
;	$R0 and $R1: High and low words of GetDLLVersion return value.
;	$R2-5: Parts of the version number.
;	$0: Human-readable JAWS version ("17.0.2417"). (2018+ build numbers may be odd.)
;	$8: Path to JAWS program folder (passed on stack).
	exch $8
	push $R0
	push $R1
	push $R2
	push $R3
	push $R4
	push $R5
	push $0
	GetDllVersion "$8\${JAWSAPP}" $R0 $R1
	; Code from NSIS docs on how to translate that to version number parts.
	IntOp $R2 $R0 / 0x00010000
	IntOp $R3 $R0 & 0x0000FFFF
	IntOp $R4 $R1 / 0x00010000
	IntOp $R5 $R1 & 0x0000FFFF
	${If} $R2 > 18
		; Adjustment from FS/VFO internal version number, e.g., 19..., to public one, e.g., 2018...
		; ToDo: Warning: Build numbers here may well not equal published ones anyway.
		IntOp $R2 $R2 + 1999
		StrCpy $JAWSBuild $R2
	${Else}
		StrCpy $JAWSBuild $R2.$R3
	${EndIf}
	; This is here for show but not actually used.
	StrCpy $0 "$R2.$R3.$R4.$R5"
	StrCpy $JAWSFullBuild $R2.$R3.$R4
	IntOp $JAWSVersion 1 * $R2
	pop $0
	pop $R5
	pop $R4
	pop $R3
	pop $R2
	pop $R1
	pop $R0
	exch $8
FunctionEnd

Function IsValidJAWSPath
; Check the validity of the JAWS directory.
	; Pick the directory name from the stack
	Pop $R6
	IfFileExists "$R6\${JAWSApp}" +1 Failed
	Push 1 ; Return the success flag
	Return

Failed:
	Push 0 ; Set the error flag
	Return
FunctionEnd

;--------------------------------
; Install-time sections and code

Section "install"
; Runs after the CheckJAWSVersions page, which fills the $JAWSFolderList variable.
; Empties $JAWSFolderList in the course of installing things.
; Register usage:
;	$0 = full Shared Settings/lang path during an installation.
;	$1 = Next JAWS folder indicator (from JAWSFolderList).
	${DetailPrint} "Install into JAWS folders $JAWSFolderList."
	${If} $JAWSFolderList == ""
		MessageBox MB_OK "No JAWS folders were selected.  The installer will now quit.$\n" /SD IDOK
		Quit
	${EndIf}

	IntOp $InstCount $InstCount * 0
	IntOp $InstSuccessCount $InstSuccessCount * 0

	; Go through JAWS folders installing and compiling scripts in each.
	${DoWhile} $JAWSFolderList != ""
		ClearErrors

		; Get the next JAWS folder indicator to consider, and remove it from the list.
		Push "$JAWSFolderList"
		push "|"
		Call SplitFirstStrPart
		Pop $1
		Pop $JAWSFolderList
		${DetailPrint} "$1:"
		; $1 looks like 2018/enu or just 2018 if there is a common language.
		push $1
		push "/"
		call SplitFirstStrPart
		pop $JAWSBuild
		pop $R8
		${If} $R8 != ""
			StrCpy $JAWSLang $R8
		${EndIf}

		; Set up My Settings and Shared Settings paths for this JAWS version.
		StrCpy $UserSettingsPath "$userRoot\$JAWSBuild\Settings\$JAWSLang"
		StrCpy $AllUserSettingsPath "$allUsersRoot\$JAWSBuild\Settings\$JAWSLang"
		${DetailPrint} "User folder: $UserSettingsPath"
		${DetailPrint} "Shared folder: $AllUserSettingsPath"
		StrCpy $INSTDIR "$UserSettingsPath"
		SetOutPath $INSTDIR
		StrCpy $0 "$AllUserSettingsPath"
		IntOp $InstCount $InstCount + 1

		!IfNDef DBG
		; Status bar shows individual files but list box won't.
		${DetailPrint} "Installing files."
		setDetailsPrint textonly
		!EndIf

		; Unpack necessary files from the installer
		file "musescore3.jkm"
		file "musescore3.jsb"
		file "musescore3.jsd"
		file "musescore3.jsm"
		file "musescore3.jss"


		!IfNDef DBG
		; List box again shows everything.
		setDetailsPrint both
		!EndIf

		${DetailPrint} "Updating JAWS user folder files."

		; Collect all JAWS user folder/language directory paths.
		; These go into the uninstaller-folder ini file for the uninstaller to retrieve.
		StrCpy $JAWSInstallDirs "$JAWSInstallDirs|$INSTDIR"
		IntOp $InstSuccessCount $InstSuccessCount + 1
	${Loop}  ; Next JAWS version

	; ChopLeft the beginning | from $InstallDirs
	StrCpy $JAWSInstallDirs $JAWSInstallDirs "" 1
	${DetailPrint} "JAWSInstallDirs is $JAWSInstallDirs."

	${If} $InstSuccessCount = 0
		StrCpy $FinishText "All installations failed!"
		${ErrMsg} "$FinishText"
		MessageBox MB_OK "All requested script installations have failed." /SD IDOK
		; Exit immediately with logging.
		call Quit
		goto instExit
	${ElseIf} $InstSuccessCount <> $InstCount
		StrCpy $FinishText "Only $InstSuccessCount of $InstCount installations succeeded."
		${ErrMsg} "!$FinishText"
		; But since some succeeded, we go on to finish.
	${EndIf}

	; Change the output directory to our real installation directory
	call getUninstallerPath
	pop $INSTDIR
	SetOutPath $INSTDIR

	; Register the installations made.
	writeIniStr "$INSTDIR\${INSTALLSINIFILE}" ${INSTALLSINISECT} "InstalledForJFWVersions" "$JAWSInstallDirs"
	; The uninstaller will read that to find things to uninstall.
	WriteUninstaller "${UnInstaller}"
instExit:
SectionEnd

Section "Start Menu Shortcuts"
	${If} $InstSuccessCount = 0
		${DetailPrint} "Installer and Start menu entries not created because no installations succeeded."
	${Else}
		CreateDirectory "$SMPROGRAMS\${NAME}"
		CreateShortCut "$SMPROGRAMS\${NAME}\Uninstall ${NAME}.lnk" "$INSTDIR\${UnInstaller}" "" "$INSTDIR\${UnInstaller}" 0
	${EndIf}
	; Log and exit if anything failed.
	${If} $InstSuccessCount <> $InstCount
		call Quit
	${Else}
		; In case it exists...
		Delete "$TEMP\${LOGFILE}"
	${EndIf}
SectionEnd

Function GetJAWSLanguages
; GetJAWSLanguages(sJAWSPath) --> sLangList
; Given a path to a JAWS All Users folder, return a |-delimited list of languages under it.
; Send in sJAWSPath on the stack and retrieve sLangList from the stack.
; Local register usage (no registers clobbered:
;	$0 = inbound JAWSPath string.
;	$1 = FindFirst/Next handle.
;	$2 = FindFirst/Next file name (no path).
;	$9 = Accumulator of language codes.
	exch $0
	push $9
	push $1
	push $2
	ClearErrors
	strcpy $9 ""
	FindFirst $1 $2 "$0\*"
	${DoUntil} ${Errors}
		${If} $2 == "."
		${OrIf} $2 == ".."
			FindNext $1 $2
			${Continue}
		${EndIf}
		; This means "If $2 is a language directory"
		${If} ${FileExists} "$0\$2\default.jcf"
			${If} $9 != ""
				strcpy $9 "$9|"
			${EndIf}
			strcpy $9 "$9$2"
		${EndIf}
		FindNext $1 $2
	${Loop}
	FindClose $1
	pop $2
	pop $1
	exch
	pop $0
	exch $9
FunctionEnd

;--------------------------------
; Uninstaller functions and sections
	; Paths containing installations according to the Installs.ini file created by the installer.
	; This is updated by un.displayUninstallOptions to omit paths that don't actually contain the scripts.
	var InstalledPathsList
	; Paths the user wants to uninstall from (subset of above).
	var PathsList

Function un.OnInit
; Uninstaller startup initialization code.
StrCpy $DetailLines ""
IntOp $DetailLineCount $DetailLineCount * 0
	!IfDef DBG
	SetDetailsView Show
	SetAutoClose False
	!EndIf
	${If} ${RunningX64}
		SetRegView 64
	${EndIf}
	!InsertMacro INSTALLOPTIONS_EXTRACT "${UNINSTINIFILE}"
FunctionEnd

Function un.getListItem
; Given a JAWS user folder path, return a displayable listItem string for it.
; Example: For a JAWS 18.0 Enu path: "JAWS 18.0 (enu) installed in <path>."
; For a JAWS 2018 Enu path: "JAWS 2018 (enu) installed in <path>."
; Also sets $JAWSBuild and $JAWSLang from the path.
; Registers (preserved):
;	$R1: Path passed on stack.
;	$R2: JAWS version from path (e.g., 18.0) or 2018.
;	$R3: JAWS language code from path (e.g., enu).
;	$R4: String index counter for finding the JAWS version part of the path.
;	$R5: Character scratch area for JAWS version finding.
	exch $R1
	push $R2
	push $R3
	push $R4
	push $R5
	; Language code from right end (three characters required)
	strcpy $R3 $R1 1 -4
	${If} $R3 != "\"
		${ErrMsg} "!GetListItem Error: Invalid path (no language folder) $R1."
		strcpy $R1 ""
		goto liExit
	${EndIf}
	strcpy $R3 $R1 3 -3
	; JAWS version from middle
	; First remove '\settings\<lng>'
	strcpy $R2 $R1 -13
	; Then find the last backslash.
	IntOp $R4 $R4 * 0
	${Do}
		IntOp $R4 $R4 - 1
		${If} $R4 = -6
			${ErrMsg} "!GetListItem Error: Invalid path (no JAWS version) $R1."
			strcpy $R1 ""
			goto liExit
		${EndIf}
		strcpy $R5 $R2 1 $R4
		${If} $R5 == "\"
			; Success, found the backslash.
			${Break}
		${EndIf}
	${Loop}
	; $R2[$R4] is the '\' before the version code.
	; $R4 is negative, counting from the end.
	IntOp $R4 $R4 + 1
	; Now it points to the first character of the version code.
	strcpy $R2 $R2 "" $R4
	strcpy $R1 "JAWS $R2 ($R3) installed in $R1"
	strcpy $JAWSBuild $R2
	strcpy $JAWSLang $R3
liExit:
	pop $R5
	pop $R4
	pop $R3
	pop $R2
	exch $R1
FunctionEnd

Function un.DisplayUninstallOptions
; Page that appears before the Uninstall section does.
	call un.getUninstallerPath
	pop $1
	ReadIniStr $R5 "$1\${INSTALLSINIFILE}" ${INSTALLSINISECT} "InstalledForJFWVersions"
	${If} $R5 == ""
		goto noScripts
	${EndIf}
	StrCpy $InstalledPathsList $R5

	; $R5 remaining install paths, $R9 paths with actual scripts,
	; $R6 next path to check, $R4 and $R8 helpers for InstallOptions list building.
	StrCpy $R9 ""
	${DoWhile} $R5 != ""
		push $R5
		push "|"
		call un.SplitFirstStrPart
		pop $R6
		pop $R5

		${IfNot} ${FileExists} "$R6\musescore3.jsb"
			; Skip this version; these scripts aren't in it.
			${Continue}
		${EndIf}
		push $R6
		call un.getListItem
		pop $R8

		; For valid versions of JAWS fill up a value list in .ini configuration file
		!InsertMacro INSTALLOPTIONS_READ $R4 "${UNINSTINIFILE}" "Field 2" "ListItems"
		${If} $R4 == ""
			!InsertMacro INSTALLOPTIONS_WRITE "${UNINSTINIFILE}" "Field 2" "ListItems" "$R8"
		${Else}
			!InsertMacro INSTALLOPTIONS_WRITE "${UNINSTINIFILE}" "Field 2" "ListItems" "$R4|$R8"
		${EndIf}
		StrCpy $R9 "$R9|$R6"
	${Loop}

	; This automates unregistration of installations whose files were manually removed by the user.
	StrCpy $R9 $R9 "" 1
	StrCpy $InstalledPathsList $R9
	push $1
	call un.getUninstallerPath
	pop $1
	writeIniStr "$1\${INSTALLSINIFILE}" ${INSTALLSINISECT} "InstalledForJFWVersions" "$InstalledPathsList"
	pop $1
	${If} $R9 == ""
noScripts:
		${IfNot} ${Silent}
			MessageBox MB_YESNO "No installations of these scripts has been found.  Would you like to remove this script uninstaller and its Start menu shortcut?" \
				/SD IDYES IDYES +1 IDNO +2
		${EndIf}
		call un.removeUninstaller
		Quit
	${EndIf}

	; This list may have only one entry,
	; but we show it anyway (unlike in the installer)
	; as a means of getting user confirmation before uninstall.
	!insertmacro INSTALLOPTIONS_READ $R4 "${UNINSTINIFILE}" "Field 2" "ListItems"
	; Select all by default.
	!InsertMacro INSTALLOPTIONS_WRITE "${UNINSTINIFILE}" "Field 2" "state" "$R4"
	!InsertMacro INSTALLOPTIONS_DISPLAY_RETURN "${UNINSTINIFILE}"
	Pop $R8 ; Pop the return value from the stack
	StrCmp $R8 "cancel" CancelUninstall +1
	!InsertMacro INSTALLOPTIONS_READ $R5 "${UNINSTINIFILE}" "Field 2" "State"
	!InsertMacro INSTALLOPTIONS_WRITE "${UNINSTINIFILE}" "Field 2" "ListItems" "" ; clean up the ListItems value

	; Convert $R5, the list of ListItems containing paths to uninstall from,
	; to $R1, just a list of paths to uninstall from.
	StrCpy $R1 ""
	${DoWhile} $R5 != ""
		push $R5
		push "|"
		call un.splitFirstStrPart
		pop $R4
		pop $R5
		; Extract the path from the selected value returned by "state"
		Push $R4
		Push " in "
		Call un.strstr
		Pop $R4
		; Strip the " in " from the beginning of the $R4 string
		StrCpy $R4 $R4 "" 4
		StrCpy $R1 "$R1|$R4"
	${Loop}
	; Strip the leading "|" from the list of paths
	StrCpy $R1 $R1 "" 1
	Goto done

CancelUninstall:
	MessageBox MB_OK "You have chosen to cancel this uninstall wizard. The uninstaller will now quit.$\n" /SD IDOK
	Quit

done:
	StrCpy $PathsList $R1
	call un.HandleInstallationPaths
FunctionEnd

Function un.removeUninstaller
; Remove the uninstaller and its folder, and its Start menu shortcut and folder.
; The uninstaller's folder also contains the ini file used to keep track of installations.
; Call this when it's time to clean up everything - i.e., no scripts remain.
	; Remove the uninstaller and its folder.
	call un.getUninstallerPath
	pop $INSTDIR
	clearErrors
	Delete "$INSTDIR\${UnInstaller}"
	; In case it exists...
	Delete "$TEMP\${LOGFILE}"
	Delete "$INSTDIR\*.*"
	RmDir "$INSTDIR"
	${If} ${Errors}
		${ErrMsg} "!Error removing the uninstaller and/or its folder."
	${EndIf}
	; Remove the uninstaller's Start menu shortcut and its folder.
	clearErrors
	Delete "$SMPROGRAMS\${NAME}\*.*"
	RmDir "$SMPROGRAMS\${NAME}"
	${If} ${Errors}
		${ErrMsg} "!Error removing the uninstaller's Start menu entry and/or its folder."
	${EndIf}
	clearErrors
FunctionEnd

Function un.HandleInstallationPaths
; Update $InstalledPathsList to contain only the paths containing scripts that the user is NOT uninstalling.
; Treating them as sets, this is  $InstalledPathsList -= $PathsList.
; $R0 paths containing installations, $R1 paths from which to uninstall.
; $0 and $1 used as "next value" for each of those lists.
	StrCpy $R0 $InstalledPathsList
	StrCpy $R1 $PathsList

	; Put in $R9 the paths containing scripts that the user is NOT uninstalling.
	; $R1 is a subset of $R0 and they are in the same order.
	StrCpy $R9 ""
	${DoWhile} $R1 != ""
		; Next path the user is uninstalling from.
		Push $R1
		push "|"
		call un.SplitFirstStrPart
		Pop $1
		pop $R1
		${DoWhile} $R0 != ""
			; Put in $R9 any path(s) before this one from $R0.
			Push $R0
			push "|"
			call un.SplitFirstStrPart
			Pop $0
			Pop $R0
			${If} $0 == $1
				${Break}
			${EndIf}
			StrCpy $R9 "$R9|$0"
		${Loop}
	${Loop}
	; Include any paths after the last one the user asked to uninstall from.
	${If} $R0 != ""
		StrCpy $R9 "$R9|$R0"
	${EndIf}
	; Strip the leading "|" and store remaining paths, if any, into $PathsList
	StrCpy $InstalledPathsList $R9 "" 1
FunctionEnd

Function un.StrStr
; searches for a given string.
; Returns the original string starting at the beginning of the one found.
	Exch $R1 ; st=haystack,old$R1, $R1=needle
	Exch    ; st=old$R1,haystack
	Exch $R2 ; st=old$R1,old$R2, $R2=haystack
	Push $R3
	Push $R4
	Push $R5
	StrLen $R3 $R1
	StrCpy $R4 0
	; $R1=needle
	; $R2=haystack
	; $R3=len(needle)
	; $R4=cnt
	; $R5=tmp
loop:
	StrCpy $R5 $R2 $R3 $R4
	StrCmp $R5 $R1 done
	StrCmp $R5 "" done
	IntOp $R4 $R4 + 1
	Goto loop
done:
	StrCpy $R1 $R2 "" $R4
	Pop $R5
	Pop $R4
	Pop $R3
	Pop $R2
	Exch $R1
FunctionEnd

; uninstall section.
Section "Uninstall"
	StrCpy $0 $PathsList
	StrCmp $0 "" CancelUninstall

	${DoWhile} $0 != ""
		Push $0
		push "|"
		Call un.SplitFirstStrPart
		pop $INSTDIR
		; Set $JAWSBuild and $JAWSLang before restoring $0.
		push $INSTDIR
		call un.getListItem
		pop $0  ; we don't need the display format this time.
		Pop $0  ; this comes from the split call above.
		call un.removeScripts
	${Loop}
	; If we have script versions remaining, do not remove any of the following, just update the Installs ini file
	StrCmp $InstalledPathsList "" +1 JustUpdateIni
	; Remove uninstaller, Start menu shortcut and folder, and installation registration(s).
	call un.removeUninstaller
	GoTo done

JustUpdateIni:
	call un.getUninstallerPath
	pop $1
	writeIniStr "$1\${INSTALLSINIFILE}" ${INSTALLSINISECT} "InstalledForJFWVersions" "$InstalledPathsList"
	Goto done

CancelUninstall:
	MessageBox MB_OK "No JAWS folders selected.  Aborting uninstall." /SD IDOK
	Quit

done:
SectionEnd

!macro DupCode UN
Function ${UN}getUninstallerPath
; Return on stack the path to the uninstaller folder.
	push "$APPDATA\${NAME}"
FunctionEnd

Function ${UN}errmsg
; Put the given error message in the Details list.
; If it starts with "!", remove it, put it in Details, and also show it in a message box.
; Usage: push msgText (with or without "!"), call.  No return value.
; Better yet, ${ErrMsg} msgText.
	exch $1
	push $2
	strcpy $2 $1 1
	${If} $2 == "!"
		strcpy $1 $1 "" 1
		${DetailPrint} $1
		messageBox MB_OK $1 /SD IDOK
	${Else}
		${DetailPrint} $1
	${EndIf}
	pop $2
	pop $1
FunctionEnd

Function ${UN}SplitFirstStrPart
; Take a string on the stack, split it at the first delimiter, and return the two parts on the stack.
; Usage: push $s; push $delim; call; pop $firstSeg; pop $rest.
; Registers are preserved.
; Author:  Presumed to be Victor Tsaran unless he got it elsewhere [DGL].
; Updated by Doug Lee 2010-04-01 to allow passed delimiter.
; Registers: $0 delim, $R0 string.
	Exch $0
	exch
	exch $R0
	Push $R1
	Push $R2
	StrLen $R1 $R0
	IntOp $R1 $R1 + 1

loop:
	IntOp $R1 $R1 - 1
	StrCpy $R2 $R0 1 -$R1
	StrCmp $R1 0 exit0
	StrCmp $R2 $0 exit1 loop

exit0:
	StrCpy $R1 ""
	Goto exit2

exit1:
	IntOp $R1 $R1 - 1
	StrCpy $R2 $R0 "" -$R1
	IntOp $R1 $R1 + 1
	StrCpy $R0 $R0 -$R1
	StrCpy $R1 $R2

exit2:
	Pop $R2
	Exch $R1 ;Rest
	; The stack is now Rest, $R0, $0.
	; It needs to become First, Rest.
	exch 2  ; now $0, $R0, Rest
	pop $0  ; now $R0, Rest
	Exch $R0 ;First
FunctionEnd

Function ${UN}removeScripts
; Remove scripts
; $INSTDIR must be accurate before calling.
	push $R9
	; Revert JAWS user folder files.

	pop $R9
	!IfNDef DBG
	; Status bar shows individual files but list box won't.
	setDetailsPrint textonly
	!EndIf
	delete "$INSTDIR\musescore3.jkm"
	delete "$INSTDIR\musescore3.jsb"
	delete "$INSTDIR\musescore3.jsd"
	delete "$INSTDIR\musescore3.jsm"
	delete "$INSTDIR\musescore3.jss"

	!IfNDef DBG
	; List box again shows everything.
	setDetailsPrint both
	!EndIf
FunctionEnd

Function ${UN}Quit
; Use this in place of the built-in Quit instruction to enable log writing.
; Does not preserve registers because it doesn't return anyway.
	${ErrMsg} "Writing log file."
	StrCpy $0 "$TEMP\${LOGFILE}"
	push $0
	call ${UN}DumpLog
	pop $1
	${If} $1 = 0
		${ErrMsg} "!Failed to write a log of this process to disk for examination by the script author."
		Quit
	${EndIf}
	${ErrMsg} "Log file written successfully."
	MessageBox MB_OK "A log of this run has been written to disk and will be shown when this installer exits.$\n\
		Sending some or all of this log to the script author may help determine the cause of any problems encountered." /SD IDOK
	ExecShell "open" "$0"
	${If} ${Errors}
		${ErrMsg} "!Unable to display $0"
	${EndIf}
	Quit
FunctionEnd

!IfNDef LVM_GETITEMCOUNT
!define LVM_GETITEMCOUNT 0x1004
!define LVM_GETITEMTEXT 0x102D
!EndIf
Function ${UN}DumpLog
; Dump the contents of the Details list (log) to a file.
; Usage: push filename, call, pop result code.
; This will dump the log to the file specified. For example:
;	GetTempFileName $0
;	Push $0
;	Call DumpLog
;	Pop $1
; written by KiCHiK
; Taken directly from the NSI chm documentation by Doug Lee, 2011-02-15.
; Modified by Doug Lee 2017-09-02 to fall back on a global DetailLines variable when the Details list is not available at call time.
; This is an overwrite operation.  If the uninstaller is launched by the installer,
; you must grab the uninstaller log before the installer can replace it.
; Register use within this function (DGL):
;	$0 hwnd of Details list control (when found).
;	$1 Space used in transferring Details list items to disk???
;	$2 Loop line counter.
;	$3 System alloc space for transfer of Details list items to disk.
;	$4 Line being written out to disk from within the loop.
;	$5 Log file name (in), and return value (out).
;	$6 Line count.
;	$7 1 when ListView is used and 0 otherwise.
	Exch $5
	Push $0
	Push $1
	Push $2
	Push $3
	Push $4
	Push $6
	Push $7
	; Default to using list unless we can't.
	IntOp $7 $7 * 0
	IntOp $7 $7 + 1

	FindWindow $0 "#32770" "" $HWNDPARENT
	GetDlgItem $0 $0 1016
	${If} $0 = 0
		${ErrMsg} "Unable to find the Details list; log may be less detailed."
		IntOp $7 $7 * 0
	${EndIf}
	FileOpen $5 $5 "w"
	${If} $5 = 0
		${ErrMsg} "!Unable to open the log file for writing."
		StrCpy $5 0
		Goto exit
	${EndIf}
	${If} $7 = 1
		SendMessage $0 ${LVM_GETITEMCOUNT} 0 0 $6
		System::Alloc ${NSIS_MAX_STRLEN}
		Pop $3
		System::Call "*(i, i, i, i, i, i, i, i, i) i \
			(0, 0, 0, 0, 0, r3, ${NSIS_MAX_STRLEN}) .r1"
	${Else}
		IntOp $6 $DetailLineCount * 1
	${EndIf}
	StrCpy $2 0
	loop: StrCmp $2 $6 done
		${If} $7 = 1
			System::Call "User32::SendMessageA(i, i, i, i) i \
				($0, ${LVM_GETITEMTEXT}, $2, r1)"
			System::Call "*$3(&t${NSIS_MAX_STRLEN} .r4)"
		${Else}
			Push $DetailLines
			Push "$\n"
			call ${UN}SplitFirstStrPart
			Pop $4
			pop $DetailLines
		${EndIf}
		FileWrite $5 "$4$\r$\n"
		IntOp $2 $2 + 1
		Goto loop
	done:
		FileClose $5
		${If} $7 = 1
			System::Free $1
			System::Free $3
		${EndIf}
exit:
	Pop $7
	Pop $6
	Pop $4
	Pop $3
	Pop $2
	Pop $1
	Pop $0
	Exch $5
FunctionEnd

!macroend
!insertmacro DupCode ""
!insertmacro DupCode "un."

