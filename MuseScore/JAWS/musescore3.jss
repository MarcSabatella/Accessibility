; MuseScore3 JAWS scripts
;
; Authors: Peter Torpey, Marc Sabatella, Doug Lee
; License: GPL v2 (see license_gpl2.txt)

include "hjConst.jsh"
include "MSAAConst.jsh"
include "MuseScore3.jsm"


Script ScriptFileName ()

var
	string text
text = FormatString(msgScriptInfo, GetAppFileName(), ScriptVersion, ReleaseDate)
if isSameScript()
	SayMessage(OT_USER_BUFFER, text)
else
	SayMessage(OT_MESSAGE, text)
endif

EndScript


void Function WindowActivatedEvent (handle win)

;SayMessage(OT_MESSAGE, "window activated")
var
	string name,
	string text
name = GetWindowName(win)
if name == "MuseScore3"
	; HACK
	; MuseScore has dialogs the do not read well by default
	; we need to read these manually
	; generally, they have only "MuseScore3" as their names
	; known examples include the telemetry dialog at first startup
	; also the error dialog if you try to add text with no selection
	text = GetTypeAndTextStringsForWindow(win)
	SayMessage(OT_MESSAGE, text)
	return
endif
; pass through
WindowActivatedEvent(win)

EndFunction


void Function FocusChangedEventEx (handle newWin, int newObj, int newChild, handle oldWin, int oldObj, int oldChild, int depth)

;SayMessage(OT_MESSAGE, "focus changed")
var
	int role,
	object obj
obj = GetCurrentObject(0)
role = obj.accRole(0)
if (role == ROLE_SYSTEM_CLIENT)
	; navigating a palette scroll area (for example, the key signature selection in New Score Wizard)
	; read current name only (a name changed event was actually generated, but it is not seen by this script)
	say(GetObjectName(), OT_STATUS)
	return
elif newWin == oldWin && newObj == oldObj && newChild == oldChild
	; navigating the score view
	; do nothing (the ValueChangedEvent will read the new value)
	return
endif
; pass through
FocusChangedEventEx(newWin, newObj, newChild, oldWin, oldObj, oldChild, depth)

EndFunction


void Function ValueChangedEvent (handle win, int id, int child, int type, string name, string value, int focused)

;SayMessage(OT_MESSAGE, "value changed")
; read value changes for focused objects, and output to Braille display
if focused
	; object focused - force display/read of value (only)
	BrailleString(value)
	say(value, OT_STATUS)
	return
endif
; pass through
ValueChangedEvent (win, id, child, type, name, value, focused)

EndFunction


void Function NameChangedEvent (handle win, int id, int child, int type, string oldname, string newname)

;SayMessage(OT_MESSAGE, "name changed")
; TODO - make sure key signatures are read in new score wizard for both Qt 5.9 and 5.12
; currently handled in FocusChangedEventEx instead
; pass through
NameChangedEvent(win, id, child, type, oldname, newname)

EndFunction


Script PassKey ()

TypeCurrentScriptKey()

EndScript
