-- Builds dist/JTHotkeys.dotm by driving Word's VBA editor (UI scripting).
-- Usage: osascript tools/build-dotm.applescript <repoDir> <file1.bas> [file2.bas ...]
-- Needs Accessibility permission for the terminal app.

on findRow(theOutline, label)
	tell application "System Events"
		repeat with r in (rows of theOutline)
			try
				if value of (first static text of (first UI element of r)) is label then return r
			end try
		end repeat
	end tell
	error "Row not found: " & label
end findRow

on selectProject(docName)
	tell application "System Events" to tell process "Microsoft Word"
		set frontmost to true
		delay 0.4
		set vbe to (first window whose name starts with "Microsoft Visual Basic")
		set r to my findRow(outline 1 of vbe, "Project (" & docName & ")")
		set selected of r to true
		delay 0.8
		if name of (first window whose name starts with "Microsoft Visual Basic") is not ("Microsoft Visual Basic - " & docName) then error "Could not select project " & docName
		-- Wait for the VBA editor's own menu bar (it has File > Import File...).
		repeat 20 times
			if exists menu item "Import File..." of menu "File" of menu bar 1 then return
			perform action "AXRaise" of vbe
			set frontmost to true
			delay 0.3
		end repeat
		error "The VBA editor menu never appeared"
	end tell
end selectProject

-- Loads/unloads JTHotkeys.dotm via Tools > Templates and Add-ins.
-- Returns whether it was loaded before.
on setAddinLoaded(wantOn)
	set wasOn to false
	tell application "System Events" to tell process "Microsoft Word"
		set frontmost to true
		delay 0.4
		click menu item "Templates and Add-ins..." of menu "Tools" of menu bar 1
		delay 1.5
		set w to window "Templates and Add-ins"
		repeat with r in rows of table 1 of scroll area 1 of w
			if value of text field 1 of r is "JTHotkeys.dotm" then
				set wasOn to ((value of checkbox 1 of r) is 1)
				if wasOn is not wantOn then click checkbox 1 of r
			end if
		end repeat
		delay 0.5
		click button "OK" of w
		delay 1
	end tell
	return wasOn
end setAddinLoaded

on run argv
	set repoDir to item 1 of argv
	set basFiles to rest of argv
	set containerDocs to (POSIX path of (path to home folder)) & "Library/Containers/com.microsoft.Word/Data/Documents/JTHotkeys.dotm"

	-- Unload the installed copy so macro names aren't ambiguous during the build.
	tell application "Microsoft Word" to activate
	delay 0.5
	set addinWasLoaded to my setAddinLoaded(false)

	with timeout of 60 seconds
		tell application "Microsoft Word"
			activate
			set d to make new document
			set docName to name of d
		end tell
	end timeout
	delay 1

	tell application "System Events" to tell process "Microsoft Word"
		set frontmost to true
		click menu item "Visual Basic Editor" of menu 1 of menu item "Macro" of menu "Tools" of menu bar 1
		delay 2
	end tell

	repeat with f in basFiles
		my selectProject(docName)
		tell application "System Events" to tell process "Microsoft Word"
			click menu item "Import File..." of menu "File" of menu bar 1
			delay 1.5
			keystroke "g" using {command down, shift down}
			delay 1
			keystroke (repoDir & "/" & f)
			delay 0.8
			key code 36
			delay 1.2
			click button "Open" of splitter group 1 of window "Import File"
			delay 1.5
		end tell
	end repeat

	my selectProject(docName)
	tell application "System Events" to tell process "Microsoft Word"
		click menu item "Compile Project" of menu "Debug" of menu bar 1
		delay 2
		if name of window 1 does not start with "Microsoft Visual Basic" then error "Compile error - check the VBA editor: " & name of window 1
		set vbe to (first window whose name starts with "Microsoft Visual Basic")
		click (first button of vbe whose description is "close button")
		delay 1
	end tell

	with timeout of 60 seconds
		tell application "Microsoft Word"
			-- Store the key assignments inside the template.
			run VB macro macro name "JT_BakeHotkeys"
			save as document docName file name containerDocs file format format templateME
			close document "JTHotkeys.dotm" saving no
		end tell
	end timeout
	do shell script "mv " & quoted form of containerDocs & " " & quoted form of (repoDir & "/dist/JTHotkeys.dotm")
	if addinWasLoaded then my setAddinLoaded(true)
	return "Built dist/JTHotkeys.dotm"
end run
