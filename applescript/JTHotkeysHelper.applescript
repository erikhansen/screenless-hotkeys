-- JT Hotkeys Helper.app: clicks the ChatGPT add-in button on Word's Home tab.
--
-- Word can't assign keyboard shortcuts to add-in buttons, and clicking a
-- button needs the macOS Accessibility permission. So the JT_OpenChatGPT macro
-- (Control+Option+Command+G) starts this small app, and the app holds that
-- permission. install.sh builds it into ~/Applications only when it is missing
-- or its version changes, because rebuilding makes macOS forget the permission.

property helperVersion : "2"

on run
	try
		set result_ to clickChatGPT()
	on error errMsg number errNum
		if errNum is -1743 then
			-- Not allowed to control System Events (Automation permission).
			speak("J T Hotkeys Helper needs permission to control System Events. Turn it on in Privacy and Security, Automation.")
			openSettings("Privacy_Automation")
		else if errNum is -25211 or errNum is -1719 or errMsg contains "assistive" then
			-- Not allowed to use accessibility features.
			speak("J T Hotkeys Helper needs Accessibility permission. Turn it on in Privacy and Security, Accessibility, then press the hotkey again.")
			openSettings("Privacy_Accessibility")
		else
			speak("Could not open ChatGPT. " & errMsg)
		end if
		return
	end try
	if result_ is "clicked" then
		speak("ChatGPT.")
	else if result_ is "no window" then
		speak("Open a Word document first, then press the hotkey again.")
	else
		speak("The ChatGPT button was not found on the Home tab.")
	end if
end run

on clickChatGPT()
	tell application "System Events" to tell process "Microsoft Word"
		if (count of windows) = 0 then return "no window"
		set tg to first UI element of window 1 whose role is "AXTabGroup"
		set homeTab to first radio button of tg whose name is "Home"
		-- Clicking an already-selected tab collapses the ribbon, so only switch when needed.
		if (value of homeTab as text) is not "1" or (count of (UI elements of tg whose role is "AXScrollArea")) = 0 then
			click homeTab
			delay 0.3
		end if
		set sa to first UI element of tg whose role is "AXScrollArea"
		repeat with g in UI elements of sa
			try
				click (first button of g whose name is "ChatGPT")
				return "clicked"
			end try
		end repeat
	end tell
	return "not found"
end clickChatGPT

on speak(t)
	do shell script "/usr/bin/say " & quoted form of t & " > /dev/null 2>&1 &"
end speak

on openSettings(anchor)
	do shell script "/usr/bin/open " & quoted form of ("x-apple.systempreferences:com.apple.preference.security?" & anchor)
end openSettings
