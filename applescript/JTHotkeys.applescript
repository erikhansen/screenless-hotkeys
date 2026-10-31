-- Helper called from Word VBA via AppleScriptTask("JTHotkeys.scpt", "speak", text).
-- Lives in ~/Library/Application Scripts/com.microsoft.Word/ (install.sh puts it there).
on speak(textToSay)
	do shell script "/usr/bin/say " & quoted form of textToSay & " > /dev/null 2>&1 &"
	return ""
end speak
