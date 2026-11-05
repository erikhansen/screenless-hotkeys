-- Helper called from Word VBA via AppleScriptTask("JTHotkeys.scpt", "speak", text).
-- Lives in ~/Library/Application Scripts/com.microsoft.Word/ (install.sh puts it there).
on speak(textToSay)
	do shell script "/usr/bin/say " & quoted form of textToSay & " > /dev/null 2>&1 &"
	return ""
end speak

-- Same as speak, but first stops any speech already in progress.
on speakNow(textToSay)
	do shell script "/usr/bin/pkill -x say; /usr/bin/say " & quoted form of textToSay & " > /dev/null 2>&1 &"
	return ""
end speakNow

-- Highest CUPS job number so far (all printers, including finished jobs).
-- Word VBA uses it to tell when its print job has reached the queue.
on lastPrintJob(unused)
	try
		return do shell script "/usr/bin/lpstat -W all -o 2>/dev/null | /usr/bin/awk '{print $1}' | /usr/bin/sed 's/.*-//' | /usr/bin/sort -n | /usr/bin/tail -1"
	on error
		return ""
	end try
end lastPrintJob

-- Starts JT Hotkeys Helper.app, which clicks the ChatGPT add-in button.
-- It is started through launchd (a LaunchAgent that install.sh sets up), not
-- with "open": an app opened by Word borrows Word's Accessibility permission,
-- while one started by launchd gets its own entry in Privacy & Security.
on openChatGPT(unused)
	set uid to do shell script "/usr/bin/id -u"
	try
		do shell script "/bin/launchctl kickstart gui/" & uid & "/com.screenless-hotkeys.helper"
	on error
		return "missing"
	end try
	return ""
end openChatGPT
