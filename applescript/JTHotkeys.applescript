-- Helper called from Word VBA via AppleScriptTask("JTHotkeys.scpt", "speak", text).
-- Lives in ~/Library/Application Scripts/com.microsoft.Word/ (install.sh puts it there).
on speak(textToSay)
	do shell script "/usr/bin/say " & quoted form of textToSay & " > /dev/null 2>&1 &"
	return ""
end speak

-- Highest CUPS job number so far (all printers, including finished jobs).
-- Word VBA uses it to tell when its print job has reached the queue.
on lastPrintJob(unused)
	try
		return do shell script "/usr/bin/lpstat -W all -o 2>/dev/null | /usr/bin/awk '{print $1}' | /usr/bin/sed 's/.*-//' | /usr/bin/sort -n | /usr/bin/tail -1"
	on error
		return ""
	end try
end lastPrintJob
