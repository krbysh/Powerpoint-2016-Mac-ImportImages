
on myFileName(aInt)
	global nameList
	get text item aInt of nameList
end myFileName

on myFileNum()
	global aFol
	global nameList
	local num
	tell application "Finder"
		tell folder aFol
			set indList to a reference to (every file whose name extension is "png")
		end tell
		set nameList to name of indList
		set num to length of nameList
	end tell
end myFileNum

on myFolderName(Fol)
	global aFol
	set aFol to "Macintosh HD" & replaceText(Fol, "/", ":")
end myFolderName

on replaceText(tText, bStr, aStr)
	local tmp
	set tmp to AppleScript's text item delimiters
	set AppleScript's text item delimiters to bStr
	set theList to every text item of tText
	set AppleScript's text item delimiters to aStr
	set tText to theList as string
	set AppleScript's text item delimiters to tmp
	return tText
end replaceText
