; (c) 2020-06-20 by Zbigniew Gralewski zbigniew.gralewski (at) bitback.pl
;
; Quick info:
; Script massively renames eml files, for example from "foo.eml" into "2020-01-01_101002__foo.eml" using the format RRRR-MM-DD_HH_MM_SS__originalname.eml.
;
;
; ATTENTION!
; This script modifies and renames files in current directory. You cannot revert what is done if you don't have a copy so make a copy!
;
;
; WHAT IS IT FOR:
; To simplify manual sorting of eml files into yearly subfolders until Thunderbird maildir format is stable enough to allow moving emails in it.
;
;
; UNDERSTAND HOW DOES IT WORK:
; Loops all eml files in current directory. Then in the loop:
;	Skips the file if it is already converted or has bad or non existent date.
;	Opens each file and looks for "Date:" field and extracts it's value using regular expressions.
;	If the date is non existing or is in bad format the file will have "_nodate" or "_baddate" in the beginning of new name.
;	Converts date format from "1 Jun 2020" into "2020-06-01".
;	Renames each file so the date is in the beginning of the name of each eml file.
;
;
; USAGE:
; 0. Install Autohotkey just to be able to run AHK files.
; 1. Put all eml files manually into some directory (work on a copy or make a backup!).
; 2. Put copy of this script into that folder and run it, eml files will have new names.
; 3. Select files, move them manually into subfolders as you like. Remember that Thunderbird needs "cur" directory to see emails in it.
; 4. Remember to delete recursively msf files in order to force Thunderbird to refresh the folder and reindex all eml files.
; 5. Send feedback how happy you are ;)
;
;
; NOTES:
; The script is secured to only do it's job one time. It will skip all the files already renamed.
; Remember to delete msf files if you have problems with thunderbird not indexing files after moving them manually into "cur" directories.
; You may want to use "_del_msf.bat" file with the content: "@del /S *.msf" to use it in TB maildir directory to clear msf files recursively.
;
;
#SingleInstance force
#Persistent
#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%

ScanEmlFiles()
goto, end

ScanEmlFiles()
{
	; loop filename by filename
	Loop, Files, %A_WorkingDir%\*.eml
	{
		filename := A_LoopFileName
		date := "_nodate"
		
		if InStr(filename, "_nodate")
			continue
		
		if InStr(filename, "_baddate")
			continue
			
		if (RegExMatch(filename, "^\d{4}-\d{2}-\d{2}_\d{6}.+$"))
			continue
		
		; loop line by line
		Loop, Read, %filename%
		{
			line := Trim(A_LoopReadLine)
			if (SubStr(line, 1, 5) = "Date:") ; Look for date field
			{
				date := ParseDateField(line)
				break ; found date field so break
			}
		}
		; here we have filename and newname set
		newname = %date%__%filename%
		;MsgBox, %filename%`n`n%newname%
		FileMove, %filename%, %newname%
	}
}

ParseDateField(datestring)
{
	; expected string formats and parsed subpattern numbers:
	;
	; Date: Mon, 22 Jul 2019 11:45:41 +0200
	; sub        ^1 ^2  ^3   ^4       ^5
	;
	; Date: 1 Jun 2020 15:56:28 +0200
	; sub   ^1^2  ^3   ^4       ^5 
	;
	pattern := "^Date:\s+(?:[a-zA-Z]{3},\s+)?(\d+)\s+([a-zA-Z]{3})\s+(\d{4})\s+(\d\d:\d\d:\d\d)\s+(.+)"
	r := RegExMatch(datestring, pattern, sub)
	if (r = 1)
	{
		sub1 := Format("{:02}",sub1)
		sub2 := StrReplace(sub2, "Jan", "01")
		sub2 := StrReplace(sub2, "Feb", "02")
		sub2 := StrReplace(sub2, "Mar", "03")
		sub2 := StrReplace(sub2, "Apr", "04")
		sub2 := StrReplace(sub2, "May", "05")
		sub2 := StrReplace(sub2, "Jun", "06")
		sub2 := StrReplace(sub2, "Jul", "07")
		sub2 := StrReplace(sub2, "Aug", "08")
		sub2 := StrReplace(sub2, "Sep", "09")
		sub2 := StrReplace(sub2, "Oct", "10")
		sub2 := StrReplace(sub2, "Nov", "11")
		sub2 := StrReplace(sub2, "Dec", "12")
		sub4 := StrReplace(sub4, ":", "")
		;MsgBox, %filename%`n1:%sub1%`n2:%sub2%`n3:%sub3%`n4:%sub4%`n5:%sub5%`n6:%sub6%`n7:%sub7%`n8:%sub8%`n9:%sub9%`n10:%sub10%

		newname = %sub3%-%sub2%-%sub1%_%sub4%
	}
	else
	{
		newname = _baddate
	}
	return newname
}

#r::Reload
#e::ExitApp

end:
ExitApp
