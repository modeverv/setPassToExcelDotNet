set repoRoot to "/Users/seijiro/Sync/sync_work/me/SetPassToExceldotNet"
set evidenceDir to repoRoot & "/docs/manual-evidence/macos"
set workbookDir to evidenceDir & "/workbooks"
set sessionLog to evidenceDir & "/session.log"
set checkValue to "日本語シート名保持確認"

set testCases to {}
set end of testCases to {"simple_en.xlsx", "en", "pass", true}
set end of testCases to {"simple_ja.xlsx", "ja", "パスワード", true}
set end of testCases to {"japanese_en.xlsx", "en", "pass", false}
set end of testCases to {"japanese_ja.xlsx", "ja", "パスワード", false}
set end of testCases to {"excel_en.xlsm", "en", "pass", false}
set end of testCases to {"excel_ja.xlsm", "ja", "パスワード", false}
set end of testCases to {"excel_image_en.xlsx", "en", "pass", false}
set end of testCases to {"excel_image_ja.xlsx", "ja", "パスワード", false}

on appendLog(sessionLog, msg)
	do shell script ("printf '%s\\n' " & quoted form of msg & " >> " & quoted form of sessionLog)
end appendLog

on captureShot(pathToPng)
	delay 1
	do shell script ("/usr/sbin/screencapture -x " & quoted form of pathToPng)
end captureShot

on fileBase(fileName)
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to "."
	set pieces to text items of fileName
	set AppleScript's text item delimiters to oldDelims
	if (count of pieces) is 1 then return fileName
	return items 1 thru -2 of pieces as text
end fileBase

on uniqueJapaneseSheetName(wb)
	tell application "Microsoft Excel"
		set sheetNames to name of every worksheet of wb
	end tell
	set candidate to "日本語シート"
	set suffixNo to 2
	repeat while sheetNames contains candidate
		set candidate to "日本語シート" & suffixNo
		set suffixNo to suffixNo + 1
	end repeat
	return candidate
end uniqueJapaneseSheetName

on closeActiveWorkbookIfAny()
	tell application "Microsoft Excel"
		try
			close active workbook saving no
		end try
	end tell
end closeActiveWorkbookIfAny

appendLog(sessionLog, "== macOS Excel manual verification: " & ((current date) as text) & " ==")

tell application "Microsoft Excel"
	activate
	set display alerts to false
end tell

repeat with tc in testCases
	set fileName to item 1 of tc
	set passwordType to item 2 of tc
	set correctPassword to item 3 of tc
	set shouldCheckWrongPassword to item 4 of tc
	set baseName to my fileBase(fileName)
	set workbookPath to workbookDir & "/" & fileName
	set openShot to evidenceDir & "/" & baseName & "-open.png"
	set addedShot to evidenceDir & "/" & baseName & "-ja-sheet-added.png"
	set reopenShot to evidenceDir & "/" & baseName & "-reopen.png"
	set correctOpen to "FAIL"
	set wrongRejected to "N/A"
	set jaNameRetained to "FAIL"
	set reopenAfterSave to "FAIL"
	set noCorruption to "PASS"
	set addedSheetName to ""
	set noteText to ""
	
	my appendLog(sessionLog, "[" & fileName & "] start")
	my closeActiveWorkbookIfAny()
	
	tell application "Microsoft Excel"
		activate
		set display alerts to false
		try
			set wb to open workbook workbook file name workbookPath password correctPassword
			set correctOpen to "PASS"
			delay 1
			my captureShot(openShot)
			
			set addedSheetName to my uniqueJapaneseSheetName(wb)
			set newWs to make new worksheet at after last worksheet of wb
			set name of newWs to addedSheetName
			set value of range "A1" of active sheet to checkValue
			delay 1
			my captureShot(addedShot)
			save wb
			close wb saving yes
			
			set wb2 to open workbook workbook file name workbookPath password correctPassword
			set reopenAfterSave to "PASS"
			activate object worksheet addedSheetName of wb2
			set reopenedValue to value of range "A1" of active sheet
			if reopenedValue is checkValue then
				set jaNameRetained to "PASS"
			else
				set jaNameRetained to "FAIL"
				set noteText to "A1 mismatch after reopen"
			end if
			delay 1
			my captureShot(reopenShot)
			close wb2 saving no
		on error errMsg number errNo
			set noCorruption to "FAIL"
			set noteText to "Excel error " & errNo & ": " & errMsg
			try
				close active workbook saving no
			end try
		end try
	end tell
	
	if shouldCheckWrongPassword then
		my closeActiveWorkbookIfAny()
		tell application "Microsoft Excel"
			try
				set wrongWb to open workbook workbook file name workbookPath password "__wrong_password__"
				set wrongRejected to "FAIL"
				close wrongWb saving no
			on error errMsg number errNo
				set wrongRejected to "PASS"
			end try
		end tell
	end if
	
	my appendLog(sessionLog, "[" & fileName & "] password-type=" & passwordType & " correct-password-open=" & correctOpen & " wrong-password-rejected=" & wrongRejected & " japanese-sheet-name-retained=" & jaNameRetained & " reopen-after-save=" & reopenAfterSave & " no-corruption=" & noCorruption & " added-sheet=" & addedSheetName & " note=" & noteText)
end repeat

my closeActiveWorkbookIfAny()
appendLog(sessionLog, "== completed: " & ((current date) as text) & " ==")
