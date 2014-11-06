; CSV_TPL_16.au3
; AutoIt script for converting scan.CSV files to Tecan EVOware TPL files
;
; Version 1.6, Okt 2014 , dion methorst : )
;
; changelog:
; july 2014: improved filename retrieval, no more dependance on length of string
;
; Samplenames read from barcodes are retrieved from C:\APPS\EVO\CSV\scan.csv.
; Then converted to appropriate TPL file format and written to C:\APPS\EVO\TPL\plateID.TPL (for instance)
;
; *******************************************************************************************************************************
; 	scan.csv file format
;
;	F411F711
;	1;1;1;Tube 13*75mm 16 Pos;tube1;013/509677;Sample_1
;	1;1;2;Tube 13*75mm 16 Pos;tube1;013/509677;Sample_2
;	1;1;3;Tube 13*75mm 16 Pos;tube1;013/509677;Sample_3
;	1;1;4;Tube 13*75mm 16 Pos;tube1;013/509677;Sample_4
;	1;1;5;Tube 13*75mm 16 Pos;tube1;013/509677;Sample_5
;	1;1;6;Tube 13*75mm 16 Pos;tube1;013/509677;Sample_6
;	1;1;7;Tube 13*75mm 16 Pos;tube1;013/509677;Sample_7
;	1;1;8;Tube 13*75mm 16 Pos;tube1;013/509677;Sample_8
;	1;1;9;Tube 13*75mm 16 Pos;tube1;013/509677;Sample_9
;	1;1;10;Tube 13*75mm 16 Pos;tube1;013/509677;Sample_10
;	1;1;11;Tube 13*75mm 16 Pos;tube1;013/509677;$$$
;	1;1;12;Tube 13*75mm 16 Pos;tube1;013/509677;$$$
;	1;1;13;Tube 13*75mm 16 Pos;tube1;013/509677;stand
;	1;1;14;Tube 13*75mm 16 Pos;tube1;013/509677;$$$
;	1;1;15;Tube 13*75mm 16 Pos;tube1;013/509677;$$$
;	1;1;16;Tube 13*75mm 16 Pos;tube1;013/509677;$$$
;	21;1;;96 Well DeepWell portait;Labware7;023/004256;PlateID_ANYNAME_ANYLENGTH
; *******************************************************************************************************************************
;
; start of script
;
;include library files for <related> functions
#include <Array.au3>
#include <file.au3>
#include <Date.au3>
#include <String.au3>

;imported function for reading a CSV file into an array
; #FUNCTION# ====================================================================================================================
; Name...........: _ParseCSV
; Description ...: Reads a CSV-file
; Syntax.........: _ParseCSV($sFile, $sDelimiters=',', $sQuote='"', $iFormat=0)
; Parameters ....: $sFile       - File to read or string to parse
;                  $sDelimiters - [optional] Fieldseparators of CSV, mulitple are allowed (default: ,;)
;                  $sQuote      - [optional] Character to quote strings (default: ")
;                  $iFormat     - [optional] Encoding of the file (default: 0):
;                  |-1     - No file, plain data given
;                  |0 or 1 - automatic (ASCII)
;                  |2      - Unicode UTF16 Little Endian reading
;                  |3      - Unicode UTF16 Big Endian reading
;                  |4 or 5 - Unicode UTF8 reading
; Return values .: Success - 2D-Array with CSV data (0-based)
;                  Failure - 0, sets @error to:
;                  |1 - could not open file
;                  |2 - error on parsing data
;                  |3 - wrong format chosen
; Author ........: ProgAndy
; Modified.......:
; Remarks .......:
; Related .......: _WriteCSV
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _ParseCSV($sFile, $sDelimiters=';', $sQuote='"', $iFormat=0)
	Local Static $aEncoding[6] = [0, 0, 32, 64, 128, 256]
	If $iFormat < -1 Or $iFormat > 6 Then
		Return SetError(3,0,0)
	ElseIf $iFormat > -1 Then
		Local $hFile = FileOpen($sFile, $aEncoding[$iFormat]), $sLine, $aTemp, $aCSV[1], $iReserved, $iCount
		If @error Then Return SetError(1,@error,0)
		$sFile = FileRead($hFile)
		FileClose($hFile)
	EndIf
	If $sDelimiters = "" Or IsKeyword($sDelimiters) Then $sDelimiters = ';'
	If $sQuote = "" Or IsKeyword($sQuote) Then $sQuote = '"'
	$sQuote = StringLeft($sQuote, 1)
	Local $srDelimiters = StringRegExpReplace($sDelimiters, '[\\\^\-\[\]]', '\\\0')
	Local $srQuote = StringRegExpReplace($sQuote, '[\\\^\-\[\]]', '\\\0')
	Local $sPattern = StringReplace(StringReplace('(?m)(?:^|[,])\h*(["](?:[^"]|["]{2})*["]|[^,\r\n]*)(\v+)?',';', $srDelimiters, 0, 1),'"', $srQuote, 0, 1)
	Local $aREgex = StringRegExp($sFile, $sPattern, 3)
	If @error Then Return SetError(2,@error,0)
	$sFile = '' ; save memory
	Local $iBound = UBound($aREgex), $iIndex=0, $iSubBound = 1, $iSub = 0
	Local $aResult[$iBound][$iSubBound]
	For $i = 0 To $iBound-1
		Select
			Case StringLen($aREgex[$i])<3 And StringInStr(@CRLF, $aREgex[$i])
				$iIndex += 1
				$iSub = 0
				ContinueLoop
			Case StringLeft(StringStripWS($aREgex[$i], 1),1)=$sQuote
				$aREgex[$i] = StringStripWS($aREgex[$i], 3)
				$aResult[$iIndex][$iSub] = StringReplace(StringMid($aREgex[$i], 2, StringLen($aREgex[$i])-2), $sQuote&$sQuote, $sQuote, 0, 1)
			Case Else
				$aResult[$iIndex][$iSub] = $aREgex[$i]
		EndSelect
		$aREgex[$i]=0 ; save memory
		$iSub += 1
		If $iSub = $iSubBound Then
			$iSubBound += 1
			ReDim $aResult[$iBound][$iSubBound]
		EndIf
	Next
	If $iIndex = 0 Then $iIndex=1
	ReDim $aResult[$iIndex][$iSubBound]
	Return $aResult
EndFunc

$protocol = 100

Select	; determine Windows OS version to select path of Evoware Trace.txt
	Case @OSVersion = "WIN_81" OR "WIN_8" OR "WIN_7" OR "WIN_VISTA"
		$FolderPath = "C:\ProgramData\Tecan\EVOware\output\"
		;$FolderPath = "C:\APPS\EVO\CSV\" ; for testing on non evo pc's
		$FileNam2 = "scan.csv"
$var = $FolderPath &$FileNam2
	Case @OSVersion = "WIN_XP"
		$FolderPath = "C:\Program Files\Tecan\EVOware\output\"
		;$FolderPath = "C:\APPS\EVO\CSV\" ; for testing on non evo pc's
		$FileNam2 = "scan.csv"
		$var = $FolderPath &$FileNam2
EndSelect

$FolderPath = "C:\ProgramData\Tecan\EVOware\output\"
;$FolderPath = "C:\APPS\EVO\CSV\" ; for testing on non evo pc's
$FileNam2 = "scan.csv"
$var = $FolderPath &$FileNam2

$GetName = FileReadline($var, -1)											; the contents of the last line of the csv-file read into a string
$filename = StringTrimleft ( $GetName,StringInStr($GetName,";",0,-1))		;
   If FileExists($var) Then
	   $TPL = FileOpen("C:\apps\EVO\TPL\" & $filename & ".tpl", 2 + 8)	; TPL file creation
	  Else
		 MsgBox(4096, $var, "Scan.csv does NOT exist in folder:" & @CRLF & @CRLF & "C:\ProgramData\Tecan\EVOware\output\")
		 ;MsgBox(4096, $var, "Scan.csv does NOT exist in folder:" & @CRLF & @CRLF & "C:\apps\EVO\CSV")
	  EndIf

$Names = _ParseCSV($var, ";", '$', 4)
	;_ArrayDisplay($Names) ; for testing purposes
    _ArrayDelete($Names, 17)
	_ArrayDelete($Names , 16)
	_ArrayDelete($Names , 15)
	_ArrayDelete($Names , 14)
	_ArrayDelete($Names , 13)
	_ArrayDelete($Names , 12)
	_ArrayDelete($Names , 11)
	_ArrayDelete($Names , 0)
	;_ArrayDisplay($Names) ; for testing purposes

for $Trim = 0 to 9

	$Names[$Trim][0] = StringTrimleft ( $Names[$Trim][0],StringInStr($Names[$Trim][0],";",0,-1))

next
;_ArrayDisplay($Names) ; for testing puposes

; START of the main loop to write the samplenumbers extracted from the $Names array with the corresponding plateposition & dilutionprotocol
For $Write = 0 To 9 							;Ubound($Names)-1
												;Write TPL file line by line
		For $Cycle = 1 to 8

		Select
			Case $Cycle = 1
				$Cell = "A" & ($Write + 1)
				$Vx = "V1"
			   FileWriteline($TPL, "D;" & $protocol & ";" & $Names[$Write][0] & ";" & $Cell & ";" & $Vx & ";" & @CRLF)
			Case $Cycle = 2
				$Cell = "B"  & ($Write + 1)
				$Vx = "V2"
			   FileWriteline($TPL, "D;" & $protocol & ";" & $Names[$Write][0] & ";" & $Cell & ";" & $Vx & ";" & @CRLF)
			Case $Cycle = 3
				$Cell = "C"  & ($Write + 1)
				$Vx = "V3"
				FileWriteline($TPL, "D;" & $protocol & ";" & $Names[$Write][0] & ";" & $Cell & ";" & $Vx & ";" & @CRLF)
			Case $Cycle = 4
				$Cell = "D" & ($Write + 1)
				$Vx = "V4"
				FileWriteline($TPL, "D;" & $protocol & ";" & $Names[$Write][0] & ";" & $Cell & ";" & $Vx & ";" & @CRLF)
			Case $Cycle = 5
				$Cell = "E"  & ($Write + 1)
				$Vx = "V5"
				FileWriteline($TPL, "D;" & $protocol & ";" & $Names[$Write][0] & ";" & $Cell & ";" & $Vx & ";" & @CRLF)
			Case $Cycle = 6
				$Cell = "F" & ($Write + 1)
				$Vx = "V6"
				FileWriteline($TPL, "D;" & $protocol & ";" & $Names[$Write][0] & ";" & $Cell & ";" & $Vx & ";" & @CRLF)
			Case $Cycle = 7
				$Cell = "G"  & ($Write + 1)
				$Vx = "V7"
				FileWriteline($TPL, "D;" & $protocol & ";" & $Names[$Write][0] & ";" & $Cell & ";" & $Vx & ";" & @CRLF)
			Case $Cycle = 8
				$Cell = "H"  & ($Write + 1)
				$Vx = "V8"
				FileWriteline($TPL, "D;" & $protocol & ";" & $Names[$Write][0] & ";" & $Cell & ";" & $Vx & ";" & @CRLF)
		   EndSelect
	  next
Next

For $STD = 0 To 1 ; Write TPL file line by line for standards and blank

		For $Cycle2 = 1 to 8

		Select
			Case $Cycle2 = 1
				$Cell2 = "A" & ($STD + 11)
				$Vx2 = 20
			   FileWriteline($TPL, "D;" & $protocol & ";" & "STAND" & ";" & $Cell2 & ";" & $Vx2 & ";" & @CRLF)
			Case $Cycle2 = 2
				$Cell2 = "B" & ($STD + 11)
				$Vx2 = 40
			   FileWriteline($TPL, "D;" & $protocol & ";" & "STAND" & ";" & $Cell2 & ";" & $Vx2 & ";" & @CRLF)
			Case $Cycle2 = 3
				$Cell2 = "C" & ($STD + 11)
				$Vx2 = 80
			   FileWriteline($TPL, "D;" & $protocol & ";" & "STAND" & ";" & $Cell2 & ";" & $Vx2 & ";" & @CRLF)
			Case $Cycle2 = 4
				$Cell2 = "D" & ($STD + 11)
				$Vx = 160
			   FileWriteline($TPL, "D;" & $protocol & ";" & "STAND" & ";" & $Cell2 & ";" & $Vx2 & ";" & @CRLF)
			Case $Cycle2 = 5
				$Cell2 = "E" & ($STD + 11)
				$Vx2 = 320
			   FileWriteline($TPL, "D;" & $protocol & ";" & "STAND" & ";" & $Cell2 & ";" & $Vx2 & ";" & @CRLF)
			Case $Cycle2 = 6
				$Cell2 = "F" & ($STD + 11)
				$Vx2 = 640
			   FileWriteline($TPL, "D;" & $protocol & ";" & "STAND" & ";" & $Cell2 & ";" & $Vx2 & ";" & @CRLF)
			Case $Cycle2 = 7
				$Cell2 = "G" & ($STD + 11)
				$Vx2 = 1280
			   FileWriteline($TPL, "D;" & $protocol & ";" & "STAND" & ";" & $Cell2 & ";" & $Vx2 & ";" & @CRLF)
			Case $Cycle2 = 8
				$Cell2 = "H" & ($STD + 11)
				$Vx2 = 1
			   FileWriteline($TPL, "D;" & $protocol & ";" & "BLANC" & ";" & $Cell2 & ";" & $Vx2 & ";" & @CRLF)
		EndSelect
	Next
Next

; last line of tPL is always:  L;
FileWriteLine($TPL,"L;")

; close up shop & move files to designated folder
fileclose($TPL)
FileClose($var)
;FileMove("C:\Program Files\Tecan\EVOware\output" & "\*.TPL", "C:\apps\EVO\TPL" & "\*.TPL", 9)
;FileMove("C:\apps\EVO\CSV" & "\*.TPL", "C:\apps\EVO\TPL" & "\*.TPL", 9)
;MsgBox(4096, $var, "CSV file converted to TPL and written to C:\apps\EVO\TPL" & @CRLF & @CRLF & "")

$message = ""
$var = ""
$protocol = ""
$filename = ""
$TPL = ""
$Names 	= ""
$Write =""
$Trim = ""
$Cycle = ""
$Cycle2 = ""
$Cells = ""
$Cells2 = ""
$Vx = ""
$Vx2 = ""

Exit
; fin