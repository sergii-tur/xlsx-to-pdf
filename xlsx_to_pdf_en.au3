; #INDEX# =======================================================================================================================
; Title .........: xlsx_to_pdf
; AutoIt Version : 3.3.14.5
; Description ...: Converts all xlsx files from script directory to pdf files. 
;                  After conversion, names of pdf files = xlsx files' names. 
;                  There is a possibility to append string to pdf files' names.
; Author(s) .....: sergii.developer@gmail.com
; ===============================================================================================================================
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <FontConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <Excel.au3>

StartScript()

Func StartScript()
  ;------------------------------------GUI-------------------------------------;
  GUICreate("Converter xlsx -> pdf", 300, 160, @DesktopWidth / 2 - 160, @DesktopHeight / 2 - 160, -1, $WS_EX_ACCEPTFILES)
  
  Local $idLabel = GUICtrlCreateLabel("pdf file name" & @CRLF & "For example _month-year", 20, 20, 200, 50)
  GUICtrlSetFont ( $idLabel, 10)
  
  GUICtrlSetState(-1, $GUI_DROPACCEPTED)
  Local $idInput = GUICtrlCreateInput("_01-2022", 20, 70, 150, 20) ; will not accept drag&drop files
  
  Local $idBtn = GUICtrlCreateButton("Convert", 100, 100, 100)
  
  Local $idCredits = GUICtrlCreateLabel("sergii.developer@gmail.com", 150, 140, 150, 50)
  GUICtrlSetFont ( $idCredits, 8)

  GUISetState(@SW_SHOW)
  ;-----------------------------------/GUI-------------------------------------;

  ; Loop until the user exits.
  While 1
    Switch GUIGetMsg()
      Case $GUI_EVENT_CLOSE
        ExitLoop
      Case $idBtn
        Local $sInput = GUICtrlRead($idInput)
        Convert($sInput)
        ExitLoop
    EndSwitch
  WEnd
EndFunc ;==> StartScript

Func Convert($sNameTemplate)
  Local $aFileList = _FileListToArray(@ScriptDir, "*xlsx", "*xls")
  If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "", "Error: file path not found")
    Exit
  EndIf
  If @error = 4 Then
    MsgBox($MB_SYSTEMMODAL, "", "Error: no xlsx files in a folder")
    Exit
  EndIf
  If $aFileList[0] == 0 Then
    MsgBox($MB_SYSTEMMODAL, "", "Error: no xlsx files in a folder")
    Exit
  EndIf
  
  ProgressOn("Converting to pdf", "Please wait...", "0%", -1, -1, BitOR($DLG_NOTONTOP, $DLG_MOVEABLE))
  Local $iConverts = 0
  Local $iNumOfFiles = $aFileList[0]
  Local $iBarStep = 100 / $iNumOfFiles
  Local $iBar = 0

  Local $oExcel = _Excel_Open(False, False, False, True, False)
  
  For $i = $aFileList[0] To 1 Step -1
    $iBar += $iBarStep
    ProgressSet($iBar, Round($iBar) & "%")
    $sStr = $aFileList[$i]
    ; omitting hidden xlsx files whith preceding "~" char
    if StringInStr($sStr, "~") == 0 Then
      $iConverts += 1
      Local $sFile = StringTrimRight($aFileList[$i], 5)
      Local $sXlsxName = $sFile & ".xlsx"
      Local $sPdfName = $sFile & $sNameTemplate & ".pdf"
      Local $sWorkbook = @ScriptDir & "\" & $sXlsxName
      
      ;MsgBox($MB_SYSTEMMODAL, "Path to  file", $sWorkbook)
      
      If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_Export Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
      Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook, False, False, Default, Default, 0)
      ;Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)
      If @error Then
          MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_Export Example", "Error opening workbook " & $sWorkbook & @CRLF & "@error = " & @error & ", @extended = " & @extended)
          _Excel_Close($oExcel)
          Exit
      EndIf

      ; Export the whole workbook as PDF.
      Local $sOutput = @ScriptDir & "\" & $sPdfName
      _Excel_Export($oExcel, $oWorkbook, $sOutput)
      If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_Export Example 2", "Error saving the workbook to '" & $sOutput & "'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
    EndIf
  Next
  
  ProgressSet(100, "Success", "Done")
  Sleep(2000)
  ProgressOff() ; Close the progress window.
  
  MsgBox($MB_SYSTEMMODAL, "Conversion completed", $iConverts & " files have been converted to pdf")
  ;_Excel_Close($oExcel, False, True)
  _Excel_BookClose($oExcel)
EndFunc ;==> Convert

; Function _ExcelBookClose
; Description ...: Closes the active workbook and removes the specified Excel object.
; Author ........: SEO <locodarwin at yahoo dot com>
; Modified.......: 07/17/2008 by bid_daddy; litlmike
Func _ExcelBookClose($oExcel, $fSave = 1, $fAlerts = 0)
  If Not IsObj($oExcel) Then Return SetError(1, 0, 0)
  Local $sObjName = ObjName($oExcel)
  If $fSave > 1 Then $fSave = 1
  If $fSave < 0 Then $fSave = 0
  If $fAlerts > 1 Then $fAlerts = 1
  If $fAlerts < 0 Then $fAlerts = 0
  ; Save the users specified settings
  Local $fDisplayAlerts = $oExcel.Application.DisplayAlerts
  Local $fScreenUpdating = $oExcel.Application.ScreenUpdating
  ; Make necessary changes
  $oExcel.Application.DisplayAlerts = $fAlerts
  $oExcel.Application.ScreenUpdating = $fAlerts
  Switch $sObjName
    Case "_Workbook"
      If $fSave Then $oExcel.Save()
      ; Check if multiple workbooks are open
      ; Do not close application if there are
      If $oExcel.Application.Workbooks.Count > 1 Then
          $oExcel.Close
          $oExcel = ''
          $oExcel = ObjGet("", "Excel.Application")
          ; Restore the users specified settings
          $oExcel.Application.DisplayAlerts = $fDisplayAlerts
          $oExcel.Application.ScreenUpdating = $fScreenUpdating
          $oExcel = ''
      Else
          $oExcel.Application.Quit
      EndIf
    Case "_Application"
      If $fSave Then $oExcel.ActiveWorkBook.Save()
      $oExcel.Quit()
    Case Else
      Return SetError(1, 0, 0)
  EndSwitch
  Return 1
EndFunc ;==> _ExcelBookClose