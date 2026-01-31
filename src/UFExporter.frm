VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFExporter 
   Caption         =   "Export Data Range"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4785
   OleObjectBlob   =   "UFExporter.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' # ------------------------------------------------------------------------------
' # Name:        UFExporter.frm
' # Purpose:     Core UserForm for the CSV Exporter Excel VBA Add-In
' # ------------------------------------------------------------------------------
' # Name:        Exporter.bas
' # Purpose:     Helper module for launching the CSV Exporter add-in
' #
' # Original Author:  Brian Skinn
' #                     bskinn@alum.mit.edu
' #
' # Modifications:    Shep Bivens
' #   - Added tab delimiter support
' #   - Added Excel-compatible quoting modes
' #   - Improved number formatting behavior
' #   - Added choice of output encoding
' #   - UI and usability enhancements
' #
' # Created:     24 Jan 2016
' # Modified:    Jan 2026
' #
' # Copyright:
' #    (c) Brian Skinn 2016-2022
' #    Modifications (c) Shep Bivens 2026
' #
' # License:   The MIT License; see "LICENSE.txt" for full license terms.
' #
' # Original:  http://www.github.com/bskinn/excel-csvexporter
' # Fork:      https://github.com/Shep-Bivens/excel-csvexporter
' #
' # ------------------------------------------------------------------------------

Option Explicit


#If VBA7 Then
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If


' ===== EVENT-ENABLED APPLICATION =====
Private WithEvents appn As Application
Attribute appn.VB_VarHelpID = -1

' =====  Help Constants  =====
' Folder
Private Const HELP_TxBxFolder _
    = "Displays the folder where the export file is written." & vbCrLf & _
      "Click ‘Select’ to choose a folder."
Private Const HELP_LblFolder = HELP_TxBxFolder
Private Const HELP_BtnSelectFolder _
    = "Opens a folder picker to choose where the file is written."
' Filename
Private Const HELP_TxBxFilename _
    = "Name of the export file written or appended in the folder."
Private Const HELP_LblFilename = HELP_TxBxFilename
' Format
Private Const HELP_OpBtnUseCustomNumberFormat _
    = "Use the custom number format for numeric cells."
Private Const HELP_OpBtnUseCellsFormat _
    = "Use each cell's numeric format."
Private Const HELP_LblFormat _
    = "Custom Excel number format code used for numeric cells." & vbCrLf & _
      "Hold Shift and hover over the format input for help."
Private Const HELP_TxBxFormat _
    = "Search Microsoft for the following page title:" & vbCrLf & _
      """Review guidelines for customizing a number format"""
' Quote
Private Const HELP_ChBxQuote _
    = "Turn quoting on or off." & vbCrLf & _
      "On is recommended for Excel-compatible output."
Private Const HELP_OpBtnAsNeeded _
    = "Quote values containing a separator, quote," & vbCrLf & _
      "carriage return, or line feed. Matches Excel."
Private Const HELP_OpBtnQuoteAll _
    = "Quote all values, including numbers."
Private Const HELP_OpBtnQuoteNonnum _
    = "Quote only non-numeric values."
Private Const HELP_TxBxQuoteChar _
    = "Character used to surround quoted values." & vbCrLf & _
      "Standard is double quote ("")."
Private Const HELP_LblQuoteValsWith = HELP_TxBxQuoteChar
' Separator
Private Const HELP_OpBtnUseSeparator _
    = "Use the separator shown in the box." & vbCrLf & _
      "Common choices: comma (,), semicolon (;), pipe (|)."
Private Const HELP_TxBxSep _
    = "Field separator string. Usually a comma (,) for CSV files." & vbCrLf & _
      "Excel imports only a single character."
Private Const HELP_OpBtnUseTabSeparator _
    = "Use TAB as the separator between fields."
' Header
Private Const HELP_ChBxHeaderRows _
    = "Write header row(s) above the exported data range." & vbCrLf & _
      "Header columns match the selected range."
Private Const HELP_TxBxHeaderStart _
    = "First worksheet row treated as header." & vbCrLf & _
      "If blank, header starts at row 1."
Private Const HELP_TxBxHeaderStop _
    = "Last worksheet row treated as header." & vbCrLf & _
      "Must be greater than or equal to start row."
Private Const HELP_LblHeaderTo = HELP_TxBxHeaderStop
' Hidden
Private Const HELP_ChBxHiddenRows _
    = "Include hidden rows in the exported output." & vbCrLf & _
      "Unchecked skips hidden rows."
Private Const HELP_ChBxHiddenCols _
    = "Include hidden columns in the exported output." & vbCrLf & _
      "Unchecked skips hidden columns."
' Append
Private Const HELP_ChBxAppend _
    = "If checked, add new lines to the end of the file." & vbCrLf & _
      "If unchecked, overwrite the file."
' Encoding
Private Const HELP_CboEncoding _
    = "Text encoding used to write the output file." & vbCrLf & _
      "UTF-8 is recommended; BOM helps older Excel."
Private Const HELP_LblFileEncoding = HELP_CboEncoding
' Range
Private Const HELP_LblExportRgLabel _
    = "Shows current sheet, header range, and export range." & vbCrLf & _
      "Use cursor to select output range on worksheet."
Private Const HELP_LblExportRg = HELP_LblExportRgLabel
' Export
Private Const HELP_BtnExport _
    = "Write or append data to the export file." & vbCrLf & _
      "Disabled until required inputs are valid."
Private Const HELP_FrmExport = HELP_BtnExport
' Close
Private Const HELP_BtnClose _
    = "Close this window. Settings are saved for next time." & vbCrLf & _
      "Ctrl+Click shows reset options."
' Help
Private Const HELP_TxBxHelp = "Hold 'shift' and move cursor over item for help."


' =====  CONSTANTS  =====

' FileSystemObject I/O modes (match Scripting Runtime values)
Private Const FSO_ForReading  As Long = 1
Private Const FSO_ForWriting  As Long = 2
Private Const FSO_ForAppending As Long = 8
' FileSystemObject Tristate constants (match Scripting Runtime values)
Private Const FSO_TristateUseDefault As Long = -2
Private Const FSO_TristateTrue As Long = -1   ' Unicode
Private Const FSO_TristateFalse As Long = 0   ' ASCII/ANSI

Private Const SHIFT_MASK As Long = 1
Private Const CTRL_MASK  As Long = 2
Private Const ALT_MASK   As Long = 4
Private Const NoFolderStr As String = "Click ‘Select’ to choose a folder."
Private Const InvalidSelStr As String = "<invalid selection>"
Private Const NoHeaderRngStr As String = "<no header>"
Private Const BadHeaderDefStr As String = "<invalid definition>"
Private Const APP_NAME As String = "ExcelCSVExporter"
Private Const SEC_NAME As String = "Settings"
Private Const KEY_ENCODING As String = "Encoding"
Private Const KEY_SEP_MODE As String = "SeparatorMode"        ' "Custom" or "Tab"
Private Const KEY_CUSTOM_SEP As String = "CustomSeparator"    ' last custom separator string
Private Const KEY_FMT_MODE As String = "NumberFormatMode"     ' "Custom" or "Cell"
Private Const KEY_CUSTOM_FMT As String = "CustomNumberFormat" ' last custom format string


' =====  GLOBALS  =====
Private EffectiveSeparator As String
Private ExportRange As Range
Private fs As Object
Private HiddenByChart As Boolean
Private WorkFolder As Object


' =====  EVENT-ENABLED APPLICATION EVENTS  =====

Private Sub appn_SheetActivate(ByVal Sh As Object)
    ' Update the export range object, the
    ' export range reporting text, and the
    ' status of the 'Export' button any time a sheet
    ' is switched to
    
    ' Short-circuit drop-out if form is hidden, and not by
    ' navigation across a chart-sheet
    If Not HiddenByChart And Not UFExporter.Visible Then Exit Sub
    
    ' If a chartsheet is selected by the user, hide the form and
    ' sit quietly.
    If TypeOf Sh Is Chart Then
        ' Only set the hidden-by-chart flag if the form was visible
        ' when the chart was activated
        If UFExporter.Visible Then
            HiddenByChart = True
            UFExporter.Hide
        End If
        
        ' Always want to not update things when a chart sheet is selected
        Exit Sub
    Else
        ' Only need to do something special here if the form
        ' was hidden by navigation onto a chart, in which case
        ' it's desired to reset the flag and re-show the form.
        If HiddenByChart Then
            HiddenByChart = False
            UFExporter.Show
        End If
    End If
    
    setExportRange
    setExportRangeText
    setExportEnabled
    
End Sub

Private Sub appn_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    ' Update the export range object, the
    ' export range reporting text, and the
    ' status of the 'Export' button any time
    ' a new cell selection is made
    
    ' Short-circuit drop-out if form is hidden. No need to check for
    ' HiddenByChart here(?)
    If Not UFExporter.Visible Then Exit Sub
    
    setExportRange
    setExportRangeText
    setExportEnabled
    
End Sub


' =====  FORM EVENTS  =====

Private Sub BtnExport_Click()
    
    Dim filePath As String, tStrm As Object, mode As Long
    Dim errNum As Long
    Dim errDesc As String
    Dim resp As VbMsgBoxResult
    
    ' Should only ever be possible to click if form is in a good state for exporting
    SaveSettings
    
    ' Check whether separator appears in the data to be exported;
    ' advise and query user if so
    If SeparatorCanBreakCSV(ExportRange, TxBxFormat.Value, EffectiveSeparator) Then
        resp = MsgBox( _
                    "Separator is present in data to be exported!" & _
                    Chr(10) & Chr(10) & _
                    "This may cause the generated file to load incorrectly." & _
                    Chr(10) & Chr(10) & Chr(10) & _
                    "Continue with export?", _
                vbOKCancel + vbExclamation, _
                "Separator Present in Data" _
        )
        
        If resp = vbCancel Then Exit Sub
        
    End If
    
    ' Store full file path
    filePath = fs.BuildPath(WorkFolder.Path, TxBxFilename.Value)
    
    ' Convert append setting to IOMode
    If ChBxAppend.Value Then
        mode = FSO_ForAppending
    Else
        mode = FSO_ForWriting
    End If
    
    Dim useUtf8 As Boolean
    Dim keepBom As Boolean
    
    Select Case Me.CboEncoding.Value
        Case "UTF-8"
            useUtf8 = True
            keepBom = False
        Case "UTF-8 with BOM"
            useUtf8 = True
            keepBom = True
        Case Else
            useUtf8 = False
    End Select
    
    ' Warn before overwriting existing file when not appending
    If Not ChBxAppend.Value Then
        If fs.FileExists(filePath) Then
            resp = MsgBox( _
                "The file already exists:" & vbCrLf & vbCrLf & _
                filePath & vbCrLf & vbCrLf & _
                "It will be overwritten." & vbCrLf & vbCrLf & _
                "Continue?", _
                vbOKCancel + vbExclamation, _
                "Overwrite Existing File" _
            )
            If resp = vbCancel Then Exit Sub
        End If
    End If
    
    If Not useUtf8 Then
    
        ' ===== ANSI path =====
        
        Dim needsUnicode As Boolean
        needsUnicode = False
        
        If ChBxHeaderRows.Value Then
            If Not getHeaderRange Is Nothing Then
                needsUnicode = RangeRequiresUnicode(getHeaderRange, TxBxFormat.Value, EffectiveSeparator, True)
            End If
        End If
        
        If needsUnicode = False Then
            needsUnicode = RangeRequiresUnicode(ExportRange, TxBxFormat.Value, EffectiveSeparator, False)
        End If
        
        If needsUnicode Then
            MsgBox "Output contains characters that cannot be written using ANSI encoding." & _
                   Chr(10) & Chr(10) & _
                   "Select a UTF-8 encoding option and export again.", _
                   vbOKOnly + vbExclamation, _
                   "Encoding Mismatch"
            Exit Sub
        End If
        
        On Error Resume Next
        Set tStrm = fs.OpenTextFile(filePath, mode, True, FSO_TristateUseDefault)
        errNum = Err.Number: Err.Clear: On Error GoTo 0
    
        If errNum <> 0 Then
            MsgBox "File cannot be written at this location." & Chr(10) & Chr(10) & _
                   "Check if file/folder is set to read-only.", vbOKOnly + vbCritical, "Error"
            Exit Sub
        End If
    
        If ChBxHeaderRows.Value Then
            writeCSV dataRg:=getHeaderRange, tStrm:=tStrm, nFormat:=TxBxFormat.Value, _
                     Separator:=EffectiveSeparator, overrideHidden:=True
        End If
    
        writeCSV dataRg:=ExportRange, tStrm:=tStrm, nFormat:=TxBxFormat.Value, _
                 Separator:=EffectiveSeparator, overrideHidden:=False
    
        On Error Resume Next
            tStrm.Close
        errNum = Err.Number
        errDesc = Err.Description
        Err.Clear: On Error GoTo 0
        
        If errNum <> 0 Then
            ShowWriteError "Unable to close output file.", errNum, errDesc
            Exit Sub
        End If
    
    Else
        ' ===== UTF-8 path =====
        
        Dim stm As Object
        On Error Resume Next
            Set stm = OpenUtf8AdoStream(filePath, ChBxAppend.Value)
        errNum = Err.Number
        errDesc = Err.Description
        Err.Clear: On Error GoTo 0
        
        If errNum <> 0 Then
            ShowWriteError "Unable to open output file for UTF-8 writing.", errNum, errDesc
            Exit Sub
        End If
    
        ' Write header/body to the UTF-8 stream
        If ChBxHeaderRows.Value Then
            writeCSV_UTF8 dataRg:=getHeaderRange, stm:=stm, nFormat:=TxBxFormat.Value, _
                          Separator:=EffectiveSeparator, overrideHidden:=True
        End If
    
        writeCSV_UTF8 dataRg:=ExportRange, stm:=stm, nFormat:=TxBxFormat.Value, _
                      Separator:=EffectiveSeparator, overrideHidden:=False
    
        On Error Resume Next
            SaveUtf8AdoStream stm, filePath, keepBom
        errNum = Err.Number
        errDesc = Err.Description
        Err.Clear: On Error GoTo 0
        
        If errNum <> 0 Then
            ShowWriteError "Unable to save UTF-8 output file.", errNum, errDesc
            Exit Sub
        End If
    End If
    
End Sub

Private Sub BtnClose_Click()

    ' Special: Ctrl+Click Close prompts for cleanup
    If ControlDown() Then
        HandleCloseWithCleanupPrompt
        Exit Sub
    End If

    ' Normal close: save + hide
    SaveSettings
    Me.StartUpPosition = 0  ' vbStartUpManual
    Me.Hide

End Sub

Private Sub BtnSelectFolder_Click()

    Dim fd As FileDialog
    Dim result As Long, errNum As Long
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Select"
        .Title = "Choose Output Folder"
        If InStr(UCase(.InitialFileName), "SYSTEM32") Then
            .InitialFileName = Environ("USERPROFILE") & "\Documents"
        End If
        
        result = .Show
    End With
    
    ' Drop if box cancelled
    If result = 0 Then Exit Sub
    
    ' Made it here; try updating the linked folder, with error handling
    On Error Resume Next
        Set WorkFolder = fs.GetFolder(fd.SelectedItems(1))
    errNum = Err.Number: Err.Clear: On Error GoTo 0
    
    If errNum <> 0 Then
        MsgBox "Invalid folder selection", _
                vbOKOnly + vbCritical, _
                "Error"
        Exit Sub
    End If
    
    ' Update display textbox
    TxBxFolder.Value = WorkFolder.Path
    
    ' Update the Export button
    setExportEnabled

End Sub

Private Sub OpBtnAsNeeded_Click()
    setExportEnabled
End Sub

Private Sub ChBxQuote_Click()
    If ChBxQuote.Value <> False Then
        IfQuoteCharNotDoubleQuoteOutputMessage
    End If
End Sub

Private Sub ChBxHeaderRows_Change()
    ' Set the header rows box colors appropriately
    setHeaderTBxColors
    
    ' Update the range report box
    setExportRangeText
    
    ' Update Export button
    setExportEnabled
    
End Sub

Private Sub TxBxQuoteChar_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(TxBxQuoteChar.Value) = 0 Then
        TxBxQuoteChar.Value = """"
    End If
    IfQuoteCharNotDoubleQuoteOutputMessage
End Sub

Private Sub OpBtnQuoteNonnum_Click()
    setExportEnabled
End Sub

Private Sub OpBtnUseCustomNumberFormat_Click()
    setExportEnabled
End Sub

Private Sub OpBtnUseCellsFormat_Click()
    setExportEnabled
End Sub

Private Sub TxBxHeaderStart_Change()
    ' Set header rows box colors
    setHeaderTBxColors
    
    ' Update the range report box
    setExportRangeText
    
    ' Update export button
    setExportEnabled
    
End Sub

Private Sub TxBxHeaderStop_Change()
    ' Set header rows box colors
    setHeaderTBxColors
    
    ' Update the range report box
    setExportRangeText
    
    ' Update export button
    setExportEnabled
    
End Sub

Private Sub TxBxFilename_Change()

    ' If filename is nonzero-length and valid, set color black.
    ' Else, complain by setting color red
    If validFilename(TxBxFilename.Value) Then
        TxBxFilename.ForeColor = RGB(0, 0, 0)
    Else
        TxBxFilename.ForeColor = RGB(255, 0, 0)
    End If
    
    ' Update the Export button
    setExportEnabled
    
End Sub

Private Sub TxBxFormat_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(TxBxFormat.Value) = 0 Then TxBxFormat.Value = "@"
    setExportEnabled
End Sub

Private Sub OpBtnUseSeparator_Click()
    EffectiveSeparator = TxBxSep.Value
    IfSeparatorNotSingleCharOutputMessage
    setExportEnabled
End Sub

Private Sub OpBtnUseTabSeparator_Click()
    ' Tab delimiter
    EffectiveSeparator = vbTab
    setExportEnabled
End Sub

Private Sub TxBxSep_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If OpBtnUseSeparator.Value Then
        EffectiveSeparator = TxBxSep.Value
    End If
    IfSeparatorNotSingleCharOutputMessage
    setExportEnabled
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) = 0 Then ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub TxBxHelp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    TxBxHelp.Value = HELP_TxBxHelp
End Sub

Private Sub LblFolder_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_LblFolder
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub TxBxFolder_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_TxBxFolder
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub BtnSelectFolder_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_BtnSelectFolder
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub LblFilename_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_LblFilename
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub TxBxFilename_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_TxBxFilename
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub OpBtnUseCustomNumberFormat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_OpBtnUseCustomNumberFormat
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub OpBtnUseCellsFormat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_OpBtnUseCellsFormat
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub LblFormat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_LblFormat
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub TxBxFormat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_TxBxFormat
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub ChBxQuote_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_ChBxQuote
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
    If ChBxQuote.Value Then
        IfQuoteCharNotDoubleQuoteOutputMessage
    End If
    
End Sub

Private Sub OpBtnAsNeeded_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_OpBtnAsNeeded
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub OpBtnQuoteAll_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_OpBtnQuoteAll
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub OpBtnQuoteNonnum_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_OpBtnQuoteNonnum
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub LblQuoteValsWith_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_LblQuoteValsWith
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub TxBxQuoteChar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_TxBxQuoteChar
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub OpBtnUseSeparator_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_OpBtnUseSeparator
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub TxBxSep_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_TxBxSep
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub OpBtnUseTabSeparator_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_OpBtnUseTabSeparator
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub ChBxHeaderRows_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_ChBxHeaderRows
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub TxBxHeaderStart_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_TxBxHeaderStart
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub LblHeaderTo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_LblHeaderTo
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub TxBxHeaderStop_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_TxBxHeaderStop
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub ChBxHiddenRows_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_ChBxHiddenRows
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub ChBxHiddenCols_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_ChBxHiddenCols
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub ChBxAppend_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_ChBxAppend
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub LblFileEncoding_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_LblFileEncoding
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub CboEncoding_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_CboEncoding
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub LblExportRgLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_LblExportRgLabel
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub LblExportRg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_LblExportRg
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub FrmExport_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_FrmExport
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub BtnExport_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_BtnExport
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub BtnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If (Shift And SHIFT_MASK) <> 0 Then ' Shift
        TxBxHelp.Value = HELP_BtnClose
    Else ' No Shift
        TxBxHelp.Value = HELP_TxBxHelp
    End If
End Sub

Private Sub UserForm_Activate()
    ' Always update the export range info box when
    ' focus is gained, unless a show/focus attempt
    ' is made when a chart-sheet is active, in which case
    ' re-hide the form and suppress the update.
    
    If TypeOf ActiveSheet Is Chart Then
        UFExporter.Hide
        Exit Sub
    End If
    
    setExportRange
    setExportRangeText
    
End Sub

Private Sub UserForm_Initialize()
    ' Link filesystem
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Link Application for events
    Set appn = Application
    
    ' Set to no folder or file selected
    TxBxFolder.Value = NoFolderStr
    TxBxFilename.Value = ""
    
    ' Format
    OpBtnUseCellsFormat.Value = True
    TxBxFormat.Value = "@"
    
    'Quote is default to double quotes (")
    OpBtnAsNeeded.Value = True
    ChBxQuote.Value = True
    TxBxQuoteChar.Value = """"
    
    ' Comma is default separator
    TxBxSep.Value = ","
    OpBtnUseSeparator.Value = True
    EffectiveSeparator = ","
    
    ' Header Rows choices
    ChBxHeaderRows.Value = False
    TxBxHeaderStart.Value = "1"
    TxBxHeaderStop.Value = "1"
    
    'Hidden Rows and Columns
    ChBxHiddenRows.Value = False
    ChBxHiddenCols.Value = False
    
    ' Append
    ChBxAppend.Value = False
    
    ' Populate encoding choices
    With Me.CboEncoding
        .Clear
        .AddItem "ANSI"
        .AddItem "UTF-8"
        .AddItem "UTF-8 with BOM"
    End With
    CboEncoding.Value = "UTF-8"
    
    ' Export Range
    LblExportRg.Caption = ""
    
    ' Load prior choices
    LoadSettings
    
    ' Default is for filename to be empty; thus disable export button
    setExportEnabled
    
    ' Help
    TxBxHelp.Value = ""
    
End Sub


' =====  FORM MANAGEMENT ROUTINES  =====

Private Sub setExportEnabled()
    Dim sepOK As Boolean
    Dim fmtOK As Boolean

    sepOK = True
    fmtOK = (OpBtnUseCellsFormat.Value Or Len(TxBxFormat.Value) > 0)

    If ( _
        sepOK And _
        validFilename(TxBxFilename.Value) And _
        fmtOK And _
        (Not WorkFolder Is Nothing) And _
        (Not ExportRange Is Nothing) And _
        (Not (ChBxHeaderRows.Value And Not checkHeaderRowValues)) _
    ) Then
        BtnExport.Enabled = True
    Else
        BtnExport.Enabled = False
    End If
End Sub

Private Sub setExportRange()

    Dim isSheetEmpty As Boolean
    Dim rgContent As Range

    ' Detect if the sheet is actually empty
    With Selection.Parent
        If .UsedRange.Address = "$A$1" And IsEmpty(.UsedRange) Then
            isSheetEmpty = True
        Else
            isSheetEmpty = False
        End If
    End With

    If Selection.Areas.Count <> 1 Then
        Set ExportRange = Nothing

    Else
        ' Whole rows, columns, or whole sheet selected
        If Selection.Address = Selection.EntireRow.Address Or _
           Selection.Address = Selection.EntireColumn.Address Then

            If isSheetEmpty Then
                Set ExportRange = Nothing
            Else
                Set rgContent = GetContentUsedRange(Selection.Parent)

                If rgContent Is Nothing Then
                    Set ExportRange = Nothing
                Else
                    Set ExportRange = Intersect(Selection, rgContent)
                End If
            End If

        Else
            ' Explicit rectangular selection — take exactly what user selected
            Set ExportRange = Selection
        End If
    End If

    setExportEnabled
End Sub

Private Function GetContentUsedRange(ByVal ws As Worksheet) As Range
    Dim lastRowCell As Range, lastColCell As Range
    Dim lastRow As Long, lastCol As Long

    Set lastRowCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastRowCell Is Nothing Then Exit Function

    Set lastColCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                    SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)

    lastRow = lastRowCell.Row
    lastCol = lastColCell.Column

    Set GetContentUsedRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
End Function

Private Sub setExportRangeText()
    ' Helper to encapsulate setting the export range reporting text in
    ' 'LblExportRg'
    
    Dim workStr As String
    
    If Not TypeOf Selection Is Range Then Exit Sub
    
    workStr = "  Worksheet: " _
        & Selection.Parent.name _
        & Chr(10) _
        & "  Header: " _
        & getHeaderRangeAddress _
        & Chr(10) _
        & "  Range: " _
        & getExportRangeAddress
    
    LblExportRg.Caption = workStr
    
End Sub

Private Function getExportRangeAddress() As String
    ' Helper for concise generation of the export range address
    ' without dollar signs.
    '
    ' Or, if ExportRange Is Nothing, which represents some sort of
    ' selection error state, then report the invalid selection string
    
    If ExportRange Is Nothing Then
        getExportRangeAddress = InvalidSelStr
    Else
        getExportRangeAddress = ExportRange.Address(RowAbsolute:=False, _
                                                    ColumnAbsolute:=False)
    End If
    
End Function

Private Function getHeaderRangeAddress() As String
    ' Helper for concise generation of the header range address
    ' without dollar signs.
    '
    ' Or, if header export is deselected, report accordingly
    
    Dim headerRange As Range
    
    ' Store the return value from retrieving the header range
    Set headerRange = getHeaderRange
    
    If ChBxHeaderRows.Value Then
        ' Header range has to be defined in order for there to be
        ' an address to return. The validity of the header row
        ' definition in the userform is already checked
        ' within getHeaderRange, and so it doesn't need(?) to be
        ' checked again here.
        If Not headerRange Is Nothing Then
            getHeaderRangeAddress = getHeaderRange.Address( _
                        RowAbsolute:=False, ColumnAbsolute:=False _
            )
        Else
            ' Though, it's clearer to change the error message in the display box
            ' depending on whether the header definition is invalid,
            ' or if the actual range selection on ActiveSheet is bad
            If Not checkHeaderRowValues Then
                getHeaderRangeAddress = BadHeaderDefStr
            Else
                getHeaderRangeAddress = InvalidSelStr
            End If
        End If
    Else
        getHeaderRangeAddress = NoHeaderRngStr
    End If
    
End Function

Private Function getHeaderRange() As Range
    ' Helper to actually generate a reference to the header range,
    ' given the currently set export range.
    '
    ' If any of the form is in a state where the header range
    ' can't be defined, returns Nothing.
    
    Dim headerFullRows As Range
    Dim startRow As Long, stopRow As Long
    Dim errNum As Long
    
    Set getHeaderRange = Nothing
    
    If Not ChBxHeaderRows.Value Then Exit Function
    If Not checkHeaderRowValues Then Exit Function
    If ExportRange Is Nothing Then Exit Function
    
    ' Handle the case where the start value is blank (implicit start at '1')
    On Error Resume Next
        startRow = CLng(TxBxHeaderStart.Value)
    errNum = Err.Number: Err.Clear: On Error GoTo 0
    
    Select Case errNum
    Case 13
        startRow = 1
    End Select
    
    ' Stop row shouldn't(?) need special handling, given that it's already
    ' proofed by the above checks
    stopRow = CLng(TxBxHeaderStop.Value)
    
    Set headerFullRows = ExportRange.Worksheet.rows(startRow)
    Set headerFullRows = headerFullRows.Resize(stopRow - startRow + 1)
    
    Set getHeaderRange = Intersect(ExportRange.EntireColumn, headerFullRows)

End Function

Private Function checkHeaderRowValues() As Boolean
    ' Proofreads the values in the row start/stop for the header inclusion
    '
    ' True means values are ok (numbers, and start <= stop)
    ' False means something (unspecified) is wrong;
    '  could be non-numeric, or start > stop
    
    Dim errNum As Long, startRow As Long, stopRow As Long
    Dim startStr As String, stopStr As String
    
    ' Cope with empty textboxes
    If TxBxHeaderStart.Value = "" Then
        startStr = "0"
    Else
        startStr = TxBxHeaderStart.Value
    End If
    
    If TxBxHeaderStop.Value = "" Then
        stopStr = "0"
    Else
        stopStr = TxBxHeaderStop.Value
    End If
    
    ' Default to failure
    checkHeaderRowValues = False
    
    On Error Resume Next
        startRow = CInt(startStr)
        stopRow = CInt(stopStr)
    errNum = Err.Number: Err.Clear: On Error GoTo 0
    
    ' One or more non-numeric values
    If errNum <> 0 Then Exit Function
    
    ' Might as well make it so an empty start row means row 1
    startRow = Application.WorksheetFunction.Max(1, startRow)
    
    ' Value check
    If startRow > stopRow Then Exit Function
    
    ' Checks ok; return True
    checkHeaderRowValues = True
    
End Function

Private Sub setHeaderTBxColors()
    ' Helper encapsulating the color setting logic
    If checkHeaderRowValues Then
        TxBxHeaderStart.ForeColor = RGB(0, 0, 0)
        TxBxHeaderStop.ForeColor = RGB(0, 0, 0)
    Else
        TxBxHeaderStart.ForeColor = RGB(255, 0, 0)
        TxBxHeaderStop.ForeColor = RGB(255, 0, 0)
    End If
    
End Sub

' =====  HELPER FUNCTIONS  =====

Private Sub writeCSV(dataRg As Range, ByVal tStrm As Object, nFormat As String, _
                    Separator As String, overrideHidden As Boolean)
    Dim cel As Range
    Dim idxRow As Long, idxCol As Long
    Dim workStr As String, outText As String
    Dim errNum As Long

    For idxRow = 1 To dataRg.rows.Count
        If overrideHidden Or ChBxHiddenRows.Value Or Not dataRg.Cells(idxRow, 1).EntireRow.Hidden Then
            workStr = ""

            For idxCol = 1 To dataRg.Columns.Count
                If ChBxHiddenCols.Value Or Not dataRg.Cells(idxRow, idxCol).EntireColumn.Hidden Then
                    Set cel = dataRg.Cells(idxRow, idxCol)

                    outText = GetCellOutText(cel, nFormat)

                    ' --- quoting logic ---
                    If ChBxQuote.Value Then
                        If OpBtnQuoteAll.Value Or (OpBtnQuoteNonnum.Value And Not IsNumeric(cel.Value)) Then
                            ' LEGACY modes: wrap only (unchanged behavior)
                            workStr = workStr & TxBxQuoteChar.Value & outText & TxBxQuoteChar.Value

                        ElseIf OpBtnAsNeeded.Value Then
                            ' NEW: Excel-style (quote if needed + escape)
                            If NeedsQuotes_Excel(outText, Separator, TxBxQuoteChar.Value) Then
                                workStr = workStr & TxBxQuoteChar.Value & _
                                          QuoteEscape_Excel(outText, TxBxQuoteChar.Value) & _
                                          TxBxQuoteChar.Value
                            Else
                                workStr = workStr & outText
                            End If

                        Else
                            workStr = workStr & outText
                        End If
                    Else
                        workStr = workStr & outText
                    End If

                    workStr = workStr & Separator
                End If
            Next idxCol

            workStr = Left$(workStr, Len(workStr) - Len(Separator))

            On Error Resume Next
                tStrm.WriteLine workStr
            errNum = Err.Number: Err.Clear: On Error GoTo 0

            If errNum <> 0 Then
                MsgBox "Unknown error occurred while writing data line", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
        End If
    Next idxRow
End Sub

Private Sub writeCSV_UTF8(dataRg As Range, stm As Object, nFormat As String, _
                          Separator As String, overrideHidden As Boolean)

    Dim cel As Range
    Dim idxRow As Long, idxCol As Long
    Dim workStr As String, outText As String
    Dim errNum As Long

    For idxRow = 1 To dataRg.rows.Count
        If overrideHidden Or ChBxHiddenRows.Value Or Not dataRg.Cells(idxRow, 1).EntireRow.Hidden Then
            workStr = ""

            For idxCol = 1 To dataRg.Columns.Count
                If ChBxHiddenCols.Value Or Not dataRg.Cells(idxRow, idxCol).EntireColumn.Hidden Then
                    Set cel = dataRg.Cells(idxRow, idxCol)

                    outText = GetCellOutText(cel, nFormat)

                    If ChBxQuote.Value Then
                        If OpBtnQuoteAll.Value Or (OpBtnQuoteNonnum.Value And Not IsNumeric(cel.Value)) Then
                            workStr = workStr & TxBxQuoteChar.Value & outText & TxBxQuoteChar.Value

                        ElseIf OpBtnAsNeeded.Value Then
                            If NeedsQuotes_Excel(outText, Separator, TxBxQuoteChar.Value) Then
                                workStr = workStr & TxBxQuoteChar.Value & _
                                          QuoteEscape_Excel(outText, TxBxQuoteChar.Value) & _
                                          TxBxQuoteChar.Value
                            Else
                                workStr = workStr & outText
                            End If

                        Else
                            workStr = workStr & outText
                        End If
                    Else
                        workStr = workStr & outText
                    End If

                    workStr = workStr & Separator
                End If
            Next idxCol

            workStr = Left$(workStr, Len(workStr) - Len(Separator))

            On Error Resume Next
                stm.WriteText workStr & vbCrLf
            errNum = Err.Number: Err.Clear: On Error GoTo 0

            If errNum <> 0 Then
                MsgBox "Unknown error occurred while writing data line", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
        End If
    Next idxRow
End Sub

Private Function validFilename(ByVal fName As String) As Boolean
    Dim rxChrs As Object
    Set rxChrs = CreateObject("VBScript.RegExp")

    With rxChrs
        .Global = True
        .IgnoreCase = True
        .MultiLine = False
        .Pattern = "[\\/:*?""<>|]"
        validFilename = (Len(fName) >= 1 And (Not .Test(fName)))
    End With
End Function

Private Function SeparatorCanBreakCSV(ByVal dataRg As Range, ByVal nFormat As String, ByVal sep As String) As Boolean
    Dim cel As Range
    Dim outText As String
    Dim quoted As Boolean

    For Each cel In dataRg.Cells
        outText = GetCellOutText(cel, nFormat)
        quoted = CellWillBeQuoted(cel, outText, sep)
        If (Not quoted) Then
            If InStr(1, outText, sep, vbBinaryCompare) > 0 Then
                SeparatorCanBreakCSV = True
                Exit Function
            End If
        End If
    Next cel
End Function

Private Function CellWillBeQuoted(ByVal Range As Range, ByVal outText As String, _
                                  ByVal Separator As String) As Boolean
    If ChBxQuote.Value Then
        If OpBtnQuoteAll.Value _
                Or (OpBtnQuoteNonnum.Value And Not IsNumeric(Range.Value)) Then
            CellWillBeQuoted = True
        ElseIf OpBtnAsNeeded.Value Then
            If NeedsQuotes_Excel(outText, Separator, TxBxQuoteChar.Value) Then
              CellWillBeQuoted = True
            End If
        End If
    End If
End Function

Private Function OpenUtf8AdoStream(ByVal filePath As String, ByVal append As Boolean) As Object
    ' Returns an ADODB.Stream opened and ready for UTF-8 text writes.
    ' For append: loads existing file content and positions at end, then overwrites on save.
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2          ' adTypeText
    stm.Charset = "utf-8"
    stm.Open

    If append Then
        On Error Resume Next
        stm.LoadFromFile filePath
        On Error GoTo 0
        stm.Position = stm.Size
    End If

    Set OpenUtf8AdoStream = stm
End Function

Private Sub SaveUtf8AdoStream(ByVal stm As Object, ByVal filePath As String, ByVal keepBom As Boolean)
    ' Saves the ADODB.Stream to disk, overwriting. Optionally strips UTF-8 BOM.
    stm.SaveToFile filePath, 2  ' adSaveCreateOverWrite
    stm.Close

    If Not keepBom Then
        StripUtf8BomFromFile filePath
    End If
End Sub

Private Sub StripUtf8BomFromFile(ByVal filePath As String)
    ' Removes UTF-8 BOM bytes (EF BB BF) from the beginning of the file if present.
    Dim ff As Integer: ff = FreeFile
    Dim b() As Byte
    Dim n As Long

    Open filePath For Binary Access Read As #ff
    n = LOF(ff)
    If n = 0 Then
        Close #ff
        Exit Sub
    End If

    ReDim b(0 To n - 1) As Byte
    Get #ff, , b
    Close #ff

    If n < 3 Then Exit Sub
    If b(0) <> &HEF Or b(1) <> &HBB Or b(2) <> &HBF Then Exit Sub

    If n = 3 Then
        ' File contains only BOM; write empty file.
        Dim stmEmpty As Object
        Set stmEmpty = CreateObject("ADODB.Stream")
        stmEmpty.Type = 1 ' adTypeBinary
        stmEmpty.Open
        stmEmpty.SaveToFile filePath, 2 ' overwrite
        stmEmpty.Close
        Exit Sub
    End If

    Dim b2() As Byte
    ReDim b2(0 To n - 4) As Byte

    Dim i As Long
    For i = 3 To n - 1
        b2(i - 3) = b(i)
    Next i

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' adTypeBinary
    stm.Open
    stm.Write b2
    stm.SaveToFile filePath, 2 ' overwrite (truncates)
    stm.Close
End Sub

Private Sub ShowWriteError(ByVal actionDesc As String, ByVal errNum As Long, ByVal errDesc As String)
    MsgBox actionDesc & Chr(10) & Chr(10) & _
           "Error " & errNum & ": " & errDesc, _
           vbOKOnly + vbCritical, _
           "File Write Error"
End Sub

Private Sub LoadSettings()

    ' ----- Folder and File Name -----
    Dim folderPath As String
    folderPath = GetSetting(APP_NAME, SEC_NAME, "FolderPath", "")
    If Len(folderPath) > 0 Then
        Dim errNum As Long
        On Error Resume Next
            Set WorkFolder = fs.GetFolder(folderPath)
        errNum = Err.Number: Err.Clear: On Error GoTo 0
        If errNum = 0 Then
            TxBxFolder.Value = WorkFolder.Path
        Else
            TxBxFolder.Value = NoFolderStr
        End If
    Else
        TxBxFolder.Value = NoFolderStr
    End If
    TxBxFilename.Value = GetSetting(APP_NAME, SEC_NAME, "Filename", "")

    ' ----- Format -----
    TxBxFormat.Value = GetSetting(APP_NAME, SEC_NAME, KEY_CUSTOM_FMT, "@")
    Dim fmtMode As String
    fmtMode = GetSetting(APP_NAME, SEC_NAME, KEY_FMT_MODE, "Cell")
    If fmtMode = "Custom" Then
        OpBtnUseCustomNumberFormat.Value = True
    Else
        OpBtnUseCellsFormat.Value = True
    End If
    
    ' ----- Quote -----
    ChBxQuote.Value = CBool(Val(GetSetting(APP_NAME, SEC_NAME, "QuoteEnabled", "1")))
    TxBxQuoteChar.Value = GetSetting(APP_NAME, SEC_NAME, "QuoteChar", """")
    Dim qMode As String
    qMode = GetSetting(APP_NAME, SEC_NAME, "QuoteMode", "AsNeeded")
    Select Case qMode
        Case "All":      OpBtnQuoteAll.Value = True
        Case "Nonnum":   OpBtnQuoteNonnum.Value = True
        Case "AsNeeded": OpBtnAsNeeded.Value = True
        Case Else:       OpBtnAsNeeded.Value = True
    End Select

    ' ----- Separator -----
    TxBxSep.Value = GetSetting(APP_NAME, SEC_NAME, KEY_CUSTOM_SEP, ",")
    Dim sepMode As String
    sepMode = GetSetting(APP_NAME, SEC_NAME, KEY_SEP_MODE, "Custom")
    If sepMode = "Tab" Then
        OpBtnUseTabSeparator.Value = True
        EffectiveSeparator = vbTab
    Else
        OpBtnUseSeparator.Value = True
        EffectiveSeparator = TxBxSep.Value
    End If
    
    ' ----- Header -----
    ChBxHeaderRows.Value = CBool(Val(GetSetting(APP_NAME, SEC_NAME, "HeaderRows", "0")))
    TxBxHeaderStart.Value = GetSetting(APP_NAME, SEC_NAME, "HeaderStart", "1")
    TxBxHeaderStop.Value = GetSetting(APP_NAME, SEC_NAME, "HeaderStop", "1")
    
    ' ----- Hidden -----
    ChBxHiddenRows.Value = CBool(Val(GetSetting(APP_NAME, SEC_NAME, "HiddenRows", "0")))
    ChBxHiddenCols.Value = CBool(Val(GetSetting(APP_NAME, SEC_NAME, "HiddenCols", "0")))
    
    ' ----- Append -----
    ChBxAppend.Value = CBool(Val(GetSetting(APP_NAME, SEC_NAME, "Append", "0")))

    ' ----- Encoding -----
    Dim enc As String
    enc = GetSetting(APP_NAME, SEC_NAME, KEY_ENCODING, "UTF-8")
    Select Case enc
        Case "ANSI", "UTF-8", "UTF-8 with BOM"
            Me.CboEncoding.Value = enc
        Case Else
            Me.CboEncoding.Value = "UTF-8"
    End Select

End Sub

Private Sub SaveSettings()
    
    ' ----- Folder & File Name -----
    If WorkFolder Is Nothing Then
        SaveSetting APP_NAME, SEC_NAME, "FolderPath", ""
    Else
        SaveSetting APP_NAME, SEC_NAME, "FolderPath", WorkFolder.Path
    End If
    SaveSetting APP_NAME, SEC_NAME, "Filename", TxBxFilename.Value

    ' ----- Number format persistence -----
    SaveSetting APP_NAME, SEC_NAME, KEY_CUSTOM_FMT, TxBxFormat.Value
    If OpBtnUseCustomNumberFormat.Value Then
        SaveSetting APP_NAME, SEC_NAME, KEY_FMT_MODE, "Custom"
    Else
        SaveSetting APP_NAME, SEC_NAME, KEY_FMT_MODE, "Cell"
    End If

    ' ----- Quote -----
    SaveSetting APP_NAME, SEC_NAME, "QuoteEnabled", CStr(Abs(CInt(ChBxQuote.Value)))
    SaveSetting APP_NAME, SEC_NAME, "QuoteChar", TxBxQuoteChar.Value
    If OpBtnQuoteAll.Value Then
        SaveSetting APP_NAME, SEC_NAME, "QuoteMode", "All"
    ElseIf OpBtnQuoteNonnum.Value Then
        SaveSetting APP_NAME, SEC_NAME, "QuoteMode", "Nonnum"
    ElseIf OpBtnAsNeeded.Value Then
        SaveSetting APP_NAME, SEC_NAME, "QuoteMode", "AsNeeded"
    Else
        SaveSetting APP_NAME, SEC_NAME, "QuoteMode", "AsNeeded"
    End If

    ' ----- Separator persistence -----
    SaveSetting APP_NAME, SEC_NAME, KEY_CUSTOM_SEP, TxBxSep.Value
    If OpBtnUseTabSeparator.Value Then
        SaveSetting APP_NAME, SEC_NAME, KEY_SEP_MODE, "Tab"
    Else
        SaveSetting APP_NAME, SEC_NAME, KEY_SEP_MODE, "Custom"
    End If

    ' ----- Header -----
    SaveSetting APP_NAME, SEC_NAME, "HeaderRows", CStr(Abs(CInt(ChBxHeaderRows.Value)))
    SaveSetting APP_NAME, SEC_NAME, "HeaderStart", TxBxHeaderStart.Value
    SaveSetting APP_NAME, SEC_NAME, "HeaderStop", TxBxHeaderStop.Value

    ' ----- Hidden -----
    SaveSetting APP_NAME, SEC_NAME, "HiddenRows", CStr(Abs(CInt(ChBxHiddenRows.Value)))
    SaveSetting APP_NAME, SEC_NAME, "HiddenCols", CStr(Abs(CInt(ChBxHiddenCols.Value)))

    ' ----- Append -----
    SaveSetting APP_NAME, SEC_NAME, "Append", CStr(Abs(CInt(ChBxAppend.Value)))
    
    ' ----- Encoding selector -----
    SaveSetting APP_NAME, SEC_NAME, KEY_ENCODING, Me.CboEncoding.Value
End Sub


Private Function RangeRequiresUnicode(ByVal dataRg As Range, ByVal nFormat As String, _
                                      ByVal Separator As String, ByVal includeHidden As Boolean) As Boolean

    Dim idxRow As Long, idxCol As Long
    Dim cel As Range
    Dim workStr As String, outText As String

    For idxRow = 1 To dataRg.rows.Count
        If includeHidden Or ChBxHiddenRows.Value Or Not dataRg.Cells(idxRow, 1).EntireRow.Hidden Then
            workStr = ""

            For idxCol = 1 To dataRg.Columns.Count
                If ChBxHiddenCols.Value Or Not dataRg.Cells(idxRow, idxCol).EntireColumn.Hidden Then
                    Set cel = dataRg.Cells(idxRow, idxCol)

                    outText = GetCellOutText(cel, nFormat)

                    If ChBxQuote.Value Then
                        If OpBtnQuoteAll.Value Or (OpBtnQuoteNonnum.Value And Not IsNumeric(cel.Value)) Then
                            workStr = workStr & TxBxQuoteChar.Value & outText & TxBxQuoteChar.Value

                        ElseIf OpBtnAsNeeded.Value Then
                            If NeedsQuotes_Excel(outText, Separator, TxBxQuoteChar.Value) Then
                                workStr = workStr & TxBxQuoteChar.Value & _
                                          QuoteEscape_Excel(outText, TxBxQuoteChar.Value) & _
                                          TxBxQuoteChar.Value
                            Else
                                workStr = workStr & outText
                            End If

                        Else
                            workStr = workStr & outText
                        End If
                    Else
                        workStr = workStr & outText
                    End If

                    workStr = workStr & Separator
                End If
            Next idxCol

            workStr = Left$(workStr, Len(workStr) - Len(Separator))

            If StrConv(StrConv(workStr, vbFromUnicode), vbUnicode) <> workStr Then
                RangeRequiresUnicode = True
                Exit Function
            End If
        End If
    Next idxRow
End Function

Private Function NeedsQuotes_Excel(ByVal s As String, ByVal sep As String, _
        ByVal q As String) As Boolean
    NeedsQuotes_Excel = (InStr(1, s, sep, vbBinaryCompare) > 0) _
                     Or (InStr(1, s, q, vbBinaryCompare) > 0) _
                     Or (InStr(1, s, vbCr, vbBinaryCompare) > 0) _
                     Or (InStr(1, s, vbLf, vbBinaryCompare) > 0)
End Function

Private Function QuoteEscape_Excel(ByVal s As String, ByVal q As String) As String
    QuoteEscape_Excel = Replace(s, q, q & q)
End Function

Private Function GetCellOutText(ByVal cel As Range, ByVal exporterFmt As String) As String
    Dim v As Variant

    If Len(CStr(cel)) > 250 Then
        GetCellOutText = CStr(cel.Value)
        Exit Function
    End If

    v = cel.Value2

    ' Treat Empty, "", and errors safely
    If IsError(v) Then
        GetCellOutText = CStr(v)   ' or "" if you prefer
        Exit Function
    End If

    If IsEmpty(v) Or (VarType(v) = vbString And Len(v) = 0) Then
        GetCellOutText = ""
        Exit Function
    End If

    If IsNumeric(v) Then
        If OpBtnUseCellsFormat.Value Then
            GetCellOutText = FormatNumericLikeExcel(cel, exporterFmt)
        Else
            GetCellOutText = Format$(v, exporterFmt)
        End If
    Else
        GetCellOutText = CStr(v)
    End If
End Function

Private Function FormatNumericLikeExcel(ByVal cel As Range, ByVal fallbackFmt As String) As String
    ' Use Excel’s TEXT() formatting engine for cell formats.
    ' This matches Excel’s CSV output more closely than VBA Format(),
    ' and does not depend on column width like cel.Text can.

    On Error GoTo UseFallback

    Dim fmt As String
    fmt = cel.NumberFormat

    ' WorksheetFunction.Text can throw for some formats/values; fall back safely.
    FormatNumericLikeExcel = Application.WorksheetFunction.Text(cel.Value, fmt)
    Exit Function

UseFallback:
    ' Fallback to provided custom format (or "@"/"General" depending on your design)
    FormatNumericLikeExcel = Format$(cel.Value, fallbackFmt)
End Function

Private Function ControlDown() As Boolean
    ' Returns True if Control key is down right now.
    ' (GetKeyState is a light WinAPI call; safe and common in VBA.)
#If VBA7 Then
    Dim s As Integer
    s = GetKeyState(vbKeyControl)
    ControlDown = (s < 0)
#Else
    Dim s As Integer
    s = GetKeyState(vbKeyControl)
    ControlDown = (s < 0)
#End If
End Function

Private Sub HandleCloseWithCleanupPrompt()
    ' Returns True if it closed (either cleanup or normal save),
    ' False if user cancelled and wants to return to the form.

    Dim resp As VbMsgBoxResult
    resp = MsgBox( _
        "Close options:" & vbCrLf & vbCrLf & _
        "Yes  = Remove all saved settings and close" & vbCrLf & _
        "No   = Save settings and close" & vbCrLf & _
        "Cancel = Return to the exporter", _
        vbYesNoCancel + vbQuestion, _
        "Close / Reset Settings")

    Select Case resp
        Case vbYes
            ' Remove all settings for this app
            On Error Resume Next
                DeleteSetting APP_NAME
            On Error GoTo 0

            ' Close without saving anything
            Me.StartUpPosition = 0
            Unload Me

        Case vbNo
            ' Save settings and close normally
            SaveSettings
            Me.StartUpPosition = 0
            Me.Hide
        
        Case Else
            ' Cancel: do nothing
    End Select
End Sub

Public Sub RefreshForShow()
    ' Revalidate folder selection (it can become invalid while form is hidden)
    RevalidateWorkFolder

    ' Refresh selection-dependent state
    setExportRange
    setExportRangeText
    setExportEnabled
End Sub

Private Sub RevalidateWorkFolder()
    Dim p As String
    Dim errNum As Long

    If WorkFolder Is Nothing Then
        TxBxFolder.Value = NoFolderStr
        Exit Sub
    End If

    p = WorkFolder.Path

    ' Fast path: folder still exists
    If Len(Dir$(p, vbDirectory)) > 0 Then
        TxBxFolder.Value = p
        Exit Sub
    End If

    ' Folder missing; try to rebind once (handles some network/share cases)
    On Error Resume Next
        Set WorkFolder = fs.GetFolder(p)
    errNum = Err.Number: Err.Clear: On Error GoTo 0

    If errNum = 0 Then
        TxBxFolder.Value = WorkFolder.Path
    Else
        ' Invalidate selection
        Set WorkFolder = Nothing
        TxBxFolder.Value = NoFolderStr
    End If
End Sub

Private Sub IfQuoteCharNotDoubleQuoteOutputMessage()
    If TxBxQuoteChar.Value <> """" Then
        MsgBox _
            "Excel recognizes only the double quote ("") as a text qualifier." & vbCrLf & _
            "Other quote characters may not import correctly.", _
            vbOKOnly + vbExclamation, "Quote Character"
    End If
End Sub

Private Sub IfSeparatorNotSingleCharOutputMessage()
    If Len(TxBxSep.Value) <> 1 Then
            MsgBox "Separator is not a single character." & vbCrLf & _
                   "This exporter will still write the file, but Excel’s delimiter import" & vbCrLf & _
                   "normally uses a single character (comma, tab, semicolon, etc.).", _
                   vbOKOnly + vbExclamation, "Separator Warning"
    End If
End Sub


