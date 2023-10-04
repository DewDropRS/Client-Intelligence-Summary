Attribute VB_Name = "cis_module111"
Option Explicit

Public RprtPckgCount As Integer
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private MstrCsvPath As String
Private MstrRprtDir As String
Private MstrTemplateDir As String
Private MstrCsvGrpdDir As String
Private MstrFinRprtDir As String
Private MstrLogDir As String
Public Function CsvPath() As String

    If MstrCsvPath = vbNullString Then
        MstrCsvPath = "C:\Client Intelligence Summary\data\csv_xgrouped\"
        'MstrCsvPath = "\\bsc\sftp\SFTP146\cis\prod\reports\"
    End If
    CsvPath = MstrCsvPath
    
End Function
Public Function LogDir() As String
    If MstrLogDir = vbNullString Then
        MstrLogDir = "C:\Client Intelligence Summary\logs\"
    End If
    LogDir = MstrLogDir

End Function
Public Function RprtDir() As String

    If MstrRprtDir = vbNullString Then
        MstrRprtDir = "C:\Client Intelligence Summary\reports\"
    End If
    RprtDir = MstrRprtDir

End Function
Public Function TemplateDir() As String

    If MstrTemplateDir = vbNullString Then
        MstrTemplateDir = "C:\Client Intelligence Summary\templates\"
    End If
    
    TemplateDir = MstrTemplateDir

End Function
Public Function CsvGrpdDir() As String

    If MstrCsvGrpdDir = vbNullString Then
        MstrCsvGrpdDir = "C:\Client Intelligence Summary\data\csv_grouped\"
    End If
    
    CsvGrpdDir = MstrCsvGrpdDir

End Function
Public Function FinRprtDir() As String

    If MstrFinRprtDir = vbNullString Then
        MstrFinRprtDir = "C:\Client Intelligence Summary\reports\"
    End If
    
    FinRprtDir = MstrFinRprtDir

End Function
Sub cis_clear_data()

    Clear_All_Files_And_SubFolders_In_Folder ("C:\Client Intelligence Summary\data\csv_xgrouped")
    Clear_All_Files_And_SubFolders_In_Folder ("C:\Client Intelligence Summary\data\csv_grouped")
    Clear_All_Files_And_SubFolders_In_Folder ("C:\Client Intelligence Summary\reports")
    
End Sub
Sub Clear_All_Files_And_SubFolders_In_Folder(MyPath As String)
'Delete all files and subfolders
'Be sure that no file is open in the folder
    Dim FSO As Object

    Set FSO = CreateObject("scripting.filesystemobject")

    If Right(MyPath, 1) = "\" Then
        MyPath = Left(MyPath, Len(MyPath) - 1)
    End If

    If FSO.FolderExists(MyPath) = False Then
        MsgBox MyPath & " doesn't exist"
        Exit Sub
    End If

    On Error Resume Next
    'Delete files
    FSO.deletefile MyPath & "\*.*", True

    'Delete subfolders
    FSO.deletefolder MyPath & "\*.*", True
    On Error GoTo 0

End Sub
Sub Clear_Reports_Folder(MyPath As String)
'Delete all files and subfolders
'Be sure that no file is open in the folder
    Dim FSO As Object
    'Dim MyPath As String

    Set FSO = CreateObject("scripting.filesystemobject")

    If Right(MyPath, 1) = "\" Then
        MyPath = Left(MyPath, Len(MyPath) - 1)
    End If

    If FSO.FolderExists(MyPath) = False Then
        MsgBox MyPath & " doesn't exist"
        Exit Sub
    End If

    On Error Resume Next
    'Delete files
    FSO.deletefile MyPath & "\*.pdf", True

    'Delete subfolders
    FSO.deletefolder MyPath & "\*.*", True
    On Error GoTo 0

End Sub
Sub createTextFile(dirfilestr As String)
    Dim fs As Object
    Dim a As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.fileExists(dirfilestr) = False Then
        Set a = fs.createTextFile(dirfilestr, True)
        a.Close
    End If
End Sub
Public Function DiffHrMinSec(strtdtm As Date, enddtm As Date) As String

    Dim intSeconds, intHours, intMinutes As Integer
    Dim timediff As String
    
    intSeconds = DateDiff("s", strtdtm, enddtm)
    intHours = intSeconds \ 3600
    intMinutes = intSeconds \ 60 Mod 60
    intSeconds = intSeconds Mod 60
    
    DiffHrMinSec = intHours & ":" & intMinutes & ":" & intSeconds
    
End Function
Function fileExists(s_directory As String, s_fileName As String) As Boolean
    Dim obj_fso As Object

    Set obj_fso = CreateObject("Scripting.FileSystemObject")
    fileExists = obj_fso.fileExists(s_directory & "\" & s_fileName)

End Function
Private Function lastrow8() As Long
    lastrow8 = 1
    If Application.WorksheetFunction.CountA(ActiveSheet.Cells) <> 0 Then
        lastrow8 = ActiveSheet.Cells.Find(What:="*", After:=ActiveSheet.Range("A1"), Lookat:=xlPart, _
            LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row + 1
    End If
End Function
Sub MakeDirectory(targetdir)
    Dim fldr, MyFolder  As String
    fldr = Dir(targetdir, vbDirectory)
    If Len(fldr) > 0 Then
    Else
      MkDir targetdir
    End If
End Sub
Sub MessageBoxTimer(msgstring As String)
    Const AckTime = 5
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after 5 seconds
    Select Case InfoBox.Popup(msgstring & " " & vbCrLf & "(this window closes automatically after " & AckTime & " seconds).", AckTime, "CIS Message Box")
        Case 1, -1
            Exit Sub
    End Select

End Sub
Sub MoveFiles()
    Dim theFile As String
    theFile = Dir("\\bsc\sftp\SFTP146\cis\prod\reports\*.*")
    Do Until theFile = ""
        Name "\\bsc\sftp\SFTP146\cis\prod\reports\" & theFile As "C:\Client Intelligence Summary\data\csv_xgrouped\" & theFile
        theFile = Dir
    Loop
    
End Sub
Sub SelectMultipleSheets()
    Dim sheetCount As Integer
    Dim arSheets() As String
    Dim intArrayIndex As Integer
    
    intArrayIndex = 0

    For sheetCount = 1 To Sheets.Count

        If Sheets(sheetCount).Name <> "Sheet1" Then
            Worksheets(Sheets(sheetCount).Name).Activate
            If (ActiveSheet.Range("BJ2").Value = "print") Then
                'UpdateFooter
                ReDim Preserve arSheets(intArrayIndex)
                arSheets(intArrayIndex) = Sheets(sheetCount).Name
                intArrayIndex = intArrayIndex + 1
            End If
        End If
    Next

    Sheets(arSheets).Select

End Sub
Sub Send_Email()

    Dim Email_Subject, Email_Send_From, Email_Send_To, Email_Cc, Email_Bcc, Email_Body As String
    Dim Mail_Object, Mail_Single As Variant
    
    Email_Subject = "CIS VBA Application has completed"
    Email_Send_To = "karyn.wang@blueshieldca.com"
    Email_Cc = " "
    Email_Bcc = " "
    Email_Body = "Client Intelligence Summary Report VBA application has completed." & vbCrLf & "The report/s can be found in the following directory." & vbCrLf & "\\bsc\fin\hdr\Client Intelligence Summary\reports"
    On Error GoTo debugs
    
    Set Mail_Object = CreateObject("Outlook.Application")
    Set Mail_Single = Mail_Object.CreateItem(0)
    
    With Mail_Single
        .Subject = Email_Subject
        .To = Email_Send_To
        .cc = Email_Cc
        .BCC = Email_Bcc
        .Body = Email_Body
        .send
    End With
debugs:
    If Err.Description <> "" Then MsgBox Err.Description
    
End Sub
Sub storeFolderPaths(path As Variant, ByRef PathArray() As Variant, ByRef FolderArray() As Variant, ByRef PathCount As Integer)
    Dim FirstDir As String
    FirstDir = Dir(path, vbDirectory)
    PathCount = 0
    Do While FirstDir <> ""
            If Not FirstDir Like ".*" Then
                If (GetAttr(path & FirstDir) And vbDirectory) = vbDirectory Then
                    PathCount = PathCount + 1
                    ReDim Preserve PathArray(1 To PathCount)
                    ReDim Preserve FolderArray(1 To PathCount)
                    PathArray(PathCount) = path & FirstDir & "\"
                    FolderArray(PathCount) = FirstDir
                End If
            End If
            FirstDir = Dir()
    Loop
End Sub
Sub TimePerformanceWbFile(dirfilestr As String, subName As String, stime As Variant, etime As Variant, dur As String, fcount As Integer)

    Dim fs As Object
    Dim a As Object
    Dim lastrow As Long
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.fileExists(dirfilestr) = False Then
    
        Workbooks.Add
        ActiveWorkbook.Sheets("Sheet1").Cells(1, 1).Value = "Sub Name"
        ActiveWorkbook.Sheets("Sheet1").Cells(1, 2).Value = "Start Time"
        ActiveWorkbook.Sheets("Sheet1").Cells(1, 3).Value = "End Time"
        ActiveWorkbook.Sheets("Sheet1").Cells(1, 4).Value = "Duration"
        ActiveWorkbook.Sheets("Sheet1").Cells(1, 5).Value = "File Count"
        
        ActiveWorkbook.Sheets("Sheet1").Cells(2, 1).Value = subName
        ActiveWorkbook.Sheets("Sheet1").Cells(2, 2).Value = stime
        ActiveWorkbook.Sheets("Sheet1").Cells(2, 3).Value = etime
        ActiveWorkbook.Sheets("Sheet1").Cells(2, 4).Value = dur
        ActiveWorkbook.Sheets("Sheet1").Cells(2, 5).Value = fcount
        
        ActiveWorkbook.SaveAs FileName:=dirfilestr
    
    Else
    
        Workbooks.Open FileName:=dirfilestr, ReadOnly:=False
        
        lastrow = lastrow8()
        
        ActiveWorkbook.Sheets("Sheet1").Cells(lastrow, 1).Value = subName
        ActiveWorkbook.Sheets("Sheet1").Cells(lastrow, 2).Value = stime
        ActiveWorkbook.Sheets("Sheet1").Cells(lastrow, 3).Value = etime
        ActiveWorkbook.Sheets("Sheet1").Cells(lastrow, 4).Value = dur
        ActiveWorkbook.Sheets("Sheet1").Cells(lastrow, 5).Value = fcount
    
        ActiveWorkbook.Save
        ActiveWorkbook.Close


    End If
    
End Sub
Sub UpdateFooter()
   ActiveSheet.PageSetup.CenterFooter = Range("BH2").Value & Range("BI2").Value
End Sub
Sub Step1_copy_csv_files()
    
    Dim FileArray() As Variant
    Dim fileNoExtArray() As String
    Dim PathArray() As String
    Dim CsvPathArray() As String
    Dim SheetNameArray() As String
    Dim FileCount As Integer
    Dim p, i As Integer
    Dim csvFileName, WBname, CustDir, Entity, SHname, wbfile As String
    Dim CsvP, RprtD, CsvGrpdD, LgDir As String
    
    Dim starttime, endtime As Variant
    Dim elapstm As String
    
    Dim WBnameArray() As String
    Dim WBnameWrdCt As Integer
    
    
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False

        starttime = DateTime.Now
        
        'Initialize paths
        CsvP = CsvPath
        RprtD = RprtDir
        CsvGrpdD = CsvGrpdDir
        LgDir = LogDir
                              
        csvFileName = Dir(CsvP & "*.csv")
        
        'count the number of files processed
        FileCount = 0

        'store csv filenames to FileArray with extension, store csv filenames without extension, store csv full paths, store excel full paths for creating directories
        Do While csvFileName <> ""
        
                
                WBname = Replace(csvFileName, ".csv", "")
                WBnameArray() = Split(WBname, "_")
                WBnameWrdCt = UBound(WBnameArray()) + 1
                
                If WBnameWrdCt = 2 Then
                
                    FileCount = FileCount + 1
                    ReDim Preserve FileArray(1 To FileCount)
                    FileArray(FileCount) = csvFileName
    
                    'store filenames without csv extension
                    ReDim Preserve fileNoExtArray(1 To FileCount)
                    fileNoExtArray(FileCount) = WBname
                    
                    'store csv file paths
                    ReDim Preserve CsvPathArray(1 To FileCount)
                    CsvPathArray(FileCount) = CsvP & csvFileName
    
                    'store future excel folder paths as target directory\customerNumber_ReportName (e.g. 3505300_Active_PPO)
                    ReDim Preserve PathArray(1 To FileCount)
                    Entity = WBnameArray(1)
                    PathArray(FileCount) = CsvGrpdD & Entity & "\"
                    
                    'store strings used to rename sheets in the xlsx files
                    ReDim Preserve SheetNameArray(1 To FileCount)
                    SheetNameArray(FileCount) = WBnameArray(0)
                    
                End If
                
                csvFileName = Dir
            
        Loop
        
        'remove duplicates from pathArray
        Dim eleArr_1, x, UB
        Dim PathArray2() As String
        
        x = 0
        For Each eleArr_1 In PathArray
        
                If (Not PathArray2) = -1 Then
                    ReDim Preserve PathArray2(x)
                    PathArray2(x) = eleArr_1
                    x = x + 1
                Else
                
                    UB = UBound(Filter(PathArray2, eleArr_1))
                    If UBound(Filter(PathArray2, eleArr_1)) = -1 Then
                        ReDim Preserve PathArray2(x)
                        PathArray2(x) = eleArr_1
                        x = x + 1
        
                    End If
                End If
                
        Next

        'Create folders in the target directories from pathArray2() created above
        p = 0
        Do While p <= UBound(PathArray2)
            CustDir = PathArray2(p)
            MakeDirectory (CustDir)
        p = p + 1
        Loop
                
        'loop through csv files and save as xlsx files to target directories in excel_workbooks, rename sheets to acctinfo or expdata
        i = 1
        
        Do While i <= FileCount
            
            FileCopy CsvP & FileArray(i), PathArray(i) & fileNoExtArray(i) & ".csv"
                                              
            i = i + 1
            
        Loop
            
        endtime = DateTime.Now
        
        elapstm = DiffHrMinSec(CDate(starttime), CDate(endtime))
        
        'write to vba time performance file
        wbfile = LgDir & "cis_vba_time_performance.xlsx"

        TimePerformanceWbFile dirfilestr:=wbfile, subName:="Step1_copy_csv_files", stime:=starttime, etime:=endtime, dur:=elapstm, fcount:=FileCount

        Application.DisplayAlerts = True
        Application.ScreenUpdating = True

        'MessageBoxTimer ("Step1_Save_csv_as_xlsx" & " (Duration: " & elapstm & ")" & vbCrLf & "Start Time: " & starttime & vbCrLf & " End Time: " & endtime & vbCrLf & "There were " & FileCount & " file(s) found and processed.")
          
End Sub
Sub Step2_Merge_CSVdata_to_Template()
    Dim SelData As Range
    Dim FileArray() As Variant
    
    Dim srcPathArray() As Variant
    Dim srcFolderArray() As Variant
    
    Dim FileCount As Integer
    Dim FileName, folderName, FileDir, FileSearch, CustDir, PkgDir, WBname, SheetName As String
    Dim RprtD, TmpltD, CsvGrpdD, LogD As Variant

    Dim starttime, endtime As Variant
    Dim i, p, srcPathCount As Integer
    
    
    Dim currentSheet As Worksheet
    Dim sheetIndex As Integer
    Dim sh As Worksheet, wb As Workbook
    Dim srcSh As Worksheet
    Dim textfile As String
    Dim fs, f As Object
    Const ForAppending = 8
    Dim elapstm As String
    
    Dim wbfile As String
    Dim lastrow As Long
    Dim TxtRng  As Range
    
    Dim ShNameArray() As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    starttime = DateTime.Now
    
   'Initialize Paths
    CsvGrpdD = CsvGrpdDir
    RprtD = RprtDir
    TmpltD = TemplateDir
    LogD = LogDir

    
    storeFolderPaths path:=CsvGrpdD, _
                        PathArray:=srcPathArray(), _
                        FolderArray:=srcFolderArray(), _
                        PathCount:=srcPathCount
    
    i = 1
    FileCount = 0
    Do While i <= srcPathCount
    
        
        FileDir = srcPathArray(i)
        'PkgDir = RprtD & srcFolderArray(i)
        PkgDir = RprtD
        FileSearch = "*.csv"
        
        FileName = Dir(FileDir & FileSearch)
        Workbooks.Open FileName:=TmpltD & "CIS template.xlsx"
        Set wb = Workbooks("CIS template.xlsx")
        
        Do While FileName <> ""
         FileCount = FileCount + 1
         Workbooks.Open FileName:=FileDir & FileName, ReadOnly:=True
            For Each sh In ActiveWorkbook.Sheets
               'clear anything on clipboard to maximize available memory
                Application.CutCopyMode = False
                'new-copy/paste data instead of copying entire sheet to template workbook
                ShNameArray() = Split(ActiveSheet.Name, "_")
                SheetName = ShNameArray(0) & "_data"
                Set srcSh = ActiveSheet

                wb.Worksheets(SheetName).Range("A1", "O45").ClearContents
            
                srcSh.Range(Cells(1, 1), Cells.SpecialCells(xlCellTypeLastCell)).Select
                
                Selection.Copy
                
                wb.Worksheets(SheetName).Range("A1").PasteSpecial Paste:=xlPasteValues
               
                'sh.Copy After:=wb.Sheets(wb.Sheets.Count)
            Next sh

            Workbooks(FileName).Close SaveChanges:=False
            FileName = Dir()
        Loop
        ActiveWorkbook.Worksheets("Cover").Select
        ActiveWorkbook.SaveAs FileName:=PkgDir & "CIS_" & srcFolderArray(i), FileFormat:=51
        ActiveWorkbook.Close SaveChanges:=True
        i = i + 1
    Loop
    endtime = DateTime.Now
    elapstm = DiffHrMinSec(CDate(starttime), CDate(endtime))
    
    wbfile = LogD & "cis_vba_time_performance.xlsx"
    
    TimePerformanceWbFile dirfilestr:=wbfile, subName:="Step2_Merge_CSVdata_to_Template", stime:=starttime, etime:=endtime, dur:=elapstm, fcount:=FileCount
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    'MessageBoxTimer ("Step2_Merge_DataWbks_to_Template" & " (Duration: " & elapstm & ")" & vbCrLf & "Start Time: " & starttime & vbCrLf & " End Time: " & endtime & vbCrLf & "There were " & FileCount & " file(s) found and processed.")

End Sub
Sub Step3_ProcessReportPackage()

    Dim srcPathArray() As Variant
    Dim srcFolderArray() As Variant
    Dim FileSearch, FileName, FileDir As String
    Dim i, srcPathCount As Integer
    Dim FileCount As Integer
    Dim printStr As Range
    Dim starttime, endtime As Variant
    Dim elapstm As String
    Dim yrmo, acctname, custnum, lob, level, psk, fnreporttype, OrigSheet As String
    
    Dim textfile As String
    Dim fs, f As Object
    Const ForAppending = 8
    
    Dim wbfile As String
    Dim lastrow As Long
    Dim TxtRng  As Range
    
    Dim sharray As Variant
    
    Dim RprtD, FinRprtD, LogD As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    starttime = DateTime.Now
    
    'Initialize paths
    RprtD = RprtDir
    FinRprtD = FinRprtDir
    LogD = LogDir

    'storeFolderPaths path:=RprtD, _
    '                    PathArray:=srcPathArray(), _
    '                    FolderArray:=srcFolderArray(), _
    '                    PathCount:=srcPathCount
                        

    
    'i = 1
    'FileCount = 0
    'Do While i <= srcPathCount
        'FileDir = srcPathArray(i)
    
        FileSearch = "*.xlsx"
          
        FileName = Dir(RprtD & FileSearch)
        
        
        Do While FileName <> ""
            FileCount = FileCount + 1
            Workbooks.Open FileName:=RprtD & FileName, ReadOnly:=True
            ActiveWorkbook.RefreshAll
            

            yrmo = Worksheets("companyinfo_data").Cells(6, "B").Value
            acctname = Worksheets("companyinfo_data").Cells(5, "B").Value
            custnum = Worksheets("companyinfo_data").Cells(4, "B").Value
            lob = Worksheets("companyinfo_data").Cells(7, "B").Value
            level = Worksheets("companyinfo_data").Cells(17, "B").Value
            psk = Worksheets("companyinfo_data").Cells(27, "B").Value

            If level = "Level 2" Then
                sharray = Array("Cover EXP HCC", "Experience Level II", "Hight Cost Claimants", "Notes Glossary EXP HCC")
                fnreporttype = "_CIS_EXP_HCC_"
            Else
                sharray = Array("Cover EXP HCC", "Hight Cost Claimants", "Notes Glossary EXP HCC")
                fnreporttype = "_CIS_HCC_"

            End If

            OrigSheet = ActiveSheet.Name
            ActiveWorkbook.Sheets(sharray).Select
            ActiveSheet.ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    FileName:=FinRprtD & yrmo & "_" & acctname & "_" & custnum & fnreporttype & lob & "_" & "PSK" & psk, _
                    Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, _
                    IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False
            Worksheets(OrigSheet).Select
                
            SelectMultipleSheets
            ActiveSheet.ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    FileName:=FinRprtD & yrmo & "_" & acctname & "_" & custnum & "_Client_Intelligence_Summary_" & lob & "_" & "PSK" & psk, _
                    Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, _
                    IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False
                    Workbooks(FileName).Close SaveChanges:=False
                        
            FileName = Dir()
        Loop
        'i = i + 1
    'Loop
    
    endtime = DateTime.Now
    elapstm = DiffHrMinSec(CDate(starttime), CDate(endtime))
  
    wbfile = LogD & "cis_vba_time_performance.xlsx"
    
    TimePerformanceWbFile dirfilestr:=wbfile, subName:="Step3_ProcessReportPackage", stime:=starttime, etime:=endtime, dur:=elapstm, fcount:=FileCount

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    'MessageBoxTimer ("Step3_ProcessReportPackage" & " (Duration: " & elapstm & ")" & vbCrLf & "Start Time: " & starttime & vbCrLf & " End Time: " & endtime & vbCrLf & "There were " & FileCount & " file(s) found and processed.")
    RprtPckgCount = FileCount
    
End Sub
Sub CIS_Generate_Reports()

    Dim startGenRprts, endGenRprts As Variant
    Dim textfile, elapsTmGenRprts As String
    Dim fs, f As Object
    
    Dim wbfile As String
    Dim lastrow As Long
    Dim TxtRng  As Range
    
    Dim LgDir As String
    Const ForAppending = 8
    
    LgDir = LogDir

    startGenRprts = DateTime.Now
    
    Step1_copy_csv_files
    Step2_Merge_CSVdata_to_Template
    Step3_ProcessReportPackage
    
    endGenRprts = DateTime.Now
    elapsTmGenRprts = DiffHrMinSec(CDate(startGenRprts), CDate(endGenRprts))
    
    MessageBoxTimer ("CIS_Generate_Reports" & " (Duration: " & elapsTmGenRprts & ")" & vbCrLf & "Start Time: " & startGenRprts & vbCrLf & " End Time: " & endGenRprts)


    wbfile = LgDir & "cis_vba_time_performance.xlsx"
    
    TimePerformanceWbFile dirfilestr:=wbfile, subName:="CIS_Generate_Reports", stime:=startGenRprts, etime:=endGenRprts, dur:=elapsTmGenRprts, fcount:=RprtPckgCount
        
    Send_Email

End Sub


