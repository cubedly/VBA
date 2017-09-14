Attribute VB_Name = "LabEvaluationsMacroV2"
Sub Step1()

'Selects folder to import files
MsgBox ("Select the folder containing reports as .doc")
Dim foldername As String
With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = "M:\My Documents\"
    .AllowMultiSelect = False
    .Title = "Select the folder containing reports as .doc"
    If .Show = False Then Exit Sub
    foldername = .SelectedItems(1) & "\"
    DoEvents
End With
'end

'Selects folder to export files
MsgBox ("Select the folder to export Final Laboratory Evaluations")
Dim mypathfinal As String
With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = "M:\My Documents\"
    .AllowMultiSelect = False
    .Title = "Select the folder to export Laboratory Evaluations"
    If .Show = False Then Exit Sub
    mypathfinal = .SelectedItems(1) & "\"
    DoEvents
End With
'end

'Selects folder to export files
MsgBox ("Select the folder to export subsequent documents")
Dim mypath As String
With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = "M:\My Documents\"
    .AllowMultiSelect = False
    .Title = "Select the folder to export other documents"
    If .Show = False Then Exit Sub
    mypath = .SelectedItems(1) & "\"
    DoEvents
End With

MsgBox "Export Start"

'Loop to run split document macro on all files in the folder
Dim filepathname As String, NextFile As String
NextFile = Dir(foldername & "*.doc", vbNormal)
Do While NextFile <> ""
    filepathname = foldername & NextFile
    Documents.Open FileName:=filepathname
    Call Step2(mypath) 'runs split document macro for the open file
    ActiveDocument.Close SaveChanges:=False
    NextFile = Dir()
Loop

Call Step3(mypath)

Dim ext As String, filext As String
ext = "docx"
filext = mypath & "\AllDocs.pdf"
Call Step5(filext, ext, mypath, mypathfinal)

MsgBox ("Every step has been completed successfully!")
MsgBox ("The Final Laboratory Evaluations can be found in the folder located at " & mypathfinal)

End Sub
Sub Step2(mypath)
Attribute Step2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.ExportReporttoPDF"

'Defining loop variables
Dim i As Integer, j As Integer, k As Integer

Dim l As Integer 'l = length of array
l = 0
Dim findl As Word.Range
Set findl = ActiveDocument.Content
Do While findl.Find.Execute(FindText:="Your laboratory code is", Forward:=True) = True
    l = l + 1
Loop

'arrays
Dim lcpgab() As Integer '(i,1) stores lab codes, (i,2) stores page number of lab codes, (i,3) stores a/b
    ReDim lcpgab(1 To l, 1 To 3)



'populate array with lab codes and page number
i = 1
Dim rng As Word.Range
Set rng = ActiveDocument.Content
Do While rng.Find.Execute(FindText:="Your laboratory code is", Forward:=True) = True
    lcpgab(i, 2) = rng.Information(wdActiveEndPageNumber)
    
    rng.SetRange Start:=rng.End + 2, End:=rng.End + 5
    lcpgab(i, 1) = rng.Text
    
    rng.SetRange Start:=rng.End, End:=rng.End + 1
    lcpgab(i, 3) = 0
    If rng.Text = "a" Then lcpgab(i, 3) = 1
    If rng.Text = "b" Then lcpgab(i, 3) = 2
    If rng.Text = "c" Then lcpgab(i, 3) = 3
    If rng.Text = "d" Then lcpgab(i, 3) = 4
    If rng.Text = "e" Then lcpgab(i, 3) = 5
    If rng.Text = "f" Then lcpgab(i, 3) = 6
    If rng.Text = "g" Then lcpgab(i, 3) = 7
    If rng.Text = "h" Then lcpgab(i, 3) = 8
    
    rng.Collapse wdCollapseEnd
    i = i + 1
Loop


'Parameter
Dim pm As Word.Range, pmname As String
Set pm = ActiveDocument.Content

If pm.Find.Execute(FindText:="major ions", Forward:=True) = True Then
    pmname = "MI"
End If

If pm.Find.Execute(FindText:="sediment", Forward:=True) = True Then
    pmname = "SED"
End If

If pm.Find.Execute(FindText:="trace elements in water", Forward:=True) = True Then
    pmname = "TM"
End If

If pm.Find.Execute(FindText:="total phosphorus", Forward:=True) = True Then
    pmname = "TP"
End If

If pm.Find.Execute(FindText:="turbidity", Forward:=True) = True Then
    pmname = "TU"
End If

If pm.Find.Execute(FindText:="for rain", Forward:=True) = True Then
    pmname = "RN"
End If

If pm.Find.Execute(FindText:="mercury in water", Forward:=True) = True Then
    pmname = "HG"
End If

If pm.Find.Execute(FindText:="mercury in water-low level", Forward:=True) = True Then
    pmname = "HGLL"
End If

    
 'APP or Z
Dim appz As Word.Range, appzname As String
Set appz = ActiveDocument.Content

If appz.Find.Execute(FindText:="Laboratory Proficiency Appraisal", Forward:=True) = True Then
    appzname = "APP"
End If

If appz.Find.Execute(FindText:="Score Summary", Forward:=True) = True Then
    appzname = "Z"
End If


'Export to pdf
Dim a As Integer, b As Integer

Dim lczero  As String
Dim lcab    As String
Dim pg2     As String


a = 1

For i = 1 To l
    'zeros
    If lcpgab(i, 1) < 100 Then lczero = "0"
    If lcpgab(i, 1) < 10 Then lczero = "00"
    
    'lcab
    If lcpgab(i, 3) = 0 Then lcab = "0"
    If lcpgab(i, 3) = 1 Then lcab = "a"
    If lcpgab(i, 3) = 2 Then lcab = "b"
    If lcpgab(i, 3) = 3 Then lcab = "c"
    If lcpgab(i, 3) = 4 Then lcab = "d"
    If lcpgab(i, 3) = 5 Then lcab = "e"
    If lcpgab(i, 3) = 6 Then lcab = "f"
    If lcpgab(i, 3) = 7 Then lcab = "g"
    If lcpgab(i, 3) = 8 Then lcab = "h"

    If i = 1 Then
        b = lcpgab(i + 1, 2) - 1
        ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            mypath & "F" & lczero & lcpgab(i, 1) & "_" & pmname & "_" & lcab & "_" & appzname & pg2 & ".pdf", _
            ExportFormat:=wdExportFormatPDF, _
            Range:=wdExportFromTo, From:=a, To:=b
        
        'MsgBox "F" & lcpgab(i, 1) & " pg=" & pg2 & " a=" & a & " b=" & b
        
        a = b + 1
    End If

    If i <> 1 And i <> l Then
        b = lcpgab(i + 1, 2) - 1
        If lcpgab(i, 1) = lcpgab(i - 1, 1) And lcpgab(i, 3) = lcpgab(i - 1, 3) Then pg2 = "_page2"
        ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            mypath & "F" & lczero & lcpgab(i, 1) & "_" & pmname & "_" & lcab & "_" & appzname & pg2 & ".pdf", _
            ExportFormat:=wdExportFormatPDF, _
            Range:=wdExportFromTo, From:=a, To:=b
        
        'MsgBox "F" & lcpgab(i, 1) & " pg=" & pg2 & " a=" & a & " b=" & b
        
        a = b + 1
    End If
    
    If i = l Then
        b = ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
        If lcpgab(i, 1) = lcpgab(i - 1, 1) And lcpgab(i, 3) = lcpgab(i - 1, 3) Then pg2 = "_page2"
        ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            mypath & "F" & lczero & lcpgab(i, 1) & "_" & pmname & "_" & lcab & "_" & appzname & pg2 & ".pdf", _
            ExportFormat:=wdExportFormatPDF, _
            Range:=wdExportFromTo, From:=a, To:=b
         
        'MsgBox "F" & lcpgab(i, 1) & " pg=" & pg2 & " a=" & a & " b=" & b
    
    End If
    
    
    lcab = ""
    lczero = ""
    pg2 = ""
Next i

End Sub

Sub Step3(mypath)
     
    Const DestFile As String = "AllDocs.pdf"
     
    Dim MyFiles As String, mypath2 As String
    Dim a() As String, i As Long, f As String
     
    mypath2 = mypath
     ' Populate the array a() by PDF file names
    If Right(mypath2, 1) <> "\" Then mypath2 = mypath2 & "\"
    ReDim a(1 To 2 ^ 14)
    f = Dir(mypath2 & "*.pdf")
    While Len(f)
        If StrComp(f, DestFile, vbTextCompare) Then
            i = i + 1
            a(i) = f
        End If
        f = Dir()
    Wend
     
     ' Merge PDFs
    If i Then
        ReDim Preserve a(1 To i)
        MyFiles = Join(a, ",")
        Application.StatusBar = "Merging, please wait ..."
        Call Step4(mypath2, MyFiles, DestFile)
        Application.StatusBar = False
    Else
        MsgBox "No PDF files found in" & vbLf & mypath2, vbExclamation, "Canceled"
    End If
     
End Sub
 
Sub Step4(mypath3 As String, MyFiles As String, Optional DestFile As String = "MergedFile.pdf")
     
    Dim a As Variant, i As Long, n As Long, ni As Long, p As String
    Dim AcroApp As New Acrobat.AcroApp, PartDocs() As Acrobat.CAcroPDDoc
     
    If Right(mypath3, 1) = "\" Then p = mypath3 Else p = mypath3 & "\"
    a = split(MyFiles, ",")
    ReDim PartDocs(0 To UBound(a))
     
    On Error GoTo exit_
    If Len(Dir(p & DestFile)) Then Kill p & DestFile
    For i = 0 To UBound(a)
         ' Check PDF file presence
        If Dir(p & Trim(a(i))) = "" Then
            MsgBox "File not found" & vbLf & p & a(i), vbExclamation, "Canceled"
            Exit For
        End If
         ' Open PDF document
        Set PartDocs(i) = CreateObject("AcroExch.PDDoc")
        PartDocs(i).Open p & Trim(a(i))
        If i Then
             ' Merge PDF to PartDocs(0) document
            ni = PartDocs(i).GetNumPages()
            If Not PartDocs(0).InsertPages(n - 1, PartDocs(i), 0, ni, True) Then
                MsgBox "Cannot insert pages of" & vbLf & p & a(i), vbExclamation, "Canceled"
            End If
             ' Calc the number of pages in the merged document
            n = n + ni
             ' Release the memory
            PartDocs(i).Close
            Set PartDocs(i) = Nothing
        Else
             ' Calc the number of pages in PartDocs(0) document
            n = PartDocs(0).GetNumPages()
        End If
    Next
     
    If i > UBound(a) Then
         ' Save the merged document to DestFile
        If Not PartDocs(0).Save(PDSaveFull, p & DestFile) Then
            MsgBox "Cannot save the resulting document" & vbLf & p & DestFile, vbExclamation, "Canceled"
        End If
    End If
     
exit_:
     
     ' Inform about error/success
    If Err Then
        MsgBox Err.Description, vbCritical, "Error #" & Err.Number
    ElseIf i > UBound(a) Then
    End If
     
     ' Release the memory
    If Not PartDocs(0) Is Nothing Then PartDocs(0).Close
    Set PartDocs(0) = Nothing
     
     ' Quit Acrobat application
    AcroApp.Exit
    Set AcroApp = Nothing
     
End Sub

Sub Step5(PDFPath As String, FileExtension As String, mypath4 As String, mypathfinal As String)
    'Saves a PDF file as another format using Adobe Professional.
   
    Dim objAcroApp      As Acrobat.AcroApp
    Dim objAcroAVDoc    As Acrobat.AcroAVDoc
    Dim objAcroPDDoc    As Acrobat.AcroPDDoc
    Dim objJSO          As Object
    Dim boResult        As Boolean
    Dim ExportFormat    As String
    Dim NewFilePath     As String
   
    'Check if the file exists.
    If Dir(PDFPath) = "" Then
        MsgBox "Cannot find the PDF file!" & vbCrLf & "Check the PDF path and retry.", _
                vbCritical, "File Path Error"
        Exit Sub
    End If
   
    'Check if the input file is a PDF file.
    If LCase(Right(PDFPath, 3)) <> "pdf" Then
        MsgBox "The input file is not a PDF file!", vbCritical, "File Type Error"
        Exit Sub
    End If
   
    'Initialize Acrobat by creating App object.
    Set objAcroApp = CreateObject("AcroExch.App")
   
    'Set AVDoc object.
    Set objAcroAVDoc = CreateObject("AcroExch.AVDoc")
   
    'Open the PDF file.
    boResult = objAcroAVDoc.Open(PDFPath, "")
       
    'Set the PDDoc object.
    Set objAcroPDDoc = objAcroAVDoc.GetPDDoc
   
    'Set the JS Object - Java Script Object.
    Set objJSO = objAcroPDDoc.GetJSObject
   
    'Check the type of conversion.
    Select Case LCase(FileExtension)
        Case "eps": ExportFormat = "com.adobe.acrobat.eps"
        Case "html", "htm": ExportFormat = "com.adobe.acrobat.html"
        Case "jpeg", "jpg", "jpe": ExportFormat = "com.adobe.acrobat.jpeg"
        Case "jpf", "jpx", "jp2", "j2k", "j2c", "jpc": ExportFormat = "com.adobe.acrobat.jp2k"
        Case "docx": ExportFormat = "com.adobe.acrobat.docx"
        Case "doc": ExportFormat = "com.adobe.acrobat.doc"
        Case "png": ExportFormat = "com.adobe.acrobat.png"
        Case "ps": ExportFormat = "com.adobe.acrobat.ps"
        Case "rft": ExportFormat = "com.adobe.acrobat.rft"
        Case "xlsx": ExportFormat = "com.adobe.acrobat.xlsx"
        Case "xls": ExportFormat = "com.adobe.acrobat.spreadsheet"
        Case "txt": ExportFormat = "com.adobe.acrobat.accesstext"
        Case "tiff", "tif": ExportFormat = "com.adobe.acrobat.tiff"
        Case "xml": ExportFormat = "com.adobe.acrobat.xml-1-00"
        Case Else: ExportFormat = "Wrong Input"
    End Select
    
    'Check if the format is correct and there are no errors.
    If ExportFormat <> "Wrong Input" And Err.Number = 0 Then
        '***'
        NewFilePath = mypath4
        '***'
        'Save PDF file to the new format.
        boResult = objJSO.SaveAs(mypath4 & "\allpages.docx", "com.adobe.acrobat.docx")

        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
        
        'Close the Acrobat application.
        boResult = objAcroApp.Exit
         
    Else
       
        'Something went wrong, so close the PDF file and the application.
       
        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
       
        'Close the Acrobat application.
        boResult = objAcroApp.Exit
       
        'Inform the user that something went wrong.
        MsgBox "Something went wrong!" & vbNewLine & "The conversion of the following PDF file FAILED:" & _
        vbNewLine & PDFPath, vbInformation, "Conversion failed"

    End If
       
    'Release the objects.
    Set objAcroPDDoc = Nothing
    Set objAcroAVDoc = Nothing
    Set objAcroApp = Nothing
    
    Call Step6(mypath4, mypathfinal)
       
End Sub

Sub Step6(mypath5 As String, mypathfinal As String)

Dim allpagesfile    As String
Dim folderpath      As String
    
allpagesfile = mypath5 & "\allpages.docx"

folderpath = mypathfinal

'Dim looping variables
Dim i As Integer, j As Integer, k As Integer

'Create new excel doc
Dim wbBook      As Workbook
Dim wsSheet     As Worksheet
Set wbBook = Workbooks.Add(xlWBATWorksheet)
Set wsSheet = wbBook.Worksheets("Sheet1")

'get aqc
Dim daoDB            As DAO.Database
Dim daoQueryDef      As DAO.QueryDef
Dim daoRcd           As DAO.Recordset
Dim refpath          As String
    
'change here
refpath = "M:\FP PT Program-IQM\PT Studies\AQCaddressdb\AQCAddressdb2005.mdb"
'refpath = "O:\Student Hiring\Student Projects\Lily Li\Evaluation email distribution macro\AQCAddressdb2005.mdb"

Set daoDB = OpenDatabase(refpath)
Set daoQueryDef = daoDB.QueryDefs("ReferenceQuery")
Set daoRcd = daoQueryDef.OpenRecordset
ActiveWorkbook.Worksheets(1).Range("A1").CopyFromRecordset daoRcd

Dim sizear As Integer
sizear = Cells(Rows.Count, "B").End(xlUp).Row

Dim refno() As Variant
ReDim refno(1 To sizear, 1 To 2)

For i = 1 To sizear 'populate refno array with data from aqc query
    refno(i, 1) = Sheets(1).Cells(i, 1).Value
    refno(i, 2) = Sheets(1).Cells(i, 2).Value
Next i

ActiveWorkbook.Close SaveChanges:=False

Documents.Open FileName:=allpagesfile

'l = total number lab codes found
Dim l As Integer

l = 0

Dim findl As Word.Range
Set findl = ActiveDocument.Content

Do While findl.Find.Execute(FindText:="Your laboratory code is", Forward:=True) = True
    l = l + 1
Loop

'arrays
Dim refpg() As Integer
    ReDim refpg(1 To l, 1 To 2)

Dim lc() As String
    ReDim lc(1 To l)

'populate array with lab codes and page number
i = 1
Dim rng As Word.Range
Set rng = ActiveDocument.Content

Do While rng.Find.Execute(FindText:="Your laboratory code is", Forward:=True) = True
    
    refpg(i, 2) = rng.Information(wdActiveEndPageNumber)
    rng.SetRange Start:=rng.End + 1, End:=rng.End + 5
    lc(i) = rng.Text
    
        For j = 1 To sizear
            If refno(j, 1) = lc(i) Then refpg(i, 1) = refno(j, 2)
        Next j
   
    rng.Collapse wdCollapseEnd
    i = i + 1
    
Loop

'Export to pdf
Dim a As Integer, b As Integer

a = 1

For i = 1 To l
    If i <> l Then
    
        If refpg(i, 1) <> refpg(i + 1, 1) Then
            If refpg(i + 1, 2) - refpg(i, 2) = 2 Then b = refpg(i, 2) + 1 Else b = refpg(i, 2)
            ActiveDocument.ExportAsFixedFormat OutputFileName:= _
                folderpath & "Laboratory Evaluations " & refpg(i, 1) & ".pdf", _
                ExportFormat:=wdExportFormatPDF, _
                Range:=wdExportFromTo, From:=a, To:=b
            a = b + 1
        End If
    End If
    If i = l Then
        b = ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
        ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            folderpath & "Laboratory Evaluations " & refpg(i, 1) & ".pdf", _
            ExportFormat:=wdExportFormatPDF, _
            Range:=wdExportFromTo, From:=a, To:=b
    End If
Next i

ActiveDocument.Close (wdDoNotSaveChanges)

End Sub
