Attribute VB_Name = "Module7_LabEval_FinalRpt"
Global PTStudy As String
Global Season As String
Global Year As String

Sub test1()

PTStudy = Worksheets("Study Details").Range("B1")
Season = Worksheets("Study Details").Range("B2")
Year = Worksheets("Study Details").Range("B3")

'get aqc
Dim daoDB            As DAO.Database
Dim daoQueryDef      As DAO.QueryDef
Dim daoRcd           As DAO.Recordset
Dim refpath          As String

'adds a new sheet to store the path information for each individual Lab Evaluation
ActiveWorkbook.Sheets.Add

'This is the name of the Query in the AQC database
ActiveSheet.Name = "Query List"
    
'Path to AQC database in PT Folder on ptstudies login
refpath = "M:\FP PT Program-IQM\PT Studies\AQCaddressdb\AQCAddressdb2005.mdb"

Set daoDB = OpenDatabase(refpath)
Set daoQueryDef = daoDB.QueryDefs("qry-final report email")
Set daoRcd = daoQueryDef.OpenRecordset
ActiveWorkbook.Worksheets("Query List").Range("A1").CopyFromRecordset daoRcd

'This is the folder that will contain the laboratory evaluations by AQC reference number generated by the Lab Eval Macro
MsgBox ("Select the folder containing all laboratory evaluations")
Dim foldername As String
With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = "M:\My Documents\"
    .AllowMultiSelect = False
    .Title = "Select the folder containing all laboratory evaluations"
    If .Show = False Then Exit Sub
    foldername = .SelectedItems(1) & "\"
    DoEvents
End With

'this will search the selected folder for the Lab Evaluations and link each lab on the email list to a Lab Eval document path
lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
For i = 1 To lastrow
    Cells(i, 4) = foldername & "Laboratory Evaluations " & Cells(i, 3).Value & ".pdf"
Next i

'send mail
For j = 1 To lastrow
Dim OutApp As Object, eval As String, sendto As String, ccto As String
Dim strpath, strFilter, strFile, strname As String
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    eval = Cells(j, 4).Value
    sendto = Cells(j, 1)
    ccto = Cells(j, 2)
    Application.ScreenUpdating = False
    
    On Error Resume Next
      
    With OutMail
        
        .To = sendto
        .CC = ccto
        .BCC = ""
        'modify this line to change the email subject
        .Subject = "ECCC PT Study #" & PTStudy & ": Laboratory Evaluation & Final Report"
        
        
'This is the HTML code for the formatting and content of the email
'By default it is set to sign each email with Fedelinas signature
'To change the signature please look to the line of code below that starts with " & <br><br><br><b>Fedelina DeOliveira"
'From this line to the very last line, anyones information can be input in place for Fedelinas
        .HTMLBody = "<font face = calibri (body)>Dear Study Participant,<br><br>Attached to this e-mail is the <b>Proficiency Appraisal</b> and <b>Z-Score Summary</b> " _
        & "for each program in which your laboratory participated, as well as the <b>Final Report</b> for PT <b>#" & PTStudy & "</b>.<br><br>The <b>Laboratory Evaluation</b> " _
        & "is named by the AQC reference number assigned to your laboratory, not your confidential lab code.  " _
        & "This electronic file replaces the mailed hard copy evaluations which were provided for previous studies.<br><br>" _
        & "Text files containing study statistics are available on request, as they are no longer provided as e-mail attachments automatically.<br><br>" _
        & "Please contact <a href =mailto:ec.ptstudies.ec@canada.ca>ec.ptstudies.ec@canada.ca</a> if you have any questions regarding the laboratory evaluation or final report.</font><br>" _
        & "<br><br><br><b>Fedelina DeOliveira</b><br><br>RM Technologist<br>Information and Quality Management / Water S&T Directorate/ S&T Branch<br>" _
        & "Environment and Climate Change Canada / Government of Canada<br><a href =mailto:ec.ptstudies.ec@canada.ca>ec.ptstudies.ec@canada.ca</a> / Tel: 905-336-4942 / fax. : 905-336-8914<br><br>" _
        & "Gestion de l'information et de la qualit� /Direction des sciences et de la technologie de l'eau/Branche des S&T<br>Environnement et Changement Climatique Canada / Gouvernement du Canada<br>" _
        & "<a href =mailto:ec.ptstudies.ec@canada.ca>ec.ptstudies.ec@canada.ca</a> / T�l. : 905-336-4942 / Telecopieur. : 905-336-8914" & .HTMLBody
        
        .Attachments.Add (eval)
        
        strpath = "O:\FP PT Program\PT Final Reports\PT 0" & PTStudy & "\"
        strFilter = "*.pdf"
        strname = "ECCC PT"
        strFile = Dir(strpath & strname & strFilter)
        .Attachments.Add (strpath & strFile)
        .Save
        
' ***Unhighlight to automatically send emails
'        .Send

    End With

    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
Next j

'deletes the now redundant sheet containing the paths to individual laboratory evaluations
MsgBox ("Select the 'DELETE' button in the next pop-up window")
Sheets("query List").Delete

End Sub