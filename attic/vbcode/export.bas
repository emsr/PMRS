Attribute VB_Name = "Module2"
'******************************************************************************
'
' *** COPYRIGHT:  Copyright 1999
'                 IIT Research Institute
'                 US Govt retains rights in accordance
'                 with DoD FAR Supp 252.227 - 7013.
'
'-----------------------------------------------------------------------
'
' *** CODING HISTORY:
'
'   DATE      PROGRAMMER                     DESCRIPTION
' --------   ------------                   -------------
' 4/21/99    Mike Fahler                    PMRS v. 1.00
'******************************************************************************

'variable to determine type of export file (e.g., access,foxpro, etc.)
Global DataType As String * 5

Public Sub Export(SourceFields As String, SourceDB As String, WHEREFilter As String)

Dim ExportDB As Database
Dim sConnect As String
Dim sDBName As String
Dim ExportFile As String
Dim ExportTotalLength As Integer
Dim ExportFileLength As Integer
Dim ExportPathLength As Integer
Dim ExportPath As String
Dim ExportNamewoExt As String

On Error GoTo errorhandler

frmSQLQuery.cdlProject.Flags = cdlOFNOverwritePrompt
frmSQLQuery.cdlProject.DialogTitle = "Export Data"
frmSQLQuery.cdlProject.filename = ""
    
frmSQLQuery.cdlProject.ShowSave
    
ExportFile = frmSQLQuery.cdlProject.filename

ExportTotalLength = Len(ExportFile)

ExportFileLength = Len(frmSQLQuery.cdlProject.FileTitle)

ExportPathLength = ExportTotalLength - ExportFileLength

ExportPath = Left(ExportFile, ExportPathLength)
    
ExportNamewoExt = Mid(ExportFile, ExportPathLength + 1, ExportTotalLength - ExportPathLength - 4)  '4 for . + file extension
  
Select Case DataType

Case "acces"
    sConnect = "[;database=" & ExportFile & "]."
    sDBName = ExportNamewoExt
    Set ExportDB = CreateDatabase(ExportFile, dbLangGeneral)
    ExportDB.Close

Case "fox20"
    sConnect = "[FoxPro 2.0;database=" & ExportPath & "]."
    sDBName = frmSQLQuery.cdlProject.FileTitle
'need to open database read/write otherwise assumed to be read only and cant export
    Set ExportDB = OpenDatabase(ExportPath, False, False, "FOXPRO 2.0; database=" & ExportPath)
    ExportDB.Close

Case "fox25"
    sConnect = "[FoxPro 2.5;database=" & ExportPath & "]."
    sDBName = frmSQLQuery.cdlProject.FileTitle
'need to open database read/write otherwise assumed to be read only and cant export
    Set ExportDB = OpenDatabase(ExportPath, False, False, "FOXPRO 2.5; database=" & ExportPath)
    ExportDB.Close

Case "fox26"
    sConnect = "[FoxPro 2.6;database=" & ExportPath & "]."
    sDBName = frmSQLQuery.cdlProject.FileTitle
'need to open database read/write otherwise assumed to be read only and cant export
    Set ExportDB = OpenDatabase(ExportPath, False, False, "FOXPRO 2.6; database=" & ExportPath)
    ExportDB.Close

Case "exc30"
    sConnect = "[Excel 3.0;database=" & ExportFile & "]."
    sDBName = ExportNamewoExt

Case "exc40"
    sConnect = "[Excel 4.0;database=" & ExportFile & "]."
    sDBName = ExportNamewoExt

Case "exc50"
    sConnect = "[Excel 5.0;database=" & ExportFile & "]."
    sDBName = ExportNamewoExt

Case "exc97"
    sConnect = "[Excel 8.0;database=" & ExportFile & "]."
    sDBName = ExportNamewoExt

Case "lot01"
    sConnect = "[LOTUS WK1;database=" & ExportFile & "]."
    sDBName = ExportNamewoExt

Case "lot03"
    sConnect = "[LOTUS WK3;database=" & ExportFile & "]."
    sDBName = ExportNamewoExt
    
Case "comma"
    sConnect = "[Text;database=" & ExportPath & "]."
    sDBName = frmSQLQuery.cdlProject.FileTitle

End Select

If Len(Trim(WHEREFilter)) > 0 Then
    MyDatabase.Execute "SELECT " + SourceFields + " INTO " & sConnect & sDBName & " from " & SourceDB & " WHERE " & WHEREFilter
Else
    MyDatabase.Execute "SELECT " + SourceFields + " INTO " & sConnect & sDBName & " from " & SourceDB
End If

errorhandler:
    Select Case Err.Number
    
    Case 32755  'cancel pressed
        Exit Sub
    
    Case 3010, 3204 'file exists - user ok'ed overwrite, 3204 for Access CreateDatabase method
        Kill ExportFile  'delete file to be overwritten - otherwise SQL wont work
        DoEvents
        Resume
    Case Is > 0
        MsgBox "Application Error." + Chr(10) + Chr(13) + Err.Description, vbExclamation + vbOKOnly, "Application Error"
        Exit Sub
    
    End Select

End Sub

