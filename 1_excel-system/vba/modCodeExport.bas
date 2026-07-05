Attribute VB_Name = "modCodeExport"
Option Explicit

' ============================================================
' modCodeExport
' One-click exporter for VBA modules + Power Query M code
' ============================================================

Public Sub ExportAllCode()

    On Error GoTo Fail
    
    Dim exportDir As String
    Dim vbaDir As String
    Dim mDir As String
    Dim manifestPath As String
    Dim ff As Integer
    Dim vbaCount As Long
    Dim mCount As Long
    
    ' -------- Config --------
    ' Set this to your repo folder. Leave as empty string "" to use the
    ' workbook's own folder instead.
    Dim repoPath As String
    repoPath = "C:\Users\gwolter\Desktop\GitHub\Canneberge\1_excel-system"   ' <-- PUT YOUR REPO PATH HERE, e.g. "C:\Users\Gail\Documents\GitHub\ProjectCanneberge"
    ' ------------------------
    
    If Len(repoPath) = 0 Then
    MsgBox "repoPath is not set — edit modCodeExport.", vbExclamation
    Exit Sub
    End If

    exportDir = repoPath                       ' root of repo
    vbaDir = repoPath & "\vba"                 ' matches existing folder
    mDir = repoPath & "\power-query"           ' matches existing folder

    EnsureFolder vbaDir
    EnsureFolder mDir
    
    ' Overwrite mode — clean out prior files
    CleanFolder vbaDir
    CleanFolder mDir
    
    vbaCount = ExportVBAComponents(vbaDir)
    mCount = ExportPowerQueries(mDir)
    
    ' Write manifest
    manifestPath = exportDir & "\_manifest.txt"
    ff = FreeFile
    Open manifestPath For Output As #ff
    Print #ff, "Project Canneberge - Code Export Manifest"
    Print #ff, String(60, "=")
    Print #ff, "Exported     : " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #ff, "Workbook     : " & ThisWorkbook.FullName
    Print #ff, "VBA files    : " & vbaCount
    Print #ff, "M queries    : " & mCount
    Print #ff, "Export folder: " & exportDir
    Close #ff
    
    MsgBox "Code export complete." & vbCrLf & vbCrLf & _
           "VBA files : " & vbaCount & vbCrLf & _
           "M queries : " & mCount & vbCrLf & vbCrLf & _
           "Folder: " & exportDir, vbInformation, "Export Complete"
    Exit Sub

Fail:
    MsgBox "Export failed: " & Err.Description & vbCrLf & vbCrLf & _
           "If this is error 1004 about 'programmatic access', enable:" & vbCrLf & _
           "  File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
           "  Macro Settings > 'Trust access to the VBA project object model'", _
           vbCritical, "Export Failed"
End Sub


' ============================================================
' Export ALL VBA components — modules, classes, forms, AND document modules
' ============================================================
Private Function ExportVBAComponents(ByVal targetDir As String) As Long

    Dim comp As Object      ' VBIDE.VBComponent
    Dim ext As String
    Dim outPath As String
    Dim count As Long
    
    count = 0
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        ext = ExtensionForComponent(comp.Type)
        outPath = targetDir & "\" & comp.Name & ext
        comp.Export outPath
        count = count + 1
    Next comp
    
    ExportVBAComponents = count
End Function


Private Function ExtensionForComponent(ByVal compType As Integer) As String
    ' VBIDE component type constants:
    ' 1 = vbext_ct_StdModule      -> .bas
    ' 2 = vbext_ct_ClassModule    -> .cls
    ' 3 = vbext_ct_MSForm         -> .frm  (also produces .frx binary)
    ' 100 = vbext_ct_Document     -> .cls  (ThisWorkbook, Sheet1, etc.)
    Select Case compType
        Case 1:   ExtensionForComponent = ".bas"
        Case 2:   ExtensionForComponent = ".cls"
        Case 3:   ExtensionForComponent = ".frm"
        Case 100: ExtensionForComponent = ".cls"
        Case Else: ExtensionForComponent = ".txt"
    End Select
End Function


' ============================================================
' Export all Power Query M code
' ============================================================
Private Function ExportPowerQueries(ByVal targetDir As String) As Long

    Dim q As WorkbookQuery
    Dim outPath As String
    Dim ff As Integer
    Dim count As Long
    Dim safeName As String
    
    count = 0
    
    For Each q In ThisWorkbook.Queries
        safeName = SanitizeFileName(q.Name)
        outPath = targetDir & "\" & safeName & ".m"
        
        ff = FreeFile
        Open outPath For Output As #ff
        Print #ff, q.formula
        Close #ff
        
        count = count + 1
    Next q
    
    ExportPowerQueries = count
End Function


' ============================================================
' Helpers
' ============================================================
Private Sub EnsureFolder(ByVal path As String)
    If Len(Dir(path, vbDirectory)) = 0 Then MkDir path
End Sub

Private Sub CleanFolder(ByVal path As String)
    On Error Resume Next
    Kill path & "\*.bas"
    Kill path & "\*.cls"
    Kill path & "\*.frm"
    Kill path & "\*.frx"
    Kill path & "\*.m"
    On Error GoTo 0
End Sub

Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In bad
        s = Replace(s, CStr(ch), "_")
    Next ch
    SanitizeFileName = s
End Function

