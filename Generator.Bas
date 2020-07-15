
Option Explicit

Public Const BaseDirectory As String = _
    "C:\Users\Katherine\Documents\_Excel Projects\Chemicals\Management_of_Change\"
    
    
Public Sub user_inputs(ByRef MOCname As String, ByRef MOCNumber As String)

    MOCname = InputBox("Please enter the MOC name: ", "MOC name: ")
    MOCNumber = InputBox("Please enter the MOC number: ", "MOC number: ")
    
End Sub

Private Function MakeMOCDir(ByVal MOCname As String, ByVal MOCNumber As String, ByRef output_dir As String) As Boolean
    
    MakeMOCDir = False
    
    output_dir = BaseDirectory & Application.PathSeparator & MOCNumber & "-" & MOCname & Application.PathSeparator
    On Error GoTo ALREADY_EXISTS
    MkDir output_dir
    
ALREADY_EXISTS:

    If Err.Number = 75 Then
        Dim FldrXt As VbMsgBoxResult
        FldrXt = MsgBox("Folder Already Exists. Click Yes to Replace File. Click No to Cancel.", vbYesNo, "Warning")
        If FldrXt = vbNo Then
            MakeMOCDir = False
            Exit Function
        End If
        
    ElseIf Err.Number > 0 Then
        MsgBox "File name error. Make sure user has write access to folder and no special characters are used.", vbOKOnly, "Warning"
        MakeMOCDir = False
        Exit Function
    End If
    
    MakeMOCDir = True
        
End Function

Private Function MakeDoc(ByVal MOCname As String, ByVal MOCNumber As String, ByVal output_dir As String) As Boolean

    MakeDoc = False
    Application.ScreenUpdating = False
    
    On Error GoTo FAILURE
    
    Dim doc As Document
    Set doc = Documents.Add
    
    Dim p As Paragraph
    Set p = doc.Paragraphs.Add
    p.Range.InsertBefore "Updates to " & MOCname & vbCrLf & MOCNumber
    
    Set p = doc.Paragraphs.Add
    p.Range.InsertBefore " All," & vbCrLf & "See the below link for the MOC folder for this change. Please review all documents in the folder as part of this change."

    Set p = doc.Paragraphs.Add
    p.Range.InsertBefore BaseDirectory & MOCNumber & "-" & MOCname
    doc.Hyperlinks.Add Anchor:=p.Range, Address:=BaseDirectory & MOCNumber & "-" & MOCname
        
    doc.SaveAs2 output_dir & "Shortcut Document"
    doc.Close
    
    MakeDoc = True
    
FAILURE:
    Application.ScreenUpdating = True
    
End Function

Sub Automate_MOC_Script()
    
' Automate_MOC_Script Macro

    Dim MOCname As String
    Dim MOCNumber As String
    Dim output_dir As String
    user_inputs MOCname, MOCNumber
    
    On Error GoTo ENDTHIS
    
    If Not MakeMOCDir(MOCname, MOCNumber, output_dir) Then
        MsgBox "Error making the directory."
        Exit Sub
    End If
    
    If Not MakeDoc(MOCname, MOCNumber, output_dir) Then
        MsgBox "Error making the word document."
        Exit Sub
    End If
    
    MsgBox "Complete!"
    
ENDTHIS:
   Application.ScreenUpdating = True
   
End Sub
