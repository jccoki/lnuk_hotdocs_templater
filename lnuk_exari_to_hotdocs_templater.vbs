' -------------------------------------------------------------------------------------------------------------
' @author John Christian Campomanes
' -------------------------------------------------------------------------------------------------------------
' https://msdn.microsoft.com/en-us/library/system.collections.arraylist(v=vs.110).aspx
' https://msdn.microsoft.com/en-us/library/office/gg251825.aspx
Option Explicit

Const wdContentControlRichText = 0
Const wdStory = 6
Const wdReplaceAll = 2

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type
 
Private Declare PtrSafe Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
 
Public Function GetGUID() As String
  ' see https://stackoverflow.com/questions/7031347/how-can-i-generate-guids-in-excel
  '(c) 2000 Gus Molina
  Dim udtGUID As GUID
  
  If (CoCreateGuid(udtGUID) = 0) Then
  GetGUID = _
    String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
    String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
    String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
    IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
    IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
    IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
    IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
    IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
    IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
    IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
    IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
  End If
End Function

Sub ProcessVariable(objDoc, r_start, r_end)
  Dim objRange
  Dim objContentControl
  Dim uuid_value, content_control_value
  
  Set objRange = objDoc.Range(r_start, r_end)
  
  content_control_value = Replace(objRange.Text, "[", "")
  content_control_value = Replace(content_control_value, "]", "")
  content_control_value = Replace(content_control_value, ".", "")
  
  ' generate UUID and format into ISO/IEC 9834-8:2008 standard
  uuid_value = GetGUID()
  uuid_value = LCase(uuid_value)
  uuid_value = Left(uuid_value, 8) & "-" & Mid(uuid_value, 9, 4) & "-" & _
    Mid(uuid_value, 13, 4) & "-" & Mid(uuid_value, 17, 4) & "-" & _
    Right(uuid_value, 12)

  ' target document should first be "developer" ready
  Set objContentControl = objDoc.ContentControls.Add(wdContentControlRichText, objDoc.Range(r_start, r_end))
  objContentControl.Tag = "HD:1.185.0.0:" & uuid_value
  
  ' @note specify all params as discussed in https://answers.microsoft.com/en-us/msoffice/forum/msoffice_word-msoffice_custom-mso_2010/setplaceholdertext-method-fails-if-using-late/2637c5fb-cafc-4913-8780-752069c8522b
  ' else this code will fail with type mismatch error
  objContentControl.SetPlaceholderText Nothing, Nothing, Text:=content_control_value
  ' replace text inside square brackets with normalized value
  ' variable naming convention should be the same as generated from the logic file parser
  objRange.Text = content_control_value
End Sub

Sub ProcessTemplate(template_input_path, psl_input_path)
    Dim objFSO, objRegEx, objDoc, objWord, objSelection, objCustomProperty, objComment
    Dim input_file_path, curr_dir, output_file_path, psl_output_file_path
    Dim variable_regex_pattern
    Dim custom_property_exists As Boolean

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objRegEx = CreateObject("VBScript.RegExp")
    ' we need to continue through errors since if Word isn't
    ' open the GetObject line will give an error
    Set objWord = GetObject("", "Word.Application")
    ' we've tried to get Word but if it's nothing then it isn't open
    If objWord Is Nothing Then
      Set objWord = CreateObject("Word.Application")
    End If
    ' it's good practice to reset error warnings
    On Error GoTo 0

    ' open your document and ensure its visible and activate after openning
    objWord.Visible = True
    objWord.Activate

    ' search throught the string
    objRegEx.Global = True
    
    ' regex patter to match variables inside square brackets
    variable_regex_pattern = "\[([a-zA-Z0-9_.]{1,})\]"
    
    ' build input file path from current directory
    curr_dir = ActiveWorkbook.Path

    output_file_path = objFSO.BuildPath(curr_dir, "output")
    output_file_path = objFSO.BuildPath(output_file_path, objFSO.GetFileName(template_input_path))

    psl_output_file_path = objFSO.BuildPath(curr_dir, "output")
    psl_output_file_path = objFSO.BuildPath(psl_output_file_path, objFSO.GetFileName(psl_input_path))

    Set objDoc = objWord.Documents.Open(template_input_path)
    Set objSelection = objWord.Selection

    ' clean the template from comments
    For Each objComment In objDoc.Comments
      objComment.DeleteRecursively
    Next objComment

    ' add hotdocs custom properties
    custom_property_exists = False
    For Each objCustomProperty In objDoc.CustomDocumentProperties
      If objCustomProperty.Name = "HotDocs version" Then
        custom_property_exists = True
        Exit For
      End If
    Next objCustomProperty
    
    If Not custom_property_exists Then
        objDoc.CustomDocumentProperties.Add Name:="HotDocs version", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=12
    End If
    
    ' search and process simple variable markup
    Do
      With objSelection.Find
       .Text = variable_regex_pattern
       .MatchAllWordForms = False
       .MatchSoundsLike = False
       .MatchFuzzy = False
       .MatchWildcards = True
       .Forward = True
       .Execute
      End With
      If objSelection.Find.Found Then
        objSelection.Range.Select
        ProcessVariable objDoc, objSelection.Range.Start, objSelection.Range.End
      End If
    Loop While objSelection.Find.Found
    objSelection.HomeKey (wdStory)

    ' save changes and close MS Word output
    objDoc.SaveAs (output_file_path)
    objDoc.Close
    Set objDoc = Nothing

    ' convert PSL flat document into Hotdocs compatible template
    Set objDoc = objWord.Documents.Open(psl_input_path)

    ' clean the template from comments
    For Each objComment In objDoc.Comments
      objComment.DeleteRecursively
    Next objComment

    ' add hotdocs custom properties
    custom_property_exists = False
    For Each objCustomProperty In objDoc.CustomDocumentProperties
      If objCustomProperty.Name = "HotDocs version" Then
        custom_property_exists = True
        Exit For
      End If
    Next objCustomProperty

    If Not custom_property_exists Then
        objDoc.CustomDocumentProperties.Add Name:="HotDocs version", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=12
    End If

    ' save changes and close MS Word output
    objDoc.SaveAs (psl_output_file_path)
    objDoc.Close
    Set objDoc = Nothing
End Sub