Sub template_text_in_existing_email_with_variables()

Dim objDoc As Word.Document, objSel As Word.Selection
Dim item As Outlook.MailItem
Dim oinspector As Inspector
Set oinspector = Application.ActiveInspector


strsubject = "template email subject line " 

If oinspector Is Nothing Then
        Set item = Application.ActiveExplorer.Selection.item(1)
    Else
       Set item = oinspector.CurrentItem
    End If

item.Subject = strsubject

    On Error Resume Next
    
        gimmer = InputBox("Enter variable one", "Using the VBA Input Box", "Variable 1")
        gimmea = InputBox("Enter variable two: ", "Using the VBA Input Box", "The 'Variable 2")
        gimmet = InputBox("Enter variable three ", "Using the VBA Input Box", "The variable 3")
                                       

    strbody = "We have " & gimmer & _
              " and " & gimmea & _
              " with " & gimmet & _
              "and additional information stating ******" & vbNewLine & vbNewLine & _
              "Please respond." & vbNewLine & vbNewLine 
              
     
     
    '~~> Get a Word.Selection from the open Outlook item
    Set objDoc = Application.ActiveInspector.WordEditor
    Set objSel = objDoc.Windows(1).Selection
    
   
    '~~> Type Relevant Text
    
    objSel.TypeText strbody & vbNewLine & vbNewLine
    
    Set objDoc = Nothing
    Set objSel = Nothing
    Set item = Nothing
    Set oinspector = Nothing

End Sub
