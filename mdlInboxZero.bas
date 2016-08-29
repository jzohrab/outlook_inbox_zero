Attribute VB_Name = "mdlInboxZero"
Public Type FiveSentenceEmailContent_t
  Greeting As String
  Message As String
  AddWaitingCategory As Boolean
End Type



Public Sub set_waiting_category()
  Dim i As Outlook.MailItem
  Set i = Application.ActiveInspector.CurrentItem
  If i Is Nothing Then
    Exit Sub
  End If
  
  c = i.Categories
  If InStr(1, c, "waiting", vbTextCompare) <> 0 Then
    Call remove_category(i, "waiting")
    i.Save
    Exit Sub
  End If
  
  c = IIf(c = "", "waiting", c & ", waiting")
  i.Categories = c
  i.Save
End Sub


Sub remove_category(itm As Outlook.MailItem, catName As String)
    arr = Split(itm.Categories, ",")
    If UBound(arr) >= 0 Then
        ' item has categories
        For i = 0 To UBound(arr)
            If Trim(arr(i)) = catName Then
                ' category already exists on item
                ' remove it
                arr(i) = ""
                'rebuild category list from array
                itm.Categories = Join(arr, ",")
                Exit Sub
            End If
        Next
    End If
End Sub


Sub add_five_sentences_outline()
  Dim i As Outlook.MailItem
  Set i = Application.ActiveInspector.CurrentItem
  If i Is Nothing Then
    Exit Sub
  End If

  ' Dim review_points As String
  ' review_points = "Review:" & vbCrLf & _
  ' "1. ensure clarity of main point" & vbCrLf & _
  ' "2. omit needless words/ideas" & vbCrLf & _
  ' "3. can it be simplified?"

  Dim r As FiveSentenceEmailContent_t
  r = get_five_sentence_email_content()
  Dim msg As String
  msg = r.Greeting & vbCrLf & vbCrLf & r.Message
  
  If TypeName(ActiveWindow) = "Inspector" Then
    If ActiveInspector.IsWordMail And ActiveInspector.EditorType = olEditorWord Then
      ActiveInspector.WordEditor.Application.Selection.TypeText msg
    End If
  End If

  If r.AddWaitingCategory Then
    Call set_waiting_category
  End If
End Sub



Private Function get_five_sentence_email_content() As FiveSentenceEmailContent_t
  Dim f As frm_five_sentences
  Set f = New frm_five_sentences
  f.Show vbModal
  Dim c As New Collection
  c.Add f.txt_who_i_am
  c.Add f.txt_what_i_want
  c.Add f.txt_why_i_am_asking
  c.Add f.txt_why_you_should_do_it
  c.Add f.txt_what_is_the_next_step
  
  Dim msg As String
  Dim t As TextBox
  For Each t In c
    If t.Text <> "" Then
      msg = msg & IIf(msg <> "", vbCrLf & vbCrLf, "") & t.Text
    End If
  Next t
  
  Dim r As FiveSentenceEmailContent_t
  r.Greeting = f.txt_salutation
  r.Message = msg
  r.AddWaitingCategory = f.chk_waiting
  
  get_five_sentence_email_content = r
End Function


Private Sub test_show_five_s_form()
  Dim r As FiveSentenceEmailContent_t
  r = get_five_sentence_email_content
  MsgBox r.Greeting
  MsgBox r.Message
  MsgBox r.AddWaitingCategory
End Sub
