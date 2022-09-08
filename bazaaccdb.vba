===============
Form_add_book
===============
Option Compare Database

Private Sub confirm_Click()

    If Me.save.Enabled = False Then
        If IsNull(Me.book_id) = False And IsNull(Me.title) = False And IsNull(Me.author) = False And IsNull(Me.admission_date) = False Then
            Me.save.Enabled = True
        End If
    End If

End Sub

Private Sub discard_Click()
    SendKeys "{esc}", True
    SendKeys "{NUMLOCK}", True
    DoCmd.Close
    DoCmd.OpenForm FormName:="main"
End Sub

Private Sub Form_Timer()
    
    Me.Time.Requery
    Me.Date.Requery

End Sub
===============
Form_main
===============
Option Compare Database

Private Sub Form_Load()

End Sub

Private Sub Form_Timer()
    
    Me.Time.Requery
    Me.Date.Requery

End Sub
===============
GitSave
===============
Option Compare Database
Option Explicit

   Sub GitSave()

   Dim fs As Object
   Dim f As Object
   Dim strMod As String
   Dim mdl As Object
   Dim i As Integer

   Set fs = CreateObject("Scripting.FileSystemObject")

   Set f = fs.CreateTextFile(CurrentProject.Path & "\" _
       & Replace(CurrentProject.Name, ".", "") & ".vba")

   For Each mdl In VBE.ActiveVBProject.VBComponents
       i = VBE.ActiveVBProject.VBComponents(mdl.Name).codemodule.CountOfLines
       If i > 0 Then
          strMod = VBE.ActiveVBProject.VBComponents(mdl.Name).codemodule.Lines(1, i)
       End If
       f.writeline String(15, "=") & vbCrLf & mdl.Name _
           & vbCrLf & String(15, "=") & vbCrLf & strMod
   Next

   f.Close
   Set fs = Nothing
   End Sub
