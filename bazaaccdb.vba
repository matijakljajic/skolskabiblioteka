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

Private Sub Form_Load()

    DoCmd.Maximize
    Application.RunCommand acCmdAppMaximize

End Sub

Private Sub Form_Timer()
    
    Me.Time.Requery
    
End Sub
===============
Form_main
===============
Option Compare Database

Private Sub Form_Load()

    DoCmd.Maximize
    Application.RunCommand acCmdAppMaximize

End Sub

Private Sub Form_Timer()
    
    Me.Time.Requery

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
===============
Form_add_member
===============
Option Compare Database

Private Sub confirm_Click()

    If Me.save.Enabled = False Then
        If IsNull(Me.member_id) = False And IsNull(Me.first_name) = False And IsNull(Me.last_name) = False And IsNull(Me.school) = False And IsNull(Me.admission_year) = False And IsNull(Me.class) = False And IsNull(Me.phone) = False Then
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

Private Sub Form_Load()

    DoCmd.Maximize
    Application.RunCommand acCmdAppMaximize

End Sub

Private Sub Form_Timer()
    
    Me.Time.Requery

End Sub
===============
Form_all_members
===============
Option Compare Database

Private Sub Button_search_Click()
     
    Dim SQL As String
    
    SQL = "SELECT [member].[member_id], [member].[first_name], [member].[last_name], [member].[admission_year], [member].[class], [member].[school], [member].[phone] " _
    & "FROM member " _
    & "WHERE [member].[first_name] LIKE '*" & Me.fsearch & "*' " _
    & "AND [member].[last_name] LIKE '*" & Me.lsearch & "*' "
    
    Form.RecordSource = SQL
    Form.Requery

End Sub

Private Sub Form_Load()

    DoCmd.Maximize
    Application.RunCommand acCmdAppMaximize

End Sub

Private Sub Form_Timer()

    Me.Time.Requery

End Sub
===============
Form_add_borrowing
===============
Private Sub confirm_Click()

    If Me.save.Enabled = False Then
        If IsNull(Me.book_id) = False And IsNull(Me.member_id) = False Then
            Me.save.Enabled = True
        End If
    End If
    
End Sub

Private Sub discard_Click()
    
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "book_t"
    DoCmd.SetWarnings True
        
    SendKeys "{esc}", True
    SendKeys "{NUMLOCK}", True
    DoCmd.Close
    DoCmd.OpenForm FormName:="main"
    
End Sub
===============
Form_all_books
===============
Option Compare Database

Private Sub Button_search_Click()

    Dim SQL As String
    
    SQL = "SELECT [book].[book_id], [book].[title], [book].[naziv], [book].[author], [book].[admission_date], [book].[borrowed], [book].[lost] " _
    & "FROM book " _
    & "WHERE [book].[title] LIKE '*" & Me.tsearch & "*' " _
    & "AND [book].[author] LIKE '*" & Me.asearch & "*' "

    Form.RecordSource = SQL
    Form.Requery

End Sub

Private Sub Form_Load()

    DoCmd.Maximize
    Application.RunCommand acCmdAppMaximize

End Sub

Private Sub Form_Timer()

    Me.Time.Requery

End Sub
