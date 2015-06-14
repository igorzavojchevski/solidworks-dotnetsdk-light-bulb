Public Class Form2

    Dim app As SldWorks.SldWorks
    Dim p1 As Boolean = False
    Dim p2 As Boolean = False
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ProgressBar1.Hide()
        app = CreateObject("SldWorks.Application")
        If app IsNot Nothing Then
            app.Visible = True
        Else
            Dim Msg = "SolidWorks не може да се стартува!"
            MsgBox(Msg)
            Me.Close()
        End If
    End Sub
    Private Sub Part1_Click(sender As Object, e As EventArgs) Handles Part1.Click
        ProgressBar1.Show()
        ProgressBar1.Value = 0
        GorenDel(app, ProgressBar1)
        If ProgressBar1.Value = ProgressBar1.Maximum Then
            p1 = True
            Dim Msg = "Горниот дел на светилката е креиран и зачуван!"
            MsgBox(Msg)
        End If
    End Sub
    Private Sub Part2_Click(sender As Object, e As EventArgs) Handles Part2.Click
        ProgressBar1.Show()
        ProgressBar1.Value = 0
        DolenDel(app, ProgressBar1)
        If ProgressBar1.Value = ProgressBar1.Maximum Then
            p2 = True
            Dim Msg = "Долниот дел на светилката е креиран и зачуван!"
            MsgBox(Msg)
        End If
    End Sub
    Private Sub Assem_Click(sender As Object, e As EventArgs) Handles Assem.Click
        If p1 And p2 Then
            ProgressBar1.Show()
            ProgressBar1.Value = 0
            Assembly(app, ProgressBar1)
            If ProgressBar1.Value = ProgressBar1.Maximum Then
                Dim Msg = "Светилката е креирана и зачувана!"
                MsgBox(Msg)
            End If
        ElseIf p1 And Not p2 Then
            ProgressBar1.Hide()
            MsgBox("Најпрвин креирајте го долниот дел на светилката!")
        ElseIf p2 And Not p1 Then
            ProgressBar1.Hide()
            MsgBox("Најпрвин креирајте го горниот дел на светилката!")
        Else
            ProgressBar1.Hide()
            MsgBox("Најпрвин креирајте го горниот и долниот дел на светилката!")
        End If
    End Sub
    Private Sub Render1_Click(sender As Object, e As EventArgs) Handles Render1.Click
        If p1 And p2 Then
            RenderPrev(app)
        ElseIf p1 And Not p2 Then
            ProgressBar1.Hide()
            MsgBox("Најпрвин креирајте го долниот дел на светилката!")
        ElseIf p2 And Not p1 Then
            ProgressBar1.Hide()
            MsgBox("Најпрвин креирајте го горниот дел на светилката!")
        Else
            ProgressBar1.Hide()
            MsgBox("Најпрвин креирајте го горниот и долниот дел на светилката!")
        End If
    End Sub

    Private Sub Render2_Click(sender As Object, e As EventArgs) Handles Render2.Click
        If p1 And p2 Then
            FinalRender(app)
        ElseIf p1 And Not p2 Then
            ProgressBar1.Hide()
            MsgBox("Најпрвин креирајте го долниот дел на светилката!")
        ElseIf p2 And Not p1 Then
            ProgressBar1.Hide()
            MsgBox("Најпрвин креирајте го горниот дел на светилката!")
        Else
            ProgressBar1.Hide()
            MsgBox("Најпрвин креирајте го горниот и долниот дел на светилката!")
        End If
    End Sub
    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        app.ExitApp()
        Me.Close()
    End Sub
End Class