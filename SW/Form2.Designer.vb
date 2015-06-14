<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Part1 = New System.Windows.Forms.Button()
        Me.Part2 = New System.Windows.Forms.Button()
        Me.CloseButton = New System.Windows.Forms.Button()
        Me.Assem = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Render1 = New System.Windows.Forms.Button()
        Me.Render2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Part1
        '
        Me.Part1.Location = New System.Drawing.Point(60, 40)
        Me.Part1.Name = "Part1"
        Me.Part1.Size = New System.Drawing.Size(90, 50)
        Me.Part1.TabIndex = 0
        Me.Part1.Text = "Горен дел"
        Me.Part1.UseVisualStyleBackColor = True
        '
        'Part2
        '
        Me.Part2.Location = New System.Drawing.Point(180, 40)
        Me.Part2.Name = "Part2"
        Me.Part2.Size = New System.Drawing.Size(90, 50)
        Me.Part2.TabIndex = 1
        Me.Part2.Text = "Долен дел"
        Me.Part2.UseVisualStyleBackColor = True
        '
        'CloseButton
        '
        Me.CloseButton.Location = New System.Drawing.Point(243, 276)
        Me.CloseButton.Name = "CloseButton"
        Me.CloseButton.Size = New System.Drawing.Size(89, 23)
        Me.CloseButton.TabIndex = 2
        Me.CloseButton.Text = "Крај"
        Me.CloseButton.UseVisualStyleBackColor = True
        '
        'Assem
        '
        Me.Assem.Location = New System.Drawing.Point(60, 138)
        Me.Assem.Name = "Assem"
        Me.Assem.Size = New System.Drawing.Size(210, 45)
        Me.Assem.TabIndex = 3
        Me.Assem.Text = "Креирај Светилка"
        Me.Assem.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(60, 237)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(210, 14)
        Me.ProgressBar1.TabIndex = 4
        '
        'Render1
        '
        Me.Render1.Location = New System.Drawing.Point(60, 199)
        Me.Render1.Name = "Render1"
        Me.Render1.Size = New System.Drawing.Size(104, 23)
        Me.Render1.TabIndex = 5
        Me.Render1.Text = "Render Preview"
        Me.Render1.UseVisualStyleBackColor = True
        '
        'Render2
        '
        Me.Render2.Location = New System.Drawing.Point(170, 199)
        Me.Render2.Name = "Render2"
        Me.Render2.Size = New System.Drawing.Size(99, 23)
        Me.Render2.TabIndex = 6
        Me.Render2.Text = "Final Render"
        Me.Render2.UseMnemonic = False
        Me.Render2.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(344, 311)
        Me.Controls.Add(Me.Render2)
        Me.Controls.Add(Me.Render1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Assem)
        Me.Controls.Add(Me.CloseButton)
        Me.Controls.Add(Me.Part2)
        Me.Controls.Add(Me.Part1)
        Me.Name = "Form2"
        Me.Text = "Светилка"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Part1 As System.Windows.Forms.Button
    Friend WithEvents Part2 As System.Windows.Forms.Button
    Friend WithEvents CloseButton As System.Windows.Forms.Button
    Friend WithEvents Assem As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Render1 As System.Windows.Forms.Button
    Friend WithEvents Render2 As System.Windows.Forms.Button
End Class
