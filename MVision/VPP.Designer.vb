<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class VPP
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.CogJobManagerEdit1 = New Cognex.VisionPro.QuickBuild.CogJobManagerEdit()
        Me.SuspendLayout()
        '
        'CogJobManagerEdit1
        '
        Me.CogJobManagerEdit1.Location = New System.Drawing.Point(0, 0)
        Me.CogJobManagerEdit1.Name = "CogJobManagerEdit1"
        Me.CogJobManagerEdit1.ShowLocalizationTab = False
        Me.CogJobManagerEdit1.Size = New System.Drawing.Size(1183, 879)
        Me.CogJobManagerEdit1.Subject = Nothing
        Me.CogJobManagerEdit1.TabIndex = 0
        '
        'VPP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1184, 882)
        Me.Controls.Add(Me.CogJobManagerEdit1)
        Me.Name = "VPP"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VPP"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CogJobManagerEdit1 As Cognex.VisionPro.QuickBuild.CogJobManagerEdit
End Class
