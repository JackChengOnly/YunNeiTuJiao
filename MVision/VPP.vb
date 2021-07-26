Imports Cognex.VisionPro
Imports Cognex.VisionPro.QuickBuild

Public Class VPP
    Inherits System.Windows.Forms.Form
    Private mJM As CogJobManager = Nothing
    Public Sub New(ByVal jm As CogJobManager)
        Try
            'This call is required by the Windows Form Designer.
            InitializeComponent()

            mJM = jm
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Form2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If MessageBox.Show("确定要退出 工具设置 界面 吗？", "提示信息", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = Windows.Forms.DialogResult.OK Then
                CogJobManagerEdit1.Dispose()
                GC.Collect()
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            CogJobManagerEdit1.Subject = mJM
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

End Class