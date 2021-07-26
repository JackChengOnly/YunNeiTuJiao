Option Strict Off
Option Explicit On

Imports Cognex.VisionPro
Imports Cognex.VisionPro.QuickBuild
Imports Cognex.VisionPro.ImageFile
Imports Cognex.VisionPro.Exceptions
Imports Cognex.VisionPro.ToolGroup
Imports System.Object
Imports System.Runtime
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Net.Sockets
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Threading
Imports Qplc
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports OpcRcw.Comn
Imports OpcRcw.Da



Public Class Form1
    Implements IOPCDataCallback
    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub
  
    Private myJobManager As CogJobManager
    Private myJob0, myJob1, myJob2, myJob3, myJob4 As CogJob
    Private myIndependentJob0, myIndependentJob1, myIndependentJob2, myIndependentJob3, myIndependentJob4 As CogJobIndependent
    Private mGroup0, mGroup1, mGroup2, mGroup3 As CogToolGroup
    Private mTool0, mTool1, mTool2, mTool3, mTool4, mTool5, mTool6 As CogToolBlock
    Delegate Sub myJobManagerDelegate(ByVal sender As Object, ByVal e As CogJobManagerActionEventArgs)
    Public str_Info(100) As String
    Public SN As String = ""
    Public IP() As String
    Public Passworld() As String
    Public PassworldArray(0 To 100, 0 To 1) As String
    Public PN() As String
    Public Current_PN As String
    Public PNArray(0 To 100, 0 To 1) As String
    Public CCD1_Data() As String
    Public CCD1_DataArray(0 To 100, 0 To 1) As String
    Public CCD2_Data() As String
    Public CCD2_DataArray(0 To 100, 0 To 1) As String
    Public CCD3_Data() As String
    Public CCD3_DataArray(0 To 100, 0 To 1) As String
    Public CCD4_Data() As String
    Public CCD4_DataArray(0 To 100, 0 To 1) As String
    Private Count_NG As Integer = 0
    Private Count_OK As Integer = 0
    Dim RunState As Boolean = True
    Dim ID As String
    Dim quantao As String
    Private mjob1, mjob2, mjob3, mjob4 As Thread
#Region "加载"
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If UBound(Diagnostics.Process.GetProcessesByName( _
     Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
                MessageBox.Show("当前版本的应用程序不允许重复运行 ！  ", "禁止运行", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Application.Exit()
            End If
            Control.CheckForIllegalCrossThreadCalls = False
            Form_show()
            checklicense()
            Read_Data()
            VPP_add(Current_PN)
            Thread.Sleep(500)
            'Start_Sample()
            Thread.Sleep(500)
            'Timer2.Enabled = True
            'chkGroupActive.Checked = True
            'Write_value(True, 7)
        Catch ex As Exception
            MessageBox.Show("加载失败")
            Application.Exit()
        End Try
    End Sub
#End Region
#Region "关闭"
    Private Sub Frm_Main_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            If MessageBox.Show("确定要退出 工具设置 界面 吗？", "提示信息", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = Windows.Forms.DialogResult.OK Then
                RemoveHandler myJobManager.Stopped, AddressOf myJobManager_Stopped
                RemoveHandler myJobManager.UserResultAvailable, AddressOf myJobManager_UserResultAvailable
                myJobManager.Shutdown()
                CogRecordDisplay1.Dispose()
                CogRecordDisplay2.Dispose()
                CogRecordDisplay3.Dispose()
                CogRecordDisplay4.Dispose()
                'CogRecordDisplay5.Dispose()
                'CogRecordDisplay6.Dispose()
                'CogRecordDisplay7.Dispose()
                'CogRecordDisplay8.Dispose()
                'Timer2.Enabled = False
                'Write_value(False, 8)
                'Write_value(False, 9)
                'Write_value(False, 10)
                'Stop_Sample()
                GC.Collect()
                System.Diagnostics.Process.GetCurrentProcess.Kill()
                Application.Exit()
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
#End Region
#Region "功能"
#Region "切换显示"
    Private Sub Form_show()
        TabControl1.SelectTab(0)
        TabPage_liaohao.Parent = Nothing
        TabPage_xiangji.Parent = Nothing
        TabPage_jiaozheng.Parent = Nothing
        TabPage_biaozhun.Parent = Nothing
        'TabPage_tongxun.Parent = Nothing
        ShowTSMI_btn.Visible = True        ''''''显示
        PasswordTSMI_btn.Visible = True    ''''''密码
        LogTMSI_btn.Visible = True         ''''''日志
        PartNumTSMI_btn.Visible = False    ''''''料号
        CamerasTSMI_btn.Visible = False     ''''''相机
        InspectTSMI_btn.Visible = True     ''''''检测
        wu.Visible = False                 ''''''校正
        StanderTSMI_btn.Visible = False    ''''''标准
        ConnetTSMI_btn.Visible = True     ''''''通讯
        ParameterTSMI_btn.Visible = True  ''''''参数
        SaveTSMI_btn.Visible = True        ''''''保存
    End Sub
    Private Sub ShowTSMI_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ShowTSMI_btn.Click
        Try
            TabControl1.SelectedTab = TabPage_xianshi
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub LogTMSI_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles LogTMSI_btn.Click
        Try
            TabControl1.SelectedTab = TabPage_rizhi
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub CamerasTSMI_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CamerasTSMI_btn.Click
        Try
            TabControl1.SelectedTab = TabPage_xiangji
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub CalibTSMI_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles wu.Click
        Try
            TabControl1.SelectedTab = TabPage_jiaozheng
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub StanderTSMI_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles StanderTSMI_btn.Click
        Try
            TabControl1.SelectedTab = TabPage_biaozhun
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub InspectTSMI_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles InspectTSMI_btn.Click
        Try
            If MessageBox.Show("确定要打开 工具设置 界面 吗？ 非专业人士请勿操作！", "提示信息", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = Windows.Forms.DialogResult.OK Then
                Dim form As New VPP(myJobManager)
                form.ShowDialog()
                form.Close()
                myJob0 = myJobManager.Job(0)
                myJob1 = myJobManager.Job(1)
                myJob2 = myJobManager.Job(2)
                myJob3 = myJobManager.Job(3)
            Else

            End If
            'TabControl1.SelectedTab = TabPage_jiance
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub ConnetTSMI_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ConnetTSMI_btn.Click
        Try
            TabControl1.SelectedTab = TabPage_tongxun
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub ParameterTSMI_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ParameterTSMI_btn.Click
        Try
            TabControl1.SelectedTab = TabPage_canshu
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub PartNumTSMI_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PartNumTSMI_btn.Click
        Try
            TabControl1.SelectedTab = TabPage_liaohao
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub PasswordTSMI_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PasswordTSMI_btn.Click
        Try
            ComboBox_Admin.SelectedIndex = 0
            TabControl1.SelectedTab = TabPage_mima
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub SaveTSMI_btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SaveTSMI_btn.Click
        Try
            If MessageBox.Show("确定要保存 相机VPP 吗？慎点！！！！！", "提示信息", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = Windows.Forms.DialogResult.OK Then
                CogSerializer.SaveObjectToFile(myJobManager, Application.StartupPath & "\VPP\PartNum\" + ComboBox1.Text + ".vpp", GetType(BinaryFormatter), CogSerializationOptionsConstants.Minimum)
                myJob0 = myJobManager.Job(0)
                myJob1 = myJobManager.Job(1)
                myJob2 = myJobManager.Job(2)
                myJob3 = myJobManager.Job(3)
                MessageBox.Show("保存VPP成功")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
#End Region
#Region "时钟"
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            ToolStripStatusLabel_Time.Text = Now.ToString
            Write_value(1, 5)
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
#End Region
#Region "注销"
    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        ParameterTSMI_btn.Enabled = False
        InspectTSMI_btn.Enabled = False
        SaveTSMI_btn.Enabled = False
        GroupBox_passworld.Visible = False
        ConnetTSMI_btn.Enabled = False
    End Sub
#End Region
#Region "登录"
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ComboBox_Admin.Text = PassworldArray(1, 0) And TextBox_passworld.Text = PassworldArray(1, 1) Then
            ParameterTSMI_btn.Enabled = True
        End If
        If ComboBox_Admin.Text = PassworldArray(2, 0) And TextBox_passworld.Text = PassworldArray(2, 1) Then
            ParameterTSMI_btn.Enabled = True
            InspectTSMI_btn.Enabled = True
            ConnetTSMI_btn.Enabled = True
        End If
        If ComboBox_Admin.Text = PassworldArray(3, 0) And TextBox_passworld.Text = PassworldArray(3, 1) Then
            ParameterTSMI_btn.Enabled = True
            InspectTSMI_btn.Enabled = True
            SaveTSMI_btn.Enabled = True
            GroupBox_passworld.Visible = True
            ConnetTSMI_btn.Enabled = True
        End If
    End Sub
#End Region
#Region "修改密码"
    Private Sub Button_passworld_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_passworld.Click
        PassworldArray(1, 1) = TextBox_passworld1.Text
        PassworldArray(2, 1) = TextBox_passworld2.Text
        PassworldArray(3, 1) = TextBox_passworld3.Text
        Save_passworld_Data()
    End Sub
#End Region
#End Region
#Region "作业加载"
#Region "检测加密狗"
    Delegate Sub mydelegate()
    Dim LicensedFeatures As New CogStringCollection
    Public Sub checklicense()
        Try
            If InvokeRequired Then
                Dim mydel As New mydelegate(AddressOf checklicense)
                Invoke(mydel)
                Return
            End If
            LicensedFeatures = CogMisc.GetLicensedFeatures(False)
            If LicensedFeatures.Count > 0 Then
            Else
                MessageBox.Show("请检查加密狗")
                Dim s() As System.Diagnostics.Process
                s = System.Diagnostics.Process.GetProcessesByName("MVision.vshost")
                System.Diagnostics.Process.GetCurrentProcess.Kill()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
#End Region
    Private Sub VPP_add(ByVal m As String)
        Try
            ComboBox1.Text = m
            myJobManager = CType(CogSerializer.LoadObjectFromFile("VPP\PartNum\" & m & ".vpp"), CogJobManager)
            myJob0 = myJobManager.Job(0)
            myIndependentJob0 = myJob0.OwnedIndependent
            myJob1 = myJobManager.Job(1)
            myIndependentJob1 = myJob1.OwnedIndependent
            myJob2 = myJobManager.Job(2)
            myIndependentJob2 = myJob2.OwnedIndependent
            myJob3 = myJobManager.Job(3)
            myIndependentJob3 = myJob3.OwnedIndependent
            'myJob4 = myJobManager.Job(4)
            'myIndependentJob4 = myJob4.OwnedIndependent
            myJobManager.UserQueueFlush()
            myJobManager.FailureQueueFlush()
            myJob0.ImageQueueFlush()
            myIndependentJob0.RealTimeQueueFlush()
            myJob1.ImageQueueFlush()
            myIndependentJob1.RealTimeQueueFlush()
            myJob2.ImageQueueFlush()
            myIndependentJob2.RealTimeQueueFlush()
            myJob3.ImageQueueFlush()
            myIndependentJob3.RealTimeQueueFlush()
            'myJob4.ImageQueueFlush()
            'myIndependentJob4.RealTimeQueueFlush()
            AddHandler myJobManager.Stopped, AddressOf myJobManager_Stopped
            AddHandler myJobManager.UserResultAvailable, AddressOf myJobManager_UserResultAvailable
            mGroup0 = CType(myJob0.VisionTool, CogToolGroup)
            mTool0 = CType(mGroup0.Tools("CogToolBlock1"), CogToolBlock)
            mGroup1 = CType(myJob1.VisionTool, CogToolGroup)
            mTool1 = CType(mGroup1.Tools("CogToolBlock1"), CogToolBlock)
            mGroup2 = CType(myJob2.VisionTool, CogToolGroup)
            mTool2 = CType(mGroup2.Tools("CogToolBlock1"), CogToolBlock)
            mGroup3 = CType(myJob3.VisionTool, CogToolGroup)
            mTool3 = CType(mGroup3.Tools("CogToolBlock1"), CogToolBlock)
            Add_data("CCD1")
            Add_data("CCD2")
            Add_data("CCD3")
            Add_data("CCD4")
            CogJobManagerEdit1.Subject = myJobManager
            Log_message("料号" & m & "加载成功")
        Catch ex As Exception
            Log_Error("VPP加载失败")
        End Try
    End Sub
    Private Sub myJobManager_Stopped(ByVal sender As Object, ByVal e As CogJobManagerActionEventArgs)
        Try
            If InvokeRequired Then
                Dim myDel As New myJobManagerDelegate(AddressOf myJobManager_Stopped)
                Dim eventArgs() As Object = {sender, e}
                Invoke(myDel, eventArgs)
                Return
            End If
            If CCD1_Result.Text = "OK" And CCD2_Result.Text = "OK" And CCD3_Result.Text = "OK" And CCD4_Result.Text = "OK" Then
                ' Write_value(1, 8)
                Count_OK += 1
            Else
                If CheckBox_Signal.Checked = False Then

                Else

                End If
                Count_NG += 1
            End If
            TextBox_OK.Text = Count_OK
            TextBox_NG.Text = Count_NG
            TextBox_Total.Text = Count_OK + Count_NG
            Log_message("VPP_Stop")
            ' Write_value(1, 7)
            GC.Collect()
        Catch ex As Exception
        End Try
    End Sub
#End Region
#Region "检测"
#Region "运行结果"
   
    Private Sub myJobManager_UserResultAvailable(ByVal sender As Object, ByVal e As CogJobManagerActionEventArgs)
        Try
            If InvokeRequired Then
                Dim myDel As New myJobManagerDelegate(AddressOf myJobManager_UserResultAvailable)
                Dim eventArgs() As Object = {sender, e}
                Invoke(myDel, eventArgs)
                Return
            End If
           
 
            Dim topRecord As Cognex.VisionPro.ICogRecord = myJobManager.UserResult
            Dim tmpRecord As Cognex.VisionPro.ICogRecord
            Dim jobname As String = topRecord.SubRecords("JobName").Content.ToString

            Select Case jobname
                Case myJob0.Name
                    Log_message("1")
                    If myJob0.RunStatus.Result = CogToolResultConstants.Accept Then
                        tmpRecord = topRecord.SubRecords("ShowLastRunRecordForUserQueue")
                        tmpRecord = tmpRecord.SubRecords("LastRun")
                        tmpRecord = tmpRecord.SubRecords("Image Source.OutputImage")
                        CogRecordDisplay2.Record = tmpRecord
                        CogRecordDisplay7.Record = tmpRecord
                        CCD1_Result.Text = "OK"
                        CCD1_Result.BackColor = Color.Green
                    ElseIf myJob0.RunStatus.Result = CogToolResultConstants.Reject Then
                        tmpRecord = topRecord.SubRecords("ShowLastRunRecordForUserQueue")
                        tmpRecord = tmpRecord.SubRecords("LastRun")
                        tmpRecord = tmpRecord.SubRecords("Image Source.OutputImage")
                        CogRecordDisplay2.Record = tmpRecord
                        CogRecordDisplay7.Record = tmpRecord
                        CCD1_Result.Text = "NG"
                        CCD1_Result.BackColor = Color.Red
                    Else
                        tmpRecord = topRecord.SubRecords("ShowLastRunRecordForUserQueue")
                        tmpRecord = tmpRecord.SubRecords("LastRun")
                        tmpRecord = tmpRecord.SubRecords("Image Source.OutputImage")
                        CogRecordDisplay2.Record = tmpRecord
                        CogRecordDisplay7.Record = tmpRecord
                        CCD1_Result.Text = "NG"
                        CCD1_Result.BackColor = Color.Red
                    End If
                    'CogRecordDisplay2.Record = tmpRecord
                    'CogRecordDisplay7.Record = tmpRecord
                    Thread.Sleep(200)
                    '   System.Threading.ThreadPool.QueueUserWorkItem(AddressOf Save_Image, "CCD1")
                    Dim th1 As Thread
                    th1 = New Threading.Thread(AddressOf SaveImage1)
                    th1.Start()

                Case myJob1.Name
                    Log_message("2")
                    If myJob1.RunStatus.Result = CogToolResultConstants.Accept Then
                        tmpRecord = topRecord.SubRecords("ShowLastRunRecordForUserQueue")
                        tmpRecord = tmpRecord.SubRecords("LastRun")
                        tmpRecord = tmpRecord.SubRecords("Image Source.OutputImage")
                        CogRecordDisplay4.Record = tmpRecord
                        CogRecordDisplay5.Record = tmpRecord
                        CCD2_Result.Text = "OK"
                        CCD2_Result.BackColor = Color.Green
                    ElseIf myJob1.RunStatus.Result = CogToolResultConstants.Reject Then
                        tmpRecord = topRecord.SubRecords("ShowLastRunRecordForUserQueue")
                        tmpRecord = tmpRecord.SubRecords("LastRun")
                        tmpRecord = tmpRecord.SubRecords("Image Source.OutputImage")
                        CogRecordDisplay4.Record = tmpRecord
                        CogRecordDisplay5.Record = tmpRecord
                        CCD2_Result.Text = "NG"
                        CCD2_Result.BackColor = Color.Red
                    Else
                        tmpRecord = topRecord.SubRecords("ShowLastRunRecordForUserQueue")
                        tmpRecord = tmpRecord.SubRecords("LastRun")
                        tmpRecord = tmpRecord.SubRecords("Image Source.OutputImage")
                        CogRecordDisplay4.Record = tmpRecord
                        CogRecordDisplay5.Record = tmpRecord
                        CCD2_Result.Text = "NG"
                        CCD2_Result.BackColor = Color.Red
                    End If
                   
                    Thread.Sleep(200)
                    ' System.Threading.ThreadPool.QueueUserWorkItem(AddressOf SaveImage, "CCD2")

                    '  SaveImage3(ComboBox1.Text, Label173.Text, ID)
                    Dim th3 As Thread
                    th3 = New Threading.Thread(AddressOf SaveImage3)
                    th3.Start()
                Case myJob2.Name
                    Log_message("3")
                    If myJob2.RunStatus.Result = CogToolResultConstants.Accept Then
                        tmpRecord = topRecord.SubRecords("ShowLastRunRecordForUserQueue")
                        tmpRecord = tmpRecord.SubRecords("LastRun")
                        tmpRecord = tmpRecord.SubRecords("Image Source.OutputImage")
                        CogRecordDisplay1.Record = tmpRecord
                        CogRecordDisplay8.Record = tmpRecord
                        CCD3_Result.Text = "OK"
                        CCD3_Result.BackColor = Color.Green
                    ElseIf myJob2.RunStatus.Result = CogToolResultConstants.Reject Then
                        tmpRecord = topRecord.SubRecords("ShowLastRunRecordForUserQueue")
                        tmpRecord = tmpRecord.SubRecords("LastRun")
                        tmpRecord = tmpRecord.SubRecords("Image Source.OutputImage")
                        CogRecordDisplay1.Record = tmpRecord
                        CogRecordDisplay8.Record = tmpRecord
                        CCD3_Result.Text = "NG"
                        CCD3_Result.BackColor = Color.Red
                    Else
                        tmpRecord = topRecord.SubRecords("ShowLastRunRecordForUserQueue")
                        tmpRecord = tmpRecord.SubRecords("LastRun")
                        tmpRecord = tmpRecord.SubRecords("Image Source.OutputImage")
                        CogRecordDisplay1.Record = tmpRecord
                        CogRecordDisplay8.Record = tmpRecord
                        CCD3_Result.Text = "NG"
                        CCD3_Result.BackColor = Color.Red
                    End If
                    'CogRecordDisplay1.Record = tmpRecord
                    'CogRecordDisplay8.Record = tmpRecord
                    Thread.Sleep(200)
                    ' SaveImage(ComboBox1.Text, Label173.Text, ID)
                    Dim th As Thread
                    th = New Threading.Thread(AddressOf SaveImage)
                    th.Start()

                    '  System.Threading.ThreadPool.QueueUserWorkItem(AddressOf Save_Image, "CCD3")
                Case myJob3.Name
                    Log_message("4")
                    If myJob3.RunStatus.Result = CogToolResultConstants.Accept Then
                        tmpRecord = topRecord.SubRecords("ShowLastRunRecordForUserQueue")
                        tmpRecord = tmpRecord.SubRecords("LastRun")
                        tmpRecord = tmpRecord.SubRecords("Image Source.OutputImage")
                        CogRecordDisplay3.Record = tmpRecord
                        CogRecordDisplay6.Record = tmpRecord
                        CCD4_Result.Text = "OK"
                        CCD4_Result.BackColor = Color.Green
                    ElseIf myJob3.RunStatus.Result = CogToolResultConstants.Reject Then
                        tmpRecord = topRecord.SubRecords("ShowLastRunRecordForUserQueue")
                        tmpRecord = tmpRecord.SubRecords("LastRun")
                        tmpRecord = tmpRecord.SubRecords("Image Source.OutputImage")
                        CogRecordDisplay3.Record = tmpRecord
                        CogRecordDisplay6.Record = tmpRecord
                        CCD4_Result.Text = "NG"
                        CCD4_Result.BackColor = Color.Red
                    Else
                        tmpRecord = topRecord.SubRecords("ShowLastRunRecordForUserQueue")
                        tmpRecord = tmpRecord.SubRecords("LastRun")
                        tmpRecord = tmpRecord.SubRecords("Image Source.OutputImage")
                        CogRecordDisplay3.Record = tmpRecord
                        CogRecordDisplay6.Record = tmpRecord
                        CCD4_Result.Text = "NG"
                        CCD4_Result.BackColor = Color.Red
                    End If
                    'CogRecordDisplay3.Record = tmpRecord
                    'CogRecordDisplay6.Record = tmpRecord
                    Thread.Sleep(200)
                    ' System.Threading.ThreadPool.QueueUserWorkItem(AddressOf Save_Image, "CCD4")
                    ' SaveImage2()

                    Dim th2 As Thread
                    th2 = New Threading.Thread(AddressOf SaveImage2)
                    th2.Start()


                Case myJob4.Name
                    If myJob4.RunStatus.Result = CogToolResultConstants.Accept Then
                        quantao = topRecord.SubRecords("Count").Content.ToString
                        If quantao = "1" Then
                            Label175.BackColor = Color.Red
                            Label175.Text = "存在圈套"
                            Write_value(False, 14)
                            Write_value(True, 15)
                        Else
                            Label175.BackColor = Color.Green
                            Label175.Text = "OK"
                            Write_value(False, 15)
                            Write_value(True, 14)
                          
                        End If
                    Else
                        Label175.BackColor = Color.Green
                        Label175.Text = "OK"
                        Write_value(False, 15)
                        Write_value(True, 14)
                     
                       
                    End If
            End Select

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
#End Region
    Private Sub Send_PLC(ByVal k As String)
       
    End Sub
#End Region
#Region "Log"
    Private Sub Log_message(ByVal message_string As String)
        Try
            If Log1.Lines.Count > 200 Then
                Log1.Clear()
            End If
            Log1.AppendText(Now.ToString & "--->" & message_string & vbCrLf)
        Catch ex As Exception
            Log_Error("信息记录错误")
        End Try
    End Sub
    Private Sub Log_Error(ByVal Error_string As String)
        Try
            If Log2.Lines.Count > 200 Then
                Log2.Clear()
            End If
            Log2.AppendText(Now.ToString & "--->" & Error_string)
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

#End Region
#Region "数据读写"
    Private Sub Read_Data()
        Read_passworld_Data()
        Read_CCD_Data()
        Read_PN()
        'Read_IP_Data()

    End Sub
    Private Sub Read_passworld_Data()
        Try
            Dim str_SR_Info As String
            Dim Read_SR_info As StreamReader = New StreamReader(Application.StartupPath & "\Data\passworld.txt", Encoding.Default)
            str_SR_Info = Read_SR_info.ReadToEnd().Replace(vbCrLf, "@") '.Replace("]", "\r\n").Replace("\r\n", "=")
            str_Info = str_SR_Info.Split("]")
            Passworld = str_Info(1).Split("@")
            For index = 1 To Passworld.Count - 2
                For j = 0 To 1
                    PassworldArray(index, j) = Passworld(index).Split("=")(j)
                Next
            Next
            Read_SR_info.Close()
            TextBox_passworld1.Text = PassworldArray(1, 1)
            TextBox_passworld2.Text = PassworldArray(2, 1)
            TextBox_passworld3.Text = PassworldArray(3, 1)
            Log_message("密码获取成功")
        Catch ex As Exception
            MessageBox.Show("参数读取失败")
        End Try
    End Sub

    Private Sub Save_passworld_Data()
        Try
            Dim fname As String = Application.StartupPath & "\Data\passworld.txt"
            Dim fname1 As New FileStream(fname, FileMode.OpenOrCreate, FileAccess.Write)
            Dim Save As New StreamWriter(fname1, System.Text.Encoding.GetEncoding("GB2312"))
            Save.Flush()
            Save.WriteLine("[Passworld]")
            Save.WriteLine("操作员=" + PassworldArray(1, 1))
            Save.WriteLine("管理员=" + PassworldArray(2, 1))
            Save.WriteLine("程序员=" + PassworldArray(3, 1))
            Save.Flush()
            Save.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub Read_PN()
        Try
            ComboBox1.Items.Clear()
            Dim str_SR_Info As String
            Dim Read_SR_info As StreamReader = New StreamReader(Application.StartupPath & "\Data\PN.txt", Encoding.Default)
            str_SR_Info = Read_SR_info.ReadToEnd().Replace(vbCrLf, "@") '.Replace("]", "\r\n").Replace("\r\n", "=")
            str_Info = str_SR_Info.Split("]")
            PN = str_Info(1).Split("@")
            Current_PN = PN(1)
            PN = str_Info(2).Split("@")
            For index = 1 To Passworld.Count - 1
                For j = 0 To 1
                    PNArray(index, j) = PN(index)
                Next
                ComboBox1.Items.Add(PN(index))
            Next
            Read_SR_info.Close()
            ComboBox1.Text = Current_PN
            Log_message("料号获取成功")
        Catch ex As Exception
            MessageBox.Show("参数读取失败")
        End Try
    End Sub
    Public Sub Read_CCD_Data()
        Try
            Dim str_SR_Info1 As String
            Dim Read_SR_info1 As StreamReader = New StreamReader(Application.StartupPath & "\Data\Camera1.txt", Encoding.Default)
            str_SR_Info1 = Read_SR_info1.ReadToEnd().Replace(vbCrLf, "@") '.Replace("]", "\r\n").Replace("\r\n", "=")
            str_Info = str_SR_Info1.Split("]")
            CCD1_Data = str_Info(1).Split("@")
            For index = 1 To 38
                For j = 0 To 1
                    CCD1_DataArray(index, j) = CCD1_Data(index).Split("=")(j)
                Next
            Next
            Read_SR_info1.Close()
            NumericUpDown_CCD1_Resolution.Value = CDbl(CCD1_DataArray(1, 1))
            NumericUpDown_CCD1_NG_KZ.Value = CDbl(CCD1_DataArray(2, 1))
            NumericUpDown_CCD1_Part1_1.Value = CDbl(CCD1_DataArray(3, 1))
            NumericUpDown_CCD1_Part1_2.Value = CDbl(CCD1_DataArray(4, 1))
            NumericUpDown_CCD1_Part1_3.Value = CDbl(CCD1_DataArray(5, 1))
            NumericUpDown_CCD1_Part2_1.Value = CDbl(CCD1_DataArray(6, 1))
            NumericUpDown_CCD1_Part2_2.Value = CDbl(CCD1_DataArray(7, 1))
            NumericUpDown_CCD1_Part2_3.Value = CDbl(CCD1_DataArray(8, 1))
            NumericUpDown_CCD1_Part3_1.Value = CDbl(CCD1_DataArray(9, 1))
            NumericUpDown_CCD1_Part3_2.Value = CDbl(CCD1_DataArray(10, 1))
            NumericUpDown_CCD1_Part3_3.Value = CDbl(CCD1_DataArray(11, 1))
            NumericUpDown_CCD1_Part4_1.Value = CDbl(CCD1_DataArray(12, 1))
            NumericUpDown_CCD1_Part4_2.Value = CDbl(CCD1_DataArray(13, 1))
            NumericUpDown_CCD1_Part4_3.Value = CDbl(CCD1_DataArray(14, 1))
            NumericUpDown_CCD1_Part5_1.Value = CDbl(CCD1_DataArray(15, 1))
            NumericUpDown_CCD1_Part5_2.Value = CDbl(CCD1_DataArray(16, 1))
            NumericUpDown_CCD1_Part5_3.Value = CDbl(CCD1_DataArray(17, 1))
            NumericUpDown_CCD1_Part6_1.Value = CDbl(CCD1_DataArray(18, 1))
            NumericUpDown_CCD1_Part6_2.Value = CDbl(CCD1_DataArray(19, 1))
            NumericUpDown_CCD1_Part6_3.Value = CDbl(CCD1_DataArray(20, 1))
            NumericUpDown_CCD1_Part7_1.Value = CDbl(CCD1_DataArray(21, 1))
            NumericUpDown_CCD1_Part7_2.Value = CDbl(CCD1_DataArray(22, 1))
            NumericUpDown_CCD1_Part7_3.Value = CDbl(CCD1_DataArray(23, 1))
            NumericUpDown_CCD1_Part8_1.Value = CDbl(CCD1_DataArray(24, 1))
            NumericUpDown_CCD1_Part8_2.Value = CDbl(CCD1_DataArray(25, 1))
            NumericUpDown_CCD1_Part8_3.Value = CDbl(CCD1_DataArray(26, 1))
            NumericUpDown_CCD1_Part9_1.Value = CDbl(CCD1_DataArray(27, 1))
            NumericUpDown_CCD1_Part9_2.Value = CDbl(CCD1_DataArray(28, 1))
            NumericUpDown_CCD1_Part9_3.Value = CDbl(CCD1_DataArray(29, 1))
            NumericUpDown_CCD1_Part10_1.Value = CDbl(CCD1_DataArray(30, 1))
            NumericUpDown_CCD1_Part10_2.Value = CDbl(CCD1_DataArray(31, 1))
            NumericUpDown_CCD1_Part10_3.Value = CDbl(CCD1_DataArray(32, 1))
            NumericUpDown_CCD1_Part11_1.Value = CDbl(CCD1_DataArray(33, 1))
            NumericUpDown_CCD1_Part11_2.Value = CDbl(CCD1_DataArray(34, 1))
            NumericUpDown_CCD1_Part11_3.Value = CDbl(CCD1_DataArray(35, 1))
            NumericUpDown_CCD1_Part12_1.Value = CDbl(CCD1_DataArray(36, 1))
            NumericUpDown_CCD1_Part12_2.Value = CDbl(CCD1_DataArray(37, 1))
            NumericUpDown_CCD1_Part12_3.Value = CDbl(CCD1_DataArray(38, 1))
        Catch ex As Exception
            MessageBox.Show("Camera1参数读取失败")
        End Try
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Try
            Dim str_SR_Info2 As String
            Dim Read_SR_info2 As StreamReader = New StreamReader(Application.StartupPath & "\Data\Camera2.txt", Encoding.Default)
            str_SR_Info2 = Read_SR_info2.ReadToEnd().Replace(vbCrLf, "@") '.Replace("]", "\r\n").Replace("\r\n", "=")
            str_Info = str_SR_Info2.Split("]")
            CCD2_Data = str_Info(1).Split("@")
            For index = 1 To 38
                For j = 0 To 1
                    CCD2_DataArray(index, j) = CCD2_Data(index).Split("=")(j)
                Next
            Next
            Read_SR_info2.Close()
            NumericUpDown_CCD2_Resolution.Value = CDbl(CCD2_DataArray(1, 1))
            NumericUpDown_CCD2_NG_KZ.Value = CDbl(CCD2_DataArray(2, 1))
            NumericUpDown_CCD2_Part1_1.Value = CDbl(CCD2_DataArray(3, 1))
            NumericUpDown_CCD2_Part1_2.Value = CDbl(CCD2_DataArray(4, 1))
            NumericUpDown_CCD2_Part1_3.Value = CDbl(CCD2_DataArray(5, 1))
            NumericUpDown_CCD2_Part2_1.Value = CDbl(CCD2_DataArray(6, 1))
            NumericUpDown_CCD2_Part2_2.Value = CDbl(CCD2_DataArray(7, 1))
            NumericUpDown_CCD2_Part2_3.Value = CDbl(CCD2_DataArray(8, 1))
            NumericUpDown_CCD2_Part3_1.Value = CDbl(CCD2_DataArray(9, 1))
            NumericUpDown_CCD2_Part3_2.Value = CDbl(CCD2_DataArray(10, 1))
            NumericUpDown_CCD2_Part3_3.Value = CDbl(CCD2_DataArray(11, 1))
            NumericUpDown_CCD2_Part4_1.Value = CDbl(CCD2_DataArray(12, 1))
            NumericUpDown_CCD2_Part4_2.Value = CDbl(CCD2_DataArray(13, 1))
            NumericUpDown_CCD2_Part4_3.Value = CDbl(CCD2_DataArray(14, 1))
            NumericUpDown_CCD2_Part5_1.Value = CDbl(CCD2_DataArray(15, 1))
            NumericUpDown_CCD2_Part5_2.Value = CDbl(CCD2_DataArray(16, 1))
            NumericUpDown_CCD2_Part5_3.Value = CDbl(CCD2_DataArray(17, 1))
            NumericUpDown_CCD2_Part6_1.Value = CDbl(CCD2_DataArray(18, 1))
            NumericUpDown_CCD2_Part6_2.Value = CDbl(CCD2_DataArray(19, 1))
            NumericUpDown_CCD2_Part6_3.Value = CDbl(CCD2_DataArray(20, 1))
            NumericUpDown_CCD2_Part7_1.Value = CDbl(CCD2_DataArray(21, 1))
            NumericUpDown_CCD2_Part7_2.Value = CDbl(CCD2_DataArray(22, 1))
            NumericUpDown_CCD2_Part7_3.Value = CDbl(CCD2_DataArray(23, 1))
            NumericUpDown_CCD2_Part8_1.Value = CDbl(CCD2_DataArray(24, 1))
            NumericUpDown_CCD2_Part8_2.Value = CDbl(CCD2_DataArray(25, 1))
            NumericUpDown_CCD2_Part8_3.Value = CDbl(CCD2_DataArray(26, 1))
            NumericUpDown_CCD2_Part9_1.Value = CDbl(CCD2_DataArray(27, 1))
            NumericUpDown_CCD2_Part9_2.Value = CDbl(CCD2_DataArray(28, 1))
            NumericUpDown_CCD2_Part9_3.Value = CDbl(CCD2_DataArray(29, 1))
            NumericUpDown_CCD2_Part10_1.Value = CDbl(CCD2_DataArray(30, 1))
            NumericUpDown_CCD2_Part10_2.Value = CDbl(CCD2_DataArray(31, 1))
            NumericUpDown_CCD2_Part10_3.Value = CDbl(CCD2_DataArray(32, 1))
            NumericUpDown_CCD2_Part11_1.Value = CDbl(CCD2_DataArray(33, 1))
            NumericUpDown_CCD2_Part11_2.Value = CDbl(CCD2_DataArray(34, 1))
            NumericUpDown_CCD2_Part11_3.Value = CDbl(CCD2_DataArray(35, 1))
            NumericUpDown_CCD2_Part12_1.Value = CDbl(CCD2_DataArray(36, 1))
            NumericUpDown_CCD2_Part12_2.Value = CDbl(CCD2_DataArray(37, 1))
            NumericUpDown_CCD2_Part12_3.Value = CDbl(CCD2_DataArray(38, 1))
        Catch ex As Exception
            MessageBox.Show("Camera2参数读取失败")
        End Try
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Try
            Dim str_SR_Info3 As String
            Dim Read_SR_info3 As StreamReader = New StreamReader(Application.StartupPath & "\Data\Camera3.txt", Encoding.Default)
            str_SR_Info3 = Read_SR_info3.ReadToEnd().Replace(vbCrLf, "@") '.Replace("]", "\r\n").Replace("\r\n", "=")
            str_Info = str_SR_Info3.Split("]")
            CCD3_Data = str_Info(1).Split("@")
            For index = 1 To 38
                For j = 0 To 1
                    CCD3_DataArray(index, j) = CCD3_Data(index).Split("=")(j)
                Next
            Next
            Read_SR_info3.Close()
            NumericUpDown_CCD3_Resolution.Value = CDbl(CCD3_DataArray(1, 1))
            NumericUpDown_CCD3_NG_KZ.Value = CDbl(CCD3_DataArray(2, 1))
            NumericUpDown_CCD3_Part1_1.Value = CDbl(CCD3_DataArray(3, 1))
            NumericUpDown_CCD3_Part1_2.Value = CDbl(CCD3_DataArray(4, 1))
            NumericUpDown_CCD3_Part1_3.Value = CDbl(CCD3_DataArray(5, 1))
            NumericUpDown_CCD3_Part2_1.Value = CDbl(CCD3_DataArray(6, 1))
            NumericUpDown_CCD3_Part2_2.Value = CDbl(CCD3_DataArray(7, 1))
            NumericUpDown_CCD3_Part2_3.Value = CDbl(CCD3_DataArray(8, 1))
            NumericUpDown_CCD3_Part3_1.Value = CDbl(CCD3_DataArray(9, 1))
            NumericUpDown_CCD3_Part3_2.Value = CDbl(CCD3_DataArray(10, 1))
            NumericUpDown_CCD3_Part3_3.Value = CDbl(CCD3_DataArray(11, 1))
            NumericUpDown_CCD3_Part4_1.Value = CDbl(CCD3_DataArray(12, 1))
            NumericUpDown_CCD3_Part4_2.Value = CDbl(CCD3_DataArray(13, 1))
            NumericUpDown_CCD3_Part4_3.Value = CDbl(CCD3_DataArray(14, 1))
            NumericUpDown_CCD3_Part5_1.Value = CDbl(CCD3_DataArray(15, 1))
            NumericUpDown_CCD3_Part5_2.Value = CDbl(CCD3_DataArray(16, 1))
            NumericUpDown_CCD3_Part5_3.Value = CDbl(CCD3_DataArray(17, 1))
            NumericUpDown_CCD3_Part6_1.Value = CDbl(CCD3_DataArray(18, 1))
            NumericUpDown_CCD3_Part6_2.Value = CDbl(CCD3_DataArray(19, 1))
            NumericUpDown_CCD3_Part6_3.Value = CDbl(CCD3_DataArray(20, 1))
            NumericUpDown_CCD3_Part7_1.Value = CDbl(CCD3_DataArray(21, 1))
            NumericUpDown_CCD3_Part7_2.Value = CDbl(CCD3_DataArray(22, 1))
            NumericUpDown_CCD3_Part7_3.Value = CDbl(CCD3_DataArray(23, 1))
            NumericUpDown_CCD3_Part8_1.Value = CDbl(CCD3_DataArray(24, 1))
            NumericUpDown_CCD3_Part8_2.Value = CDbl(CCD3_DataArray(25, 1))
            NumericUpDown_CCD3_Part8_3.Value = CDbl(CCD3_DataArray(26, 1))
            NumericUpDown_CCD3_Part9_1.Value = CDbl(CCD3_DataArray(27, 1))
            NumericUpDown_CCD3_Part9_2.Value = CDbl(CCD3_DataArray(28, 1))
            NumericUpDown_CCD3_Part9_3.Value = CDbl(CCD3_DataArray(29, 1))
            NumericUpDown_CCD3_Part10_1.Value = CDbl(CCD3_DataArray(30, 1))
            NumericUpDown_CCD3_Part10_2.Value = CDbl(CCD3_DataArray(31, 1))
            NumericUpDown_CCD3_Part10_3.Value = CDbl(CCD3_DataArray(32, 1))
            NumericUpDown_CCD3_Part11_1.Value = CDbl(CCD3_DataArray(33, 1))
            NumericUpDown_CCD3_Part11_2.Value = CDbl(CCD3_DataArray(34, 1))
            NumericUpDown_CCD3_Part11_3.Value = CDbl(CCD3_DataArray(35, 1))
            NumericUpDown_CCD3_Part12_1.Value = CDbl(CCD3_DataArray(36, 1))
            NumericUpDown_CCD3_Part12_2.Value = CDbl(CCD3_DataArray(37, 1))
            NumericUpDown_CCD3_Part12_3.Value = CDbl(CCD3_DataArray(38, 1))
        Catch ex As Exception
            MessageBox.Show("Camera3参数读取失败")
        End Try
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Try
            Dim str_SR_Info4 As String
            Dim Read_SR_info4 As StreamReader = New StreamReader(Application.StartupPath & "\Data\Camera4.txt", Encoding.Default)
            str_SR_Info4 = Read_SR_info4.ReadToEnd().Replace(vbCrLf, "@") '.Replace("]", "\r\n").Replace("\r\n", "=")
            str_Info = str_SR_Info4.Split("]")
            CCD4_Data = str_Info(1).Split("@")
            'For index = 1 To CCD4_Data.Count - 2
            For index = 1 To 38
                For j = 0 To 1
                    CCD4_DataArray(index, j) = CCD4_Data(index).Split("=")(j)
                Next
            Next
            Read_SR_info4.Close()
            NumericUpDown_CCD4_Resolution.Value = CDbl(CCD4_DataArray(1, 1))
            NumericUpDown_CCD4_NG_KZ.Value = CDbl(CCD4_DataArray(2, 1))
            NumericUpDown_CCD4_Part1_1.Value = CDbl(CCD4_DataArray(3, 1))
            NumericUpDown_CCD4_Part1_2.Value = CDbl(CCD4_DataArray(4, 1))
            NumericUpDown_CCD4_Part1_3.Value = CDbl(CCD4_DataArray(5, 1))
            NumericUpDown_CCD4_Part2_1.Value = CDbl(CCD4_DataArray(6, 1))
            NumericUpDown_CCD4_Part2_2.Value = CDbl(CCD4_DataArray(7, 1))
            NumericUpDown_CCD4_Part2_3.Value = CDbl(CCD4_DataArray(8, 1))
            NumericUpDown_CCD4_Part3_1.Value = CDbl(CCD4_DataArray(9, 1))
            NumericUpDown_CCD4_Part3_2.Value = CDbl(CCD4_DataArray(10, 1))
            NumericUpDown_CCD4_Part3_3.Value = CDbl(CCD4_DataArray(11, 1))
            NumericUpDown_CCD4_Part4_1.Value = CDbl(CCD4_DataArray(12, 1))
            NumericUpDown_CCD4_Part4_2.Value = CDbl(CCD4_DataArray(13, 1))
            NumericUpDown_CCD4_Part4_3.Value = CDbl(CCD4_DataArray(14, 1))
            NumericUpDown_CCD4_Part5_1.Value = CDbl(CCD4_DataArray(15, 1))
            NumericUpDown_CCD4_Part5_2.Value = CDbl(CCD4_DataArray(16, 1))
            NumericUpDown_CCD4_Part5_3.Value = CDbl(CCD4_DataArray(17, 1))
            NumericUpDown_CCD4_Part6_1.Value = CDbl(CCD4_DataArray(18, 1))
            NumericUpDown_CCD4_Part6_2.Value = CDbl(CCD4_DataArray(19, 1))
            NumericUpDown_CCD4_Part6_3.Value = CDbl(CCD4_DataArray(20, 1))
            NumericUpDown_CCD4_Part7_1.Value = CDbl(CCD4_DataArray(21, 1))
            NumericUpDown_CCD4_Part7_2.Value = CDbl(CCD4_DataArray(22, 1))
            NumericUpDown_CCD4_Part7_3.Value = CDbl(CCD4_DataArray(23, 1))
            NumericUpDown_CCD4_Part8_1.Value = CDbl(CCD4_DataArray(24, 1))
            NumericUpDown_CCD4_Part8_2.Value = CDbl(CCD4_DataArray(25, 1))
            NumericUpDown_CCD4_Part8_3.Value = CDbl(CCD4_DataArray(26, 1))
            NumericUpDown_CCD4_Part9_1.Value = CDbl(CCD4_DataArray(27, 1))
            NumericUpDown_CCD4_Part9_2.Value = CDbl(CCD4_DataArray(28, 1))
            NumericUpDown_CCD4_Part9_3.Value = CDbl(CCD4_DataArray(29, 1))
            NumericUpDown_CCD4_Part10_1.Value = CDbl(CCD4_DataArray(30, 1))
            NumericUpDown_CCD4_Part10_2.Value = CDbl(CCD4_DataArray(31, 1))
            NumericUpDown_CCD4_Part10_3.Value = CDbl(CCD4_DataArray(32, 1))
            NumericUpDown_CCD4_Part11_1.Value = CDbl(CCD4_DataArray(33, 1))
            NumericUpDown_CCD4_Part11_2.Value = CDbl(CCD4_DataArray(34, 1))
            NumericUpDown_CCD4_Part11_3.Value = CDbl(CCD4_DataArray(35, 1))
            NumericUpDown_CCD4_Part12_1.Value = CDbl(CCD4_DataArray(36, 1))
            NumericUpDown_CCD4_Part12_2.Value = CDbl(CCD4_DataArray(37, 1))
            NumericUpDown_CCD4_Part12_3.Value = CDbl(CCD4_DataArray(38, 1))
        Catch ex As Exception
            MessageBox.Show("Camera4参数读取失败")
        End Try
        Log_message("Camera参数获取成功")
    End Sub
    Private Sub Save_CCD_Data(ByVal CCD_Name As String)
        Try
            Select Case CCD_Name
                Case "CCD1"
                    Dim fname As String = Application.StartupPath & "\Data\Camera1.txt"
                    Dim fname1 As New FileStream(fname, FileMode.OpenOrCreate, FileAccess.Write)
                    Dim Save As New StreamWriter(fname1, System.Text.Encoding.GetEncoding("GB2312"))
                    Save.Flush()
                    Save.WriteLine("[CCD1]")
                    Save.WriteLine("Resolution=" + CCD1_DataArray(1, 1))
                    Save.WriteLine("KZ=" + CCD1_DataArray(2, 1))
                    Save.WriteLine("Part1_Width=" + CCD1_DataArray(3, 1))
                    Save.WriteLine("Part1_UP=" + CCD1_DataArray(4, 1))
                    Save.WriteLine("Part1_Down=" + CCD1_DataArray(5, 1))
                    Save.WriteLine("Part2_Width=" + CCD1_DataArray(6, 1))
                    Save.WriteLine("Part2_UP=" + CCD1_DataArray(7, 1))
                    Save.WriteLine("Part2_Down=" + CCD1_DataArray(8, 1))
                    Save.WriteLine("Part3_Width=" + CCD1_DataArray(9, 1))
                    Save.WriteLine("Part3_UP=" + CCD1_DataArray(10, 1))
                    Save.WriteLine("Part3_Down=" + CCD1_DataArray(11, 1))
                    Save.WriteLine("Part4_Width=" + CCD1_DataArray(12, 1))
                    Save.WriteLine("Part4_UP=" + CCD1_DataArray(13, 1))
                    Save.WriteLine("Part4_Down=" + CCD1_DataArray(14, 1))
                    Save.WriteLine("Part5_Width=" + CCD1_DataArray(15, 1))
                    Save.WriteLine("Part5_UP=" + CCD1_DataArray(16, 1))
                    Save.WriteLine("Part5_Down=" + CCD1_DataArray(17, 1))
                    Save.WriteLine("Part6_Width=" + CCD1_DataArray(18, 1))
                    Save.WriteLine("Part6_UP=" + CCD1_DataArray(19, 1))
                    Save.WriteLine("Part6_Down=" + CCD1_DataArray(20, 1))
                    Save.WriteLine("Part7_Width=" + CCD1_DataArray(21, 1))
                    Save.WriteLine("Part7_UP=" + CCD1_DataArray(22, 1))
                    Save.WriteLine("Part7_Down=" + CCD1_DataArray(23, 1))
                    Save.WriteLine("Part8_Width=" + CCD1_DataArray(24, 1))
                    Save.WriteLine("Part8_UP=" + CCD1_DataArray(25, 1))
                    Save.WriteLine("Part8_Down=" + CCD1_DataArray(26, 1))
                    Save.WriteLine("Part9_Width=" + CCD1_DataArray(27, 1))
                    Save.WriteLine("Part9_UP=" + CCD1_DataArray(28, 1))
                    Save.WriteLine("Part9_Down=" + CCD1_DataArray(29, 1))
                    Save.WriteLine("Part10_Width=" + CCD1_DataArray(30, 1))
                    Save.WriteLine("Part10_UP=" + CCD1_DataArray(31, 1))
                    Save.WriteLine("Part10_Down=" + CCD1_DataArray(32, 1))
                    Save.WriteLine("Part11_Width=" + CCD1_DataArray(33, 1))
                    Save.WriteLine("Part11_UP=" + CCD1_DataArray(34, 1))
                    Save.WriteLine("Part11_Down=" + CCD1_DataArray(35, 1))
                    Save.WriteLine("Part12_Width=" + CCD1_DataArray(36, 1))
                    Save.WriteLine("Part12_UP=" + CCD1_DataArray(37, 1))
                    Save.WriteLine("Part12_Down=" + CCD1_DataArray(38, 1))
                    Save.Flush()
                    Save.Close()
                Case "CCD2"
                    Dim fname As String = Application.StartupPath & "\Data\Camera2.txt"
                    Dim fname1 As New FileStream(fname, FileMode.OpenOrCreate, FileAccess.Write)
                    Dim Save As New StreamWriter(fname1, System.Text.Encoding.GetEncoding("GB2312"))
                    Save.Flush()
                    Save.WriteLine("[CCD2]")
                    Save.WriteLine("Resolution=" + CCD2_DataArray(1, 1))
                    Save.WriteLine("KZ=" + CCD2_DataArray(2, 1))
                    Save.WriteLine("Part1_Width=" + CCD2_DataArray(3, 1))
                    Save.WriteLine("Part1_UP=" + CCD2_DataArray(4, 1))
                    Save.WriteLine("Part1_Down=" + CCD2_DataArray(5, 1))
                    Save.WriteLine("Part2_Width=" + CCD2_DataArray(6, 1))
                    Save.WriteLine("Part2_UP=" + CCD2_DataArray(7, 1))
                    Save.WriteLine("Part2_Down=" + CCD2_DataArray(8, 1))
                    Save.WriteLine("Part3_Width=" + CCD2_DataArray(9, 1))
                    Save.WriteLine("Part3_UP=" + CCD2_DataArray(10, 1))
                    Save.WriteLine("Part3_Down=" + CCD2_DataArray(11, 1))
                    Save.WriteLine("Part4_Width=" + CCD2_DataArray(12, 1))
                    Save.WriteLine("Part4_UP=" + CCD2_DataArray(13, 1))
                    Save.WriteLine("Part4_Down=" + CCD2_DataArray(14, 1))
                    Save.WriteLine("Part5_Width=" + CCD2_DataArray(15, 1))
                    Save.WriteLine("Part5_UP=" + CCD2_DataArray(16, 1))
                    Save.WriteLine("Part5_Down=" + CCD2_DataArray(17, 1))
                    Save.WriteLine("Part6_Width=" + CCD2_DataArray(18, 1))
                    Save.WriteLine("Part6_UP=" + CCD2_DataArray(19, 1))
                    Save.WriteLine("Part6_Down=" + CCD2_DataArray(20, 1))
                    Save.WriteLine("Part7_Width=" + CCD2_DataArray(21, 1))
                    Save.WriteLine("Part7_UP=" + CCD2_DataArray(22, 1))
                    Save.WriteLine("Part7_Down=" + CCD2_DataArray(23, 1))
                    Save.WriteLine("Part8_Width=" + CCD2_DataArray(24, 1))
                    Save.WriteLine("Part8_UP=" + CCD2_DataArray(25, 1))
                    Save.WriteLine("Part8_Down=" + CCD2_DataArray(26, 1))
                    Save.WriteLine("Part9_Width=" + CCD2_DataArray(27, 1))
                    Save.WriteLine("Part9_UP=" + CCD2_DataArray(28, 1))
                    Save.WriteLine("Part9_Down=" + CCD2_DataArray(29, 1))
                    Save.WriteLine("Part10_Width=" + CCD2_DataArray(30, 1))
                    Save.WriteLine("Part10_UP=" + CCD2_DataArray(31, 1))
                    Save.WriteLine("Part10_Down=" + CCD2_DataArray(32, 1))
                    Save.WriteLine("Part11_Width=" + CCD2_DataArray(33, 1))
                    Save.WriteLine("Part11_UP=" + CCD2_DataArray(34, 1))
                    Save.WriteLine("Part11_Down=" + CCD2_DataArray(35, 1))
                    Save.WriteLine("Part12_Width=" + CCD2_DataArray(36, 1))
                    Save.WriteLine("Part12_UP=" + CCD2_DataArray(37, 1))
                    Save.WriteLine("Part12_Down=" + CCD2_DataArray(38, 1))
                    Save.Flush()
                    Save.Close()
                Case "CCD3"
                    Dim fname As String = Application.StartupPath & "\Data\Camera3.txt"
                    Dim fname1 As New FileStream(fname, FileMode.OpenOrCreate, FileAccess.Write)
                    Dim Save As New StreamWriter(fname1, System.Text.Encoding.GetEncoding("GB2312"))
                    Save.Flush()
                    Save.WriteLine("[CCD3]")
                    Save.WriteLine("Resolution=" + CCD3_DataArray(1, 1))
                    Save.WriteLine("KZ=" + CCD3_DataArray(2, 1))
                    Save.WriteLine("Part1_Width=" + CCD3_DataArray(3, 1))
                    Save.WriteLine("Part1_UP=" + CCD3_DataArray(4, 1))
                    Save.WriteLine("Part1_Down=" + CCD3_DataArray(5, 1))
                    Save.WriteLine("Part2_Width=" + CCD3_DataArray(6, 1))
                    Save.WriteLine("Part2_UP=" + CCD3_DataArray(7, 1))
                    Save.WriteLine("Part2_Down=" + CCD3_DataArray(8, 1))
                    Save.WriteLine("Part3_Width=" + CCD3_DataArray(9, 1))
                    Save.WriteLine("Part3_UP=" + CCD3_DataArray(10, 1))
                    Save.WriteLine("Part3_Down=" + CCD3_DataArray(11, 1))
                    Save.WriteLine("Part4_Width=" + CCD3_DataArray(12, 1))
                    Save.WriteLine("Part4_UP=" + CCD3_DataArray(13, 1))
                    Save.WriteLine("Part4_Down=" + CCD3_DataArray(14, 1))
                    Save.WriteLine("Part5_Width=" + CCD3_DataArray(15, 1))
                    Save.WriteLine("Part5_UP=" + CCD3_DataArray(16, 1))
                    Save.WriteLine("Part5_Down=" + CCD3_DataArray(17, 1))
                    Save.WriteLine("Part6_Width=" + CCD3_DataArray(18, 1))
                    Save.WriteLine("Part6_UP=" + CCD3_DataArray(19, 1))
                    Save.WriteLine("Part6_Down=" + CCD3_DataArray(20, 1))
                    Save.WriteLine("Part7_Width=" + CCD3_DataArray(21, 1))
                    Save.WriteLine("Part7_UP=" + CCD3_DataArray(22, 1))
                    Save.WriteLine("Part7_Down=" + CCD3_DataArray(23, 1))
                    Save.WriteLine("Part8_Width=" + CCD3_DataArray(24, 1))
                    Save.WriteLine("Part8_UP=" + CCD3_DataArray(25, 1))
                    Save.WriteLine("Part8_Down=" + CCD3_DataArray(26, 1))
                    Save.WriteLine("Part9_Width=" + CCD3_DataArray(27, 1))
                    Save.WriteLine("Part9_UP=" + CCD3_DataArray(28, 1))
                    Save.WriteLine("Part9_Down=" + CCD3_DataArray(29, 1))
                    Save.WriteLine("Part10_Width=" + CCD3_DataArray(30, 1))
                    Save.WriteLine("Part10_UP=" + CCD3_DataArray(31, 1))
                    Save.WriteLine("Part10_Down=" + CCD3_DataArray(32, 1))
                    Save.WriteLine("Part11_Width=" + CCD3_DataArray(33, 1))
                    Save.WriteLine("Part11_UP=" + CCD3_DataArray(34, 1))
                    Save.WriteLine("Part11_Down=" + CCD3_DataArray(35, 1))
                    Save.WriteLine("Part12_Width=" + CCD3_DataArray(36, 1))
                    Save.WriteLine("Part12_UP=" + CCD3_DataArray(37, 1))
                    Save.WriteLine("Part12_Down=" + CCD3_DataArray(38, 1))
                    Save.Flush()
                    Save.Close()
                Case "CCD4"
                    Dim fname As String = Application.StartupPath & "\Data\Camera4.txt"
                    Dim fname1 As New FileStream(fname, FileMode.OpenOrCreate, FileAccess.Write)
                    Dim Save As New StreamWriter(fname1, System.Text.Encoding.GetEncoding("GB2312"))
                    Save.Flush()
                    Save.WriteLine("[CCD4]")
                    Save.WriteLine("Resolution=" + CCD4_DataArray(1, 1))
                    Save.WriteLine("KZ=" + CCD4_DataArray(2, 1))
                    Save.WriteLine("Part1_Width=" + CCD4_DataArray(3, 1))
                    Save.WriteLine("Part1_UP=" + CCD4_DataArray(4, 1))
                    Save.WriteLine("Part1_Down=" + CCD4_DataArray(5, 1))
                    Save.WriteLine("Part2_Width=" + CCD4_DataArray(6, 1))
                    Save.WriteLine("Part2_UP=" + CCD4_DataArray(7, 1))
                    Save.WriteLine("Part2_Down=" + CCD4_DataArray(8, 1))
                    Save.WriteLine("Part3_Width=" + CCD4_DataArray(9, 1))
                    Save.WriteLine("Part3_UP=" + CCD4_DataArray(10, 1))
                    Save.WriteLine("Part3_Down=" + CCD4_DataArray(11, 1))
                    Save.WriteLine("Part4_Width=" + CCD4_DataArray(12, 1))
                    Save.WriteLine("Part4_UP=" + CCD4_DataArray(13, 1))
                    Save.WriteLine("Part4_Down=" + CCD4_DataArray(14, 1))
                    Save.WriteLine("Part5_Width=" + CCD4_DataArray(15, 1))
                    Save.WriteLine("Part5_UP=" + CCD4_DataArray(16, 1))
                    Save.WriteLine("Part5_Down=" + CCD4_DataArray(17, 1))
                    Save.WriteLine("Part6_Width=" + CCD4_DataArray(18, 1))
                    Save.WriteLine("Part6_UP=" + CCD4_DataArray(19, 1))
                    Save.WriteLine("Part6_Down=" + CCD4_DataArray(20, 1))
                    Save.WriteLine("Part7_Width=" + CCD4_DataArray(21, 1))
                    Save.WriteLine("Part7_UP=" + CCD4_DataArray(22, 1))
                    Save.WriteLine("Part7_Down=" + CCD4_DataArray(23, 1))
                    Save.WriteLine("Part8_Width=" + CCD4_DataArray(24, 1))
                    Save.WriteLine("Part8_UP=" + CCD4_DataArray(25, 1))
                    Save.WriteLine("Part8_Down=" + CCD4_DataArray(26, 1))
                    Save.WriteLine("Part9_Width=" + CCD4_DataArray(27, 1))
                    Save.WriteLine("Part9_UP=" + CCD4_DataArray(28, 1))
                    Save.WriteLine("Part9_Down=" + CCD4_DataArray(29, 1))
                    Save.WriteLine("Part10_Width=" + CCD4_DataArray(30, 1))
                    Save.WriteLine("Part10_UP=" + CCD4_DataArray(31, 1))
                    Save.WriteLine("Part10_Down=" + CCD4_DataArray(32, 1))
                    Save.WriteLine("Part11_Width=" + CCD4_DataArray(33, 1))
                    Save.WriteLine("Part11_UP=" + CCD4_DataArray(34, 1))
                    Save.WriteLine("Part11_Down=" + CCD4_DataArray(35, 1))
                    Save.WriteLine("Part12_Width=" + CCD4_DataArray(36, 1))
                    Save.WriteLine("Part12_UP=" + CCD4_DataArray(37, 1))
                    Save.WriteLine("Part12_Down=" + CCD4_DataArray(38, 1))
                    Save.Flush()
                    Save.Close()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        CCD1_DataArray(1, 1) = NumericUpDown_CCD1_Resolution.Value
        CCD1_DataArray(2, 1) = NumericUpDown_CCD1_NG_KZ.Value
        CCD1_DataArray(3, 1) = NumericUpDown_CCD1_Part1_1.Value
        CCD1_DataArray(4, 1) = NumericUpDown_CCD1_Part1_2.Value
        CCD1_DataArray(5, 1) = NumericUpDown_CCD1_Part1_3.Value
        CCD1_DataArray(6, 1) = NumericUpDown_CCD1_Part2_1.Value
        CCD1_DataArray(7, 1) = NumericUpDown_CCD1_Part2_2.Value
        CCD1_DataArray(8, 1) = NumericUpDown_CCD1_Part2_3.Value
        CCD1_DataArray(9, 1) = NumericUpDown_CCD1_Part3_1.Value
        CCD1_DataArray(10, 1) = NumericUpDown_CCD1_Part3_2.Value
        CCD1_DataArray(11, 1) = NumericUpDown_CCD1_Part3_3.Value
        CCD1_DataArray(12, 1) = NumericUpDown_CCD1_Part4_1.Value
        CCD1_DataArray(13, 1) = NumericUpDown_CCD1_Part4_2.Value
        CCD1_DataArray(14, 1) = NumericUpDown_CCD1_Part4_3.Value
        CCD1_DataArray(15, 1) = NumericUpDown_CCD1_Part5_1.Value
        CCD1_DataArray(16, 1) = NumericUpDown_CCD1_Part5_2.Value
        CCD1_DataArray(17, 1) = NumericUpDown_CCD1_Part5_3.Value
        CCD1_DataArray(18, 1) = NumericUpDown_CCD1_Part6_1.Value
        CCD1_DataArray(19, 1) = NumericUpDown_CCD1_Part6_2.Value
        CCD1_DataArray(20, 1) = NumericUpDown_CCD1_Part6_3.Value
        CCD1_DataArray(21, 1) = NumericUpDown_CCD1_Part7_1.Value
        CCD1_DataArray(22, 1) = NumericUpDown_CCD1_Part7_2.Value
        CCD1_DataArray(23, 1) = NumericUpDown_CCD1_Part7_3.Value
        CCD1_DataArray(24, 1) = NumericUpDown_CCD1_Part8_1.Value
        CCD1_DataArray(25, 1) = NumericUpDown_CCD1_Part8_2.Value
        CCD1_DataArray(26, 1) = NumericUpDown_CCD1_Part8_3.Value
        CCD1_DataArray(27, 1) = NumericUpDown_CCD1_Part9_1.Value
        CCD1_DataArray(28, 1) = NumericUpDown_CCD1_Part9_2.Value
        CCD1_DataArray(29, 1) = NumericUpDown_CCD1_Part9_3.Value
        CCD1_DataArray(30, 1) = NumericUpDown_CCD1_Part10_1.Value
        CCD1_DataArray(31, 1) = NumericUpDown_CCD1_Part10_2.Value
        CCD1_DataArray(32, 1) = NumericUpDown_CCD1_Part10_3.Value
        CCD1_DataArray(33, 1) = NumericUpDown_CCD1_Part11_1.Value
        CCD1_DataArray(34, 1) = NumericUpDown_CCD1_Part11_2.Value
        CCD1_DataArray(35, 1) = NumericUpDown_CCD1_Part11_3.Value
        CCD1_DataArray(36, 1) = NumericUpDown_CCD1_Part12_1.Value
        CCD1_DataArray(37, 1) = NumericUpDown_CCD1_Part12_2.Value
        CCD1_DataArray(38, 1) = NumericUpDown_CCD1_Part12_3.Value
        Save_CCD_Data("CCD1")
        Add_data("CCD1")
    End Sub
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        CCD2_DataArray(1, 1) = NumericUpDown_CCD2_Resolution.Value
        CCD2_DataArray(2, 1) = NumericUpDown_CCD2_NG_KZ.Value
        CCD2_DataArray(3, 1) = NumericUpDown_CCD2_Part1_1.Value
        CCD2_DataArray(4, 1) = NumericUpDown_CCD2_Part1_2.Value
        CCD2_DataArray(5, 1) = NumericUpDown_CCD2_Part1_3.Value
        CCD2_DataArray(6, 1) = NumericUpDown_CCD2_Part2_1.Value
        CCD2_DataArray(7, 1) = NumericUpDown_CCD2_Part2_2.Value
        CCD2_DataArray(8, 1) = NumericUpDown_CCD2_Part2_3.Value
        CCD2_DataArray(9, 1) = NumericUpDown_CCD2_Part3_1.Value
        CCD2_DataArray(10, 1) = NumericUpDown_CCD2_Part3_2.Value
        CCD2_DataArray(11, 1) = NumericUpDown_CCD2_Part3_3.Value
        CCD2_DataArray(12, 1) = NumericUpDown_CCD2_Part4_1.Value
        CCD2_DataArray(13, 1) = NumericUpDown_CCD2_Part4_2.Value
        CCD2_DataArray(14, 1) = NumericUpDown_CCD2_Part4_3.Value
        CCD2_DataArray(15, 1) = NumericUpDown_CCD2_Part5_1.Value
        CCD2_DataArray(16, 1) = NumericUpDown_CCD2_Part5_2.Value
        CCD2_DataArray(17, 1) = NumericUpDown_CCD2_Part5_3.Value
        CCD2_DataArray(18, 1) = NumericUpDown_CCD2_Part6_1.Value
        CCD2_DataArray(19, 1) = NumericUpDown_CCD2_Part6_2.Value
        CCD2_DataArray(20, 1) = NumericUpDown_CCD2_Part6_3.Value
        CCD2_DataArray(21, 1) = NumericUpDown_CCD2_Part7_1.Value
        CCD2_DataArray(22, 1) = NumericUpDown_CCD2_Part7_2.Value
        CCD2_DataArray(23, 1) = NumericUpDown_CCD2_Part7_3.Value
        CCD2_DataArray(24, 1) = NumericUpDown_CCD2_Part8_1.Value
        CCD2_DataArray(25, 1) = NumericUpDown_CCD2_Part8_2.Value
        CCD2_DataArray(26, 1) = NumericUpDown_CCD2_Part8_3.Value
        CCD2_DataArray(27, 1) = NumericUpDown_CCD2_Part9_1.Value
        CCD2_DataArray(28, 1) = NumericUpDown_CCD2_Part9_2.Value
        CCD2_DataArray(29, 1) = NumericUpDown_CCD2_Part9_3.Value
        CCD2_DataArray(30, 1) = NumericUpDown_CCD2_Part10_1.Value
        CCD2_DataArray(31, 1) = NumericUpDown_CCD2_Part10_2.Value
        CCD2_DataArray(32, 1) = NumericUpDown_CCD2_Part10_3.Value
        CCD2_DataArray(33, 1) = NumericUpDown_CCD2_Part11_1.Value
        CCD2_DataArray(34, 1) = NumericUpDown_CCD2_Part11_2.Value
        CCD2_DataArray(35, 1) = NumericUpDown_CCD2_Part11_3.Value
        CCD2_DataArray(36, 1) = NumericUpDown_CCD2_Part12_1.Value
        CCD2_DataArray(37, 1) = NumericUpDown_CCD2_Part12_2.Value
        CCD2_DataArray(38, 1) = NumericUpDown_CCD2_Part12_3.Value
        Save_CCD_Data("CCD2")
        Add_data("CCD2")
    End Sub
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        CCD3_DataArray(1, 1) = NumericUpDown_CCD3_Resolution.Value
        CCD3_DataArray(2, 1) = NumericUpDown_CCD3_NG_KZ.Value
        CCD3_DataArray(3, 1) = NumericUpDown_CCD3_Part1_1.Value
        CCD3_DataArray(4, 1) = NumericUpDown_CCD3_Part1_2.Value
        CCD3_DataArray(5, 1) = NumericUpDown_CCD3_Part1_3.Value
        CCD3_DataArray(6, 1) = NumericUpDown_CCD3_Part2_1.Value
        CCD3_DataArray(7, 1) = NumericUpDown_CCD3_Part2_2.Value
        CCD3_DataArray(8, 1) = NumericUpDown_CCD3_Part2_3.Value
        CCD3_DataArray(9, 1) = NumericUpDown_CCD3_Part3_1.Value
        CCD3_DataArray(10, 1) = NumericUpDown_CCD3_Part3_2.Value
        CCD3_DataArray(11, 1) = NumericUpDown_CCD3_Part3_3.Value
        CCD3_DataArray(12, 1) = NumericUpDown_CCD3_Part4_1.Value
        CCD3_DataArray(13, 1) = NumericUpDown_CCD3_Part4_2.Value
        CCD3_DataArray(14, 1) = NumericUpDown_CCD3_Part4_3.Value
        CCD3_DataArray(15, 1) = NumericUpDown_CCD3_Part5_1.Value
        CCD3_DataArray(16, 1) = NumericUpDown_CCD3_Part5_2.Value
        CCD3_DataArray(17, 1) = NumericUpDown_CCD3_Part5_3.Value
        CCD3_DataArray(18, 1) = NumericUpDown_CCD3_Part6_1.Value
        CCD3_DataArray(19, 1) = NumericUpDown_CCD3_Part6_2.Value
        CCD3_DataArray(20, 1) = NumericUpDown_CCD3_Part6_3.Value
        CCD3_DataArray(21, 1) = NumericUpDown_CCD3_Part7_1.Value
        CCD3_DataArray(22, 1) = NumericUpDown_CCD3_Part7_2.Value
        CCD3_DataArray(23, 1) = NumericUpDown_CCD3_Part7_3.Value
        CCD3_DataArray(24, 1) = NumericUpDown_CCD3_Part8_1.Value
        CCD3_DataArray(25, 1) = NumericUpDown_CCD3_Part8_2.Value
        CCD3_DataArray(26, 1) = NumericUpDown_CCD3_Part8_3.Value
        CCD3_DataArray(27, 1) = NumericUpDown_CCD3_Part9_1.Value
        CCD3_DataArray(28, 1) = NumericUpDown_CCD3_Part9_2.Value
        CCD3_DataArray(29, 1) = NumericUpDown_CCD3_Part9_3.Value
        CCD3_DataArray(30, 1) = NumericUpDown_CCD3_Part10_1.Value
        CCD3_DataArray(31, 1) = NumericUpDown_CCD3_Part10_2.Value
        CCD3_DataArray(32, 1) = NumericUpDown_CCD3_Part10_3.Value
        CCD3_DataArray(33, 1) = NumericUpDown_CCD3_Part11_1.Value
        CCD3_DataArray(34, 1) = NumericUpDown_CCD3_Part11_2.Value
        CCD3_DataArray(35, 1) = NumericUpDown_CCD3_Part11_3.Value
        CCD3_DataArray(36, 1) = NumericUpDown_CCD3_Part12_1.Value
        CCD3_DataArray(37, 1) = NumericUpDown_CCD3_Part12_2.Value
        CCD3_DataArray(38, 1) = NumericUpDown_CCD3_Part12_3.Value
        Save_CCD_Data("CCD3")
        Add_data("CCD3")
    End Sub
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        CCD4_DataArray(1, 1) = NumericUpDown_CCD4_Resolution.Value
        CCD4_DataArray(2, 1) = NumericUpDown_CCD4_NG_KZ.Value
        CCD4_DataArray(3, 1) = NumericUpDown_CCD4_Part1_1.Value
        CCD4_DataArray(4, 1) = NumericUpDown_CCD4_Part1_2.Value
        CCD4_DataArray(5, 1) = NumericUpDown_CCD4_Part1_3.Value
        CCD4_DataArray(6, 1) = NumericUpDown_CCD4_Part2_1.Value
        CCD4_DataArray(7, 1) = NumericUpDown_CCD4_Part2_2.Value
        CCD4_DataArray(8, 1) = NumericUpDown_CCD4_Part2_3.Value
        CCD4_DataArray(9, 1) = NumericUpDown_CCD4_Part3_1.Value
        CCD4_DataArray(10, 1) = NumericUpDown_CCD4_Part3_2.Value
        CCD4_DataArray(11, 1) = NumericUpDown_CCD4_Part3_3.Value
        CCD4_DataArray(12, 1) = NumericUpDown_CCD4_Part4_1.Value
        CCD4_DataArray(13, 1) = NumericUpDown_CCD4_Part4_2.Value
        CCD4_DataArray(14, 1) = NumericUpDown_CCD4_Part4_3.Value
        CCD4_DataArray(15, 1) = NumericUpDown_CCD4_Part5_1.Value
        CCD4_DataArray(16, 1) = NumericUpDown_CCD4_Part5_2.Value
        CCD4_DataArray(17, 1) = NumericUpDown_CCD4_Part5_3.Value
        CCD4_DataArray(18, 1) = NumericUpDown_CCD4_Part6_1.Value
        CCD4_DataArray(19, 1) = NumericUpDown_CCD4_Part6_2.Value
        CCD4_DataArray(20, 1) = NumericUpDown_CCD4_Part6_3.Value
        CCD4_DataArray(21, 1) = NumericUpDown_CCD4_Part7_1.Value
        CCD4_DataArray(22, 1) = NumericUpDown_CCD4_Part7_2.Value
        CCD4_DataArray(23, 1) = NumericUpDown_CCD4_Part7_3.Value
        CCD4_DataArray(24, 1) = NumericUpDown_CCD4_Part8_1.Value
        CCD4_DataArray(25, 1) = NumericUpDown_CCD4_Part8_2.Value
        CCD4_DataArray(26, 1) = NumericUpDown_CCD4_Part8_3.Value
        CCD4_DataArray(27, 1) = NumericUpDown_CCD4_Part9_1.Value
        CCD4_DataArray(28, 1) = NumericUpDown_CCD4_Part9_2.Value
        CCD4_DataArray(29, 1) = NumericUpDown_CCD4_Part9_3.Value
        CCD4_DataArray(30, 1) = NumericUpDown_CCD4_Part10_1.Value
        CCD4_DataArray(31, 1) = NumericUpDown_CCD4_Part10_2.Value
        CCD4_DataArray(32, 1) = NumericUpDown_CCD4_Part10_3.Value
        CCD4_DataArray(33, 1) = NumericUpDown_CCD4_Part11_1.Value
        CCD4_DataArray(34, 1) = NumericUpDown_CCD4_Part11_2.Value
        CCD4_DataArray(35, 1) = NumericUpDown_CCD4_Part11_3.Value
        CCD4_DataArray(36, 1) = NumericUpDown_CCD4_Part12_1.Value
        CCD4_DataArray(37, 1) = NumericUpDown_CCD4_Part12_2.Value
        CCD4_DataArray(38, 1) = NumericUpDown_CCD4_Part12_3.Value
        Save_CCD_Data("CCD4")
        Add_data("CCD4")
    End Sub
    Private Sub Add_data(ByVal Camera_Name As String)
        Try
            Select Case Camera_Name
                Case "CCD1"
                    mTool0.Inputs.Item("Resolution").Value = NumericUpDown_CCD1_Resolution.Value
                    mTool0.Inputs.Item("KZ").Value = NumericUpDown_CCD1_NG_KZ.Value
                    mTool0.Inputs.Item("Width").Value = NumericUpDown_CCD1_Part1_1.Value
                    mTool0.Inputs.Item("Up_1").Value = NumericUpDown_CCD1_Part1_2.Value
                    mTool0.Inputs.Item("Down_1").Value = NumericUpDown_CCD1_Part1_3.Value
                    mTool0.Inputs.Item("Up_2").Value = NumericUpDown_CCD1_Part2_2.Value
                    mTool0.Inputs.Item("Down_2").Value = NumericUpDown_CCD1_Part2_3.Value
                    mTool0.Inputs.Item("Up_3").Value = NumericUpDown_CCD1_Part3_2.Value
                    mTool0.Inputs.Item("Down_3").Value = NumericUpDown_CCD1_Part3_3.Value
                    mTool0.Inputs.Item("Up_4").Value = NumericUpDown_CCD1_Part4_2.Value
                    mTool0.Inputs.Item("Down_4").Value = NumericUpDown_CCD1_Part4_3.Value
                    mTool0.Inputs.Item("Up_5").Value = NumericUpDown_CCD1_Part5_2.Value
                    mTool0.Inputs.Item("Down_5").Value = NumericUpDown_CCD1_Part5_3.Value
                    mTool0.Inputs.Item("Up_6").Value = NumericUpDown_CCD1_Part6_2.Value
                    mTool0.Inputs.Item("Down_6").Value = NumericUpDown_CCD1_Part6_3.Value
                    mTool0.Inputs.Item("Up_7").Value = NumericUpDown_CCD1_Part7_2.Value
                    mTool0.Inputs.Item("Down_7").Value = NumericUpDown_CCD1_Part7_3.Value
                    mTool0.Inputs.Item("Up_8").Value = NumericUpDown_CCD1_Part8_2.Value
                    mTool0.Inputs.Item("Down_8").Value = NumericUpDown_CCD1_Part8_3.Value
                    mTool0.Inputs.Item("Up_9").Value = NumericUpDown_CCD1_Part9_2.Value
                    mTool0.Inputs.Item("Down_9").Value = NumericUpDown_CCD1_Part9_3.Value
                    mTool0.Inputs.Item("Up_10").Value = NumericUpDown_CCD1_Part10_2.Value
                    mTool0.Inputs.Item("Down_10").Value = NumericUpDown_CCD1_Part10_3.Value
                    mTool0.Inputs.Item("Up_11").Value = NumericUpDown_CCD1_Part11_2.Value
                    mTool0.Inputs.Item("Down_11").Value = NumericUpDown_CCD1_Part11_3.Value
                    mTool0.Inputs.Item("Up_12").Value = NumericUpDown_CCD1_Part12_2.Value
                    mTool0.Inputs.Item("Down_12").Value = NumericUpDown_CCD1_Part12_3.Value
                Case "CCD2"
                    mTool1.Inputs.Item("Resolution").Value = NumericUpDown_CCD2_Resolution.Value
                    mTool1.Inputs.Item("KZ").Value = NumericUpDown_CCD3_NG_KZ.Value
                    mTool1.Inputs.Item("Width").Value = NumericUpDown_CCD2_Part1_1.Value
                    mTool1.Inputs.Item("Up_1").Value = NumericUpDown_CCD2_Part1_2.Value
                    mTool1.Inputs.Item("Down_1").Value = NumericUpDown_CCD2_Part1_3.Value
                    mTool1.Inputs.Item("Up_2").Value = NumericUpDown_CCD2_Part2_2.Value
                    mTool1.Inputs.Item("Down_2").Value = NumericUpDown_CCD2_Part2_3.Value
                    mTool1.Inputs.Item("Up_3").Value = NumericUpDown_CCD2_Part3_2.Value
                    mTool1.Inputs.Item("Down_3").Value = NumericUpDown_CCD2_Part3_3.Value
                    mTool1.Inputs.Item("Up_4").Value = NumericUpDown_CCD2_Part4_2.Value
                    mTool1.Inputs.Item("Down_4").Value = NumericUpDown_CCD2_Part4_3.Value
                    mTool1.Inputs.Item("Up_5").Value = NumericUpDown_CCD2_Part5_2.Value
                    mTool1.Inputs.Item("Down_5").Value = NumericUpDown_CCD2_Part5_3.Value
                    mTool1.Inputs.Item("Up_6").Value = NumericUpDown_CCD2_Part6_2.Value
                    mTool1.Inputs.Item("Down_6").Value = NumericUpDown_CCD2_Part6_3.Value
                    mTool1.Inputs.Item("Up_7").Value = NumericUpDown_CCD2_Part7_2.Value
                    mTool1.Inputs.Item("Down_7").Value = NumericUpDown_CCD2_Part7_3.Value
                    mTool1.Inputs.Item("Up_8").Value = NumericUpDown_CCD2_Part8_2.Value
                    mTool1.Inputs.Item("Down_8").Value = NumericUpDown_CCD2_Part8_3.Value
                    mTool1.Inputs.Item("Up_9").Value = NumericUpDown_CCD2_Part9_2.Value
                    mTool1.Inputs.Item("Down_9").Value = NumericUpDown_CCD2_Part9_3.Value
                    mTool1.Inputs.Item("Up_10").Value = NumericUpDown_CCD2_Part10_2.Value
                    mTool1.Inputs.Item("Down_10").Value = NumericUpDown_CCD2_Part10_3.Value
                    mTool1.Inputs.Item("Up_11").Value = NumericUpDown_CCD2_Part11_2.Value
                    mTool1.Inputs.Item("Down_11").Value = NumericUpDown_CCD2_Part11_3.Value
                    mTool1.Inputs.Item("Up_12").Value = NumericUpDown_CCD2_Part12_2.Value
                    mTool1.Inputs.Item("Down_12").Value = NumericUpDown_CCD2_Part12_3.Value
                Case "CCD3"
                    mTool2.Inputs.Item("Resolution").Value = NumericUpDown_CCD3_Resolution.Value
                    mTool2.Inputs.Item("KZ").Value = NumericUpDown_CCD3_NG_KZ.Value
                    mTool2.Inputs.Item("Width").Value = NumericUpDown_CCD3_Part1_1.Value
                    mTool2.Inputs.Item("Up_1").Value = NumericUpDown_CCD3_Part1_2.Value
                    mTool2.Inputs.Item("Down_1").Value = NumericUpDown_CCD3_Part1_3.Value
                    mTool2.Inputs.Item("Up_2").Value = NumericUpDown_CCD3_Part2_2.Value
                    mTool2.Inputs.Item("Down_2").Value = NumericUpDown_CCD3_Part2_3.Value
                    mTool2.Inputs.Item("Up_3").Value = NumericUpDown_CCD3_Part3_2.Value
                    mTool2.Inputs.Item("Down_3").Value = NumericUpDown_CCD3_Part3_3.Value
                    mTool2.Inputs.Item("Up_4").Value = NumericUpDown_CCD3_Part4_2.Value
                    mTool2.Inputs.Item("Down_4").Value = NumericUpDown_CCD3_Part4_3.Value
                    mTool2.Inputs.Item("Up_5").Value = NumericUpDown_CCD3_Part5_2.Value
                    mTool2.Inputs.Item("Down_5").Value = NumericUpDown_CCD3_Part5_3.Value
                    mTool2.Inputs.Item("Up_6").Value = NumericUpDown_CCD3_Part6_2.Value
                    mTool2.Inputs.Item("Down_6").Value = NumericUpDown_CCD3_Part6_3.Value
                    mTool2.Inputs.Item("Up_7").Value = NumericUpDown_CCD3_Part7_2.Value
                    mTool2.Inputs.Item("Down_7").Value = NumericUpDown_CCD3_Part7_3.Value
                    mTool2.Inputs.Item("Up_8").Value = NumericUpDown_CCD3_Part8_2.Value
                    mTool2.Inputs.Item("Down_8").Value = NumericUpDown_CCD3_Part8_3.Value
                    mTool2.Inputs.Item("Up_9").Value = NumericUpDown_CCD3_Part9_2.Value
                    mTool2.Inputs.Item("Down_9").Value = NumericUpDown_CCD3_Part9_3.Value
                    mTool2.Inputs.Item("Up_10").Value = NumericUpDown_CCD3_Part10_2.Value
                    mTool2.Inputs.Item("Down_10").Value = NumericUpDown_CCD3_Part10_3.Value
                    mTool2.Inputs.Item("Up_11").Value = NumericUpDown_CCD3_Part11_2.Value
                    mTool2.Inputs.Item("Down_11").Value = NumericUpDown_CCD3_Part11_3.Value
                    mTool2.Inputs.Item("Up_12").Value = NumericUpDown_CCD3_Part12_2.Value
                    mTool2.Inputs.Item("Down_12").Value = NumericUpDown_CCD3_Part12_3.Value
                Case "CCD4"
                    mTool3.Inputs.Item("Resolution").Value = NumericUpDown_CCD4_Resolution.Value
                    mTool3.Inputs.Item("KZ").Value = NumericUpDown_CCD4_NG_KZ.Value
                    mTool3.Inputs.Item("Width").Value = NumericUpDown_CCD4_Part1_1.Value
                    mTool3.Inputs.Item("Up_1").Value = NumericUpDown_CCD4_Part1_2.Value
                    mTool3.Inputs.Item("Down_1").Value = NumericUpDown_CCD4_Part1_3.Value
                    mTool3.Inputs.Item("Up_2").Value = NumericUpDown_CCD4_Part2_2.Value
                    mTool3.Inputs.Item("Down_2").Value = NumericUpDown_CCD4_Part2_3.Value
                    mTool3.Inputs.Item("Up_3").Value = NumericUpDown_CCD4_Part3_2.Value
                    mTool3.Inputs.Item("Down_3").Value = NumericUpDown_CCD4_Part3_3.Value
                    mTool3.Inputs.Item("Up_4").Value = NumericUpDown_CCD4_Part4_2.Value
                    mTool3.Inputs.Item("Down_4").Value = NumericUpDown_CCD4_Part4_3.Value
                    mTool3.Inputs.Item("Up_5").Value = NumericUpDown_CCD4_Part5_2.Value
                    mTool3.Inputs.Item("Down_5").Value = NumericUpDown_CCD4_Part5_3.Value
                    mTool3.Inputs.Item("Up_6").Value = NumericUpDown_CCD4_Part6_2.Value
                    mTool3.Inputs.Item("Down_6").Value = NumericUpDown_CCD4_Part6_3.Value
                    mTool3.Inputs.Item("Up_7").Value = NumericUpDown_CCD4_Part7_2.Value
                    mTool3.Inputs.Item("Down_7").Value = NumericUpDown_CCD4_Part7_3.Value
                    mTool3.Inputs.Item("Up_8").Value = NumericUpDown_CCD4_Part8_2.Value
                    mTool3.Inputs.Item("Down_8").Value = NumericUpDown_CCD4_Part8_3.Value
                    mTool3.Inputs.Item("Up_9").Value = NumericUpDown_CCD4_Part9_2.Value
                    mTool3.Inputs.Item("Down_9").Value = NumericUpDown_CCD4_Part9_3.Value
                    mTool3.Inputs.Item("Up_10").Value = NumericUpDown_CCD4_Part10_2.Value
                    mTool3.Inputs.Item("Down_10").Value = NumericUpDown_CCD4_Part10_3.Value
                    mTool3.Inputs.Item("Up_11").Value = NumericUpDown_CCD4_Part11_2.Value
                    mTool3.Inputs.Item("Down_11").Value = NumericUpDown_CCD4_Part11_3.Value
                    mTool3.Inputs.Item("Up_12").Value = NumericUpDown_CCD4_Part12_2.Value
                    mTool3.Inputs.Item("Down_12").Value = NumericUpDown_CCD4_Part12_3.Value
            End Select
        Catch ex As Exception

        End Try
    End Sub
#End Region
#Region "手动运行"
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            myJob0.Run()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            myJob1.Run()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            myJob2.Run()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            myJob3.Run()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Try
            myJob0.Run()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Try
            myJob1.Run()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Try
            myJob2.Run()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Try
            myJob3.Run()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        myJobManager.Run()
    End Sub
#End Region
#Region "OPC"
    Dim m_server As IOPCServer = Nothing
    Dim m_group As IOPCGroupStateMgt = Nothing
    Dim m_item As IOPCItemMgt = Nothing
    Dim m_asyncIO2 As IOPCAsyncIO2 = Nothing

    Dim m_ConnectionContainer As IConnectionPointContainer = Nothing
    Dim m_ConnectionPoint As IConnectionPoint = Nothing
    Dim m_Callback As IOPCDataCallback = Nothing
    Dim m_dwCookie As Int32 = 0

    'Used by GetErrorString; Language code &H409 = ENGLISH
    Const LOCALE_ID As Integer = &H409

    Dim Transaction As Integer = 0
    Dim ServerGroup As Integer = 0
    Dim ClientGrouphandle As Integer = 0
    Dim ServerHandle1 As Integer = 0
    Dim ServerHandle2 As Integer = 0
    Dim activeGroup As Boolean = False
    Dim OutText As String
    Private Sub Command_Start_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command_Start.Click
        Start_Sample()
    End Sub
    Private Sub Command_Read_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command_Read.Click
        Read_value(6)
        Read_value(13)
    End Sub
    Private Sub Command_Write_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command_Write.Click
        Write_value(Edit_WriteVal.Text, ComboBox2.Text)
    End Sub
    Private Sub Command_Exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command_Exit.Click
        Stop_Sample()
    End Sub
    Private Sub chkGroupActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGroupActive.CheckedChanged
        Dim RequestedUpdateRate As Integer = 250
        Dim Active As Integer = 0
        Dim Loc As Integer = &H409

        Dim pTimeBias As IntPtr = IntPtr.Zero
        Dim pDeadband As IntPtr = IntPtr.Zero
        'Dim phServerGroup As IntPtr = IntPtr.Zero

        'Dim hServerGroup As Integer
        Dim RevisedUpdateRate As Integer

        ' Access unmanaged memory
        Dim hRequestedUpdateRate As GCHandle = GCHandle.Alloc(RequestedUpdateRate, GCHandleType.Pinned)
        Dim hLoc As GCHandle = GCHandle.Alloc(Loc, GCHandleType.Pinned)
        Dim hActive As GCHandle = GCHandle.Alloc(Active, GCHandleType.Pinned)
        Dim hClientGroup As GCHandle = GCHandle.Alloc(ClientGrouphandle, GCHandleType.Pinned)

        Dim OutText As String


        Try
            OutText = "Group active checked"

            If Not activeGroup Then
                ' Activate group
                activeGroup = True
                hActive.Target = 1
                m_group.SetState(hRequestedUpdateRate.AddrOfPinnedObject(), _
                    RevisedUpdateRate, _
                    hActive.AddrOfPinnedObject(), _
                    pTimeBias, _
                    pDeadband, _
                    hLoc.AddrOfPinnedObject(), _
                    hClientGroup.AddrOfPinnedObject())
            Else
                ' Deactivate group
                activeGroup = False
                hActive.Target = 0
                m_group.SetState(hRequestedUpdateRate.AddrOfPinnedObject(), _
                    RevisedUpdateRate, _
                    hActive.AddrOfPinnedObject(), _
                    pTimeBias, _
                    pDeadband, _
                    hLoc.AddrOfPinnedObject(), _
                    hClientGroup.AddrOfPinnedObject())
                ' Clear previously shown values
                Edit_OnDataVal_0.Text = ""
                Edit_OnDataQu_0.Text = ""
                Edit_OnDataTS_0.Text = ""
                Edit_OnDataVal_1.Text = ""
                Edit_OnDataQu_1.Text = ""
                Edit_OnDataTS_1.Text = ""
                Edit_OnDataVal_2.Text = ""
                Edit_OnDataQu_2.Text = ""
                Edit_OnDataTS_2.Text = ""
                Edit_OnDataVal_3.Text = ""
                Edit_OnDataQu_3.Text = ""
                Edit_OnDataTS_3.Text = ""
                Edit_OnDataVal_4.Text = ""
                Edit_OnDataQu_4.Text = ""
                Edit_OnDataTS_4.Text = ""
                Edit_OnDataVal_5.Text = ""
                Edit_OnDataQu_5.Text = ""
                Edit_OnDataTS_5.Text = ""
                Edit_OnDataVal_6.Text = ""
                Edit_OnDataQu_6.Text = ""
                Edit_OnDataTS_6.Text = ""
                Edit_OnDataVal_7.Text = ""
                Edit_OnDataQu_7.Text = ""
                Edit_OnDataTS_7.Text = ""
            End If

        Catch ex As Exception
            OutText = ""
            MessageBox.Show(OutText & Chr(13) & Chr(13) & ex.ToString(), "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Free unmanaged memory
            If hActive.IsAllocated Then hActive.Free()
            If hLoc.IsAllocated Then hLoc.Free()
            If hRequestedUpdateRate.IsAllocated Then hRequestedUpdateRate.Free()
            If hClientGroup.IsAllocated Then hClientGroup.Free()
        End Try
    End Sub
    Private Sub Edit_WriteVal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Edit_WriteVal.TextChanged
        ' Clear old value
        Command_Write.Enabled = True
        Edit_WriteRes.Text = ""
    End Sub
    Private Sub Start_Sample()
        Try
            OutText = "Create instance of the SIMATIC NET OPC-Server"

            ' Given a ProgID, this looks up the associated Type in the registry
            Dim typeofOPCserver As Type = Type.GetTypeFromProgID("OPC.SimaticNET")

            ' Create the OPC server object and querry for the IOPCServer interface of the object
            m_server = Activator.CreateInstance(typeofOPCserver)

            Edit_WriteVal.ReadOnly = False
            OutText = "Adding a Group to OPC-Server"

            Dim RevisedUpdateRate As Integer
            Dim TimeBias As Int32 = 0
            Dim Deadband As Double = 0
            Dim hTimeBias As GCHandle = GCHandle.Alloc(TimeBias, GCHandleType.Pinned)
            Dim hDeadband As GCHandle = GCHandle.Alloc(TimeBias, GCHandleType.Pinned)

            ClientGrouphandle = 10
            ' Add a group object "MyOPCGroup" and querry for interface IOPCItemMgt
            ' Parameter as following:
            ' [in] not active, so no OnDataChange callback
            ' [in] Request this Update Rate from Server
            ' [in] Client Handle, not necessary in this sample
            ' [in] No time interval to system UTC time
            ' [in] No Deadband, so all data changes are reported
            ' [in] Server uses english language to for text values
            ' [out] Server handle to identify this group in later calls
            ' [out] The answer from Server to the requested Update Rate
            ' [in] requested interface type of the group object
            ' [out] pointer to the requested interface
            m_server.AddGroup( _
                    "MyOPCGroup", _
                    0, _
                    250, _
                    ClientGrouphandle, _
                    hTimeBias.AddrOfPinnedObject(), _
                    hDeadband.AddrOfPinnedObject(), _
                    LOCALE_ID, _
                    ServerGroup, _
                    RevisedUpdateRate, _
                    GetType(IOPCGroupStateMgt).GUID, _
                    m_group)
            m_item = m_group

            ' define item tables with 2 items as in parameters for AddItem
            Dim itemdefs(15) As OPCITEMDEF


            itemdefs(0) = New OPCITEMDEF
            itemdefs(0).hClient = 1
            itemdefs(0).szItemID = "S7:[S7 connection_1]DB1100,0.1"     ''''''''数据清零
            itemdefs(0).szAccessPath = ""
            itemdefs(0).bActive = 1
            itemdefs(0).vtRequestedDataType = VarEnum.VT_BOOL
            itemdefs(0).dwBlobSize = 0
            itemdefs(0).pBlob = IntPtr.Zero

            itemdefs(1) = New OPCITEMDEF
            itemdefs(1).hClient = 2
            itemdefs(1).szItemID = "S7:[S7 connection_1]DB1100,0.2"    ''''''''' 发送条码
            itemdefs(1).szAccessPath = ""
            itemdefs(1).bActive = 1
            itemdefs(1).vtRequestedDataType = VarEnum.VT_BOOL
            itemdefs(1).dwBlobSize = 0
            itemdefs(1).pBlob = IntPtr.Zero

            itemdefs(2) = New OPCITEMDEF
            itemdefs(2).hClient = 3
            itemdefs(2).szItemID = "S7:[S7 connection_1]DB1100,0.3"    '''''''''拍照信号
            itemdefs(2).szAccessPath = ""
            itemdefs(2).bActive = 1
            itemdefs(2).vtRequestedDataType = VarEnum.VT_BOOL
            itemdefs(2).dwBlobSize = 0
            itemdefs(2).pBlob = IntPtr.Zero

            itemdefs(3) = New OPCITEMDEF
            itemdefs(3).hClient = 4
            itemdefs(3).szItemID = "S7:[S7 connection_1]DB1100,10"        ''''''''产品型号
            itemdefs(3).szAccessPath = ""
            itemdefs(3).bActive = 1
            itemdefs(3).vtRequestedDataType = VarEnum.VT_UI2
            itemdefs(3).dwBlobSize = 0
            itemdefs(3).pBlob = IntPtr.Zero

            itemdefs(4) = New OPCITEMDEF
            itemdefs(4).hClient = 5
            itemdefs(4).szItemID = "S7:[S7 connection_1]DB1100,STRING12.38"   ''''''''''''''工件条码
            itemdefs(4).szAccessPath = ""
            itemdefs(4).bActive = 1
            itemdefs(4).vtRequestedDataType = VarEnum.VT_BSTR
            itemdefs(4).dwBlobSize = 0
            itemdefs(4).pBlob = IntPtr.Zero

            itemdefs(5) = New OPCITEMDEF
            itemdefs(5).hClient = 6
            itemdefs(5).szItemID = "S7:[S7 connection_1]DB1100,200.0"     ''''''''心跳
            itemdefs(5).szAccessPath = ""
            itemdefs(5).bActive = 1
            itemdefs(5).vtRequestedDataType = VarEnum.VT_BOOL
            itemdefs(5).dwBlobSize = 0
            itemdefs(5).pBlob = IntPtr.Zero

            itemdefs(6) = New OPCITEMDEF
            itemdefs(6).hClient = 7
            itemdefs(6).szItemID = "S7:[S7 connection_1]DB1100,200.1"  ''''''''''''准备好信号
            itemdefs(6).szAccessPath = ""
            itemdefs(6).bActive = 1
            itemdefs(6).vtRequestedDataType = VarEnum.VT_BOOL
            itemdefs(6).dwBlobSize = 0
            itemdefs(6).pBlob = IntPtr.Zero

            itemdefs(7) = New OPCITEMDEF
            itemdefs(7).hClient = 8
            itemdefs(7).szItemID = "S7:[S7 connection_1]DB1100,200.2"  '''''''''清零成功
            itemdefs(7).szAccessPath = ""
            itemdefs(7).bActive = 1
            itemdefs(7).vtRequestedDataType = VarEnum.VT_BOOL
            itemdefs(7).dwBlobSize = 0
            itemdefs(7).pBlob = IntPtr.Zero

            itemdefs(8) = New OPCITEMDEF
            itemdefs(8).hClient = 9
            itemdefs(8).szItemID = "S7:[S7 connection_1]DB1100,200.3"  '''''''''条码成功
            itemdefs(8).szAccessPath = ""
            itemdefs(8).bActive = 1
            itemdefs(8).vtRequestedDataType = VarEnum.VT_BOOL
            itemdefs(8).dwBlobSize = 0
            itemdefs(8).pBlob = IntPtr.Zero

            itemdefs(9) = New OPCITEMDEF
            itemdefs(9).hClient = 10
            itemdefs(9).szItemID = "S7:[S7 connection_1]DB1100,200.4"  '''''''''拍照OK
            itemdefs(9).szAccessPath = ""
            itemdefs(9).bActive = 1
            itemdefs(9).vtRequestedDataType = VarEnum.VT_BOOL
            itemdefs(9).dwBlobSize = 0
            itemdefs(9).pBlob = IntPtr.Zero

            itemdefs(10) = New OPCITEMDEF
            itemdefs(10).hClient = 11
            itemdefs(10).szItemID = "S7:[S7 connection_1]DB1100,200.5"  '''''''''拍照NG
            itemdefs(10).szAccessPath = ""
            itemdefs(10).bActive = 1
            itemdefs(10).vtRequestedDataType = VarEnum.VT_BOOL
            itemdefs(10).dwBlobSize = 0
            itemdefs(10).pBlob = IntPtr.Zero

            itemdefs(11) = New OPCITEMDEF
            itemdefs(11).hClient = 12
            itemdefs(11).szItemID = "S7:[S7 connection_1]DB1100,STRING52.38"  '''''''''拍照NG
            itemdefs(11).szAccessPath = ""
            itemdefs(11).bActive = 1
            itemdefs(11).vtRequestedDataType = VarEnum.VT_BSTR
            itemdefs(11).dwBlobSize = 0
            itemdefs(11).pBlob = IntPtr.Zero

            itemdefs(12) = New OPCITEMDEF
            itemdefs(12).hClient = 13
            itemdefs(12).szItemID = "S7:[S7 connection_1]DB1100,0.4"  '''''''''轴承拍照
            itemdefs(12).szAccessPath = ""
            itemdefs(12).bActive = 1
            itemdefs(12).vtRequestedDataType = VarEnum.VT_BOOL
            itemdefs(12).dwBlobSize = 0
            itemdefs(12).pBlob = IntPtr.Zero

            itemdefs(13) = New OPCITEMDEF
            itemdefs(13).hClient = 14
            itemdefs(13).szItemID = "S7:[S7 connection_1]DB1100,200.6"  '''''''''拍照OK
            itemdefs(13).szAccessPath = ""
            itemdefs(13).bActive = 1
            itemdefs(13).vtRequestedDataType = VarEnum.VT_BOOL
            itemdefs(13).dwBlobSize = 0
            itemdefs(13).pBlob = IntPtr.Zero

            itemdefs(14) = New OPCITEMDEF
            itemdefs(14).hClient = 15
            itemdefs(14).szItemID = "S7:[S7 connection_1]DB1100,200.7"  '''''''''拍照NG
            itemdefs(14).szAccessPath = ""
            itemdefs(14).bActive = 1
            itemdefs(14).vtRequestedDataType = VarEnum.VT_BOOL
            itemdefs(14).dwBlobSize = 0
            itemdefs(14).pBlob = IntPtr.Zero

            OutText = "Adding Items to the Group"

            Dim Results As IntPtr = IntPtr.Zero
            Dim pErrors As IntPtr = IntPtr.Zero

            ' Adding two Items to the Group
            m_item.AddItems(15, _
                itemdefs, _
                Results, _
                pErrors)


            'Evaluate return ErrorCodes to exclude possible Errors
            Dim errors(15) As Integer
            Marshal.Copy(pErrors, errors, 0, 15)
            Dim result As OPCITEMRESULT
            Dim pos As IntPtr = Results
  
            For i = 0 To 14
                If errors(i) = 0 Then
                    ' First Item was added succesfully
                    If i > 0 Then
                        pos = New IntPtr(pos.ToInt32() + Marshal.SizeOf(GetType(OPCITEMRESULT)))
                    End If
                    result = Marshal.PtrToStructure(pos, GetType(OPCITEMRESULT))
                    ServerHandle1 = result.hServer
                Else
                    ' First Item was not added
                    ' Raise Exception
                    Dim ex As Exception = New Exception("Error while adding " & (i + 1) & "th Item")
                    Throw ex
                End If
                ' Destroy indirect structure elements
                Marshal.DestroyStructure(pos, GetType(OPCITEMRESULT))
            Next
            ' Free allocated COM-ressouces
            Marshal.FreeCoTaskMem(Results)
            Marshal.FreeCoTaskMem(pErrors)

         

            OutText = "Adding asynchron Interface"


            m_asyncIO2 = m_group


            activeGroup = False

            m_ConnectionContainer = m_group

            OutText = "Establishing asynchronous callbacks"
            m_ConnectionContainer.FindConnectionPoint(GetType(IOPCDataCallback).GUID, m_ConnectionPoint)
            m_ConnectionPoint.Advise(Me, m_dwCookie)
            Command_Start.Enabled = False
            Command_Read.Enabled = True
            Command_Write.Enabled = True
            Command_Exit.Enabled = True
            chkGroupActive.Enabled = True
            'chkGroupActive.Checked = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, OutText, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub Stop_Sample()
        Try
            chkGroupActive.Checked = False
            OutText = "Removing Objects"

            ' Unadvise and remove the callback
            m_ConnectionPoint.Unadvise(m_dwCookie)
            Marshal.ReleaseComObject(m_ConnectionPoint)
            m_ConnectionPoint = Nothing

            ' Removes items
            Dim pItems(2) As Integer
            ' Select item by server handle, received at AddItem
            pItems(0) = ServerHandle1
            pItems(1) = ServerHandle2

            Dim pErrors As IntPtr = IntPtr.Zero

            m_item.RemoveItems(2, _
                        pItems, _
                        pErrors)

            'Evaluating Return ErrorCodes to exclude possible Errors
            Dim errors(2) As Integer
            Marshal.Copy(pErrors, errors, 0, 2)
            If errors(0) = 0 Then
                ' First Item was deleted succesfully
                ServerHandle1 = 0
            Else
                ' First Item was not added
                ' Raise Exception
                Dim ex As Exception = New Exception("Error while deleting first Item")
                Throw ex
            End If

            ' Free allocated COM-ressouces
            Marshal.FreeCoTaskMem(pErrors)

            ' *****************************************************
            ' removes Group
            ' *****************************************************
            m_server.RemoveGroup(ServerGroup, True)

            ' *****************************************************
            ' Release all interfaces
            ' *****************************************************
            m_asyncIO2 = Nothing
            m_item = Nothing
            m_group = Nothing
            m_server = Nothing
            ' The Server is terminated now

            ' Reset dialog
            Edit_ReadVal.Text = ""
            Edit_ReadQu.Text = ""
            Edit_ReadTS.Text = ""
            Edit_OnDataVal_0.Text = ""
            Edit_OnDataQu_0.Text = ""
            Edit_OnDataTS_0.Text = ""
            Edit_OnDataVal_1.Text = ""
            Edit_OnDataQu_1.Text = ""
            Edit_OnDataTS_1.Text = ""
            Edit_WriteRes.Text = ""
            Edit_WriteVal.Text = ""

            Command_Start.Enabled = True
            Command_Read.Enabled = False
            Command_Write.Enabled = False
            Command_Exit.Enabled = False
            chkGroupActive.Enabled = False
            Edit_WriteVal.ReadOnly = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, OutText, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Exit Sub
    End Sub
    Private Sub Read_value(ByVal i As Integer)
        'Timer1.Stop()
        Dim ErrorString As String
        Try
            OutText = "Reading Value of Item"
            ErrorString = ""
            Dim CancelID As Integer
            Dim pServer(1) As Integer
            ' Select item by server handle, received at AddItem
            pServer(0) = i

            Dim pItemValues As IntPtr = IntPtr.Zero, pErrors As IntPtr = IntPtr.Zero
            Transaction += 1
            m_asyncIO2.Read(1, _
                pServer, _
                Transaction, _
                CancelID, _
                pErrors)
            'check if an error occured
            Dim errors(1) As Integer
            Marshal.Copy(pErrors, errors, 0, 1)
            If errors(0) <> 0 Then
                ' Errors occured - raise Exception
                Dim ex As Exception = New Exception("Read Error")
                Throw ex
            End If

            ' Free allocated COM-ressouces
            Marshal.FreeCoTaskMem(pErrors)

        Catch ex As Exception
            ' Timer1.Enabled = False
            If MessageBox.Show("OPC通讯异常，请重启软件？", "提示信息", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = Windows.Forms.DialogResult.OK Then

                Me.Close()
            Else
                Timer1.Enabled = True
            End If
            MessageBox.Show(ex.Message, OutText, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Exit Sub
    End Sub
    Private Sub Write_value(ByVal j As String, ByVal K As String)
        Dim ErrorString As String
        OutText = "Writing Value"
        ErrorString = ""
        Try
            Dim phServer(1) As Integer
            phServer(0) = Val(K)

            Dim phItemValues(14) As Object
            phItemValues(0) = j
            Dim dwCancelID As Integer
            Dim pErrors As IntPtr = IntPtr.Zero
            Transaction += 1
            ' Select item by server handle, received at AddItem
            m_asyncIO2.Write(15, _
                 phServer, _
                 phItemValues, _
                 Transaction, _
                 dwCancelID, _
                 pErrors)
            'check if an error occured
            Dim errors(15) As Integer
            Marshal.Copy(pErrors, errors, 0, 15)
            If errors(0) <> 0 Then
                ' Errors occured - raise Exception
                Edit_WriteRes.Text = "Bad"
                Dim ex As Exception = New Exception("Read Error")
                Throw ex
            End If

            ' Free allocated COM-ressouces
            Marshal.FreeCoTaskMem(pErrors)

        Catch ex As Exception
            MessageBox.Show(ex.Message, OutText, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Public Function ToDateTime(ByVal ft As OpcRcw.Da.FILETIME) As DateTime
        'Converts FILETIME to DateTime
        Dim High As Long = ft.dwHighDateTime
        Dim lng As Long = (High << 32) + ft.dwLowDateTime
        Return DateTime.FromFileTimeUtc(lng)
    End Function

    Private Sub OPCAsync_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    ' delegate is used by the UI update mechanism (see below)
    Private Delegate Sub UpdateTextboxDelegate(ByVal ctrlTextbox As TextBox, ByVal strText As String)

    Private Sub UpdateTextbox(ByVal ctrlTextbox As TextBox, ByVal strText As String)
        ' can update a text box from an arbitrary thread
        If InvokeRequired Then 'thread switch required
            Dim handler As New UpdateTextboxDelegate(AddressOf DoUpdateTextbox)
            Dim listArguments() As Object = {ctrlTextbox, strText}
            BeginInvoke(handler, listArguments)
        Else ' called from within user interface thread
            DoUpdateTextbox(ctrlTextbox, strText)
        End If
    End Sub

    ' can update a text box only from the user interface thread itself
    Private Sub DoUpdateTextbox(ByVal ctrlTextbox As TextBox, ByVal strText As String)
        ctrlTextbox.Text = strText
    End Sub

    ' sends data update notifications.
    Overridable Sub OnDataChange( _
    ByVal dwTransid As Integer, _
    ByVal hGroup As Integer, _
    ByVal hrMasterquality As Integer, _
    ByVal hrMastererror As Integer, _
    ByVal dwCount As Integer, _
    ByVal phClientItems() As Integer, _
    ByVal pvValues() As Object, _
    ByVal pwQualities() As Short, _
    ByVal pftTimeStamps() As OpcRcw.Da.FILETIME, _
    ByVal pErrors() As Integer) Implements IOPCDataCallback.OnDataChange
        Try
            If dwCount >= 1 Then
                If pErrors(0) = 0 Then
                    Dim dt As DateTime = ToDateTime(pftTimeStamps(0))
                    Select Case phClientItems(0)
                        Case 1  '''''''''''''''数据清零信号

                            If pvValues(0) = True Then
                                Write_value(False, 10)
                                Write_value(False, 11)
                                Write_value(False, 14)
                                Write_value(False, 15)
                                Write_value(True, 8)
                            End If
                        Case 2 '''''''''''''''''''发送条码
                            If pvValues(0) = True Then
                                Read_value(5)
                                Read_value(12)
                              
                                Write_value(True, 9)

                            End If

                        Case 3  ''''''''''''''''''''拍照信号
                            If pvValues(0) = True Then
                                Dim fname As String
                                fname = "\\10.230.211.8\image\" + ComboBox1.Text + Format(Now, "yyyy年M月d日") + "\" + Label173.Text
                                If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
                                    MkDir(fname)                     '创建文件夹         
                                End If
                                CCD1_Result.Text = ""
                                CCD2_Result.Text = ""
                                CCD3_Result.Text = ""
                                CCD4_Result.Text = ""
                                CogRecordDisplay1.Image = Nothing
                                CogRecordDisplay2.Image = Nothing
                                CogRecordDisplay3.Image = Nothing
                                CogRecordDisplay4.Image = Nothing
                                If RunState = True Then
                                    RunState = False
                                    myJob0.Run()
                                    myJob1.Run()
                                    myJob2.Run()
                                    myJob3.Run()
                                    Thread.Sleep(1500)
                                    If CCD2_Result.Text = "OK" And CCD1_Result.Text = "OK" And CCD3_Result.Text = "OK" And CCD4_Result.Text = "OK" Then
                                        Write_value(False, 11)
                                        Write_value(True, 10)
                                    Else

                                        Write_value(True, 11)
                                        Write_value(False, 10)
                                    End If
                                    RunState = True
                                End If
                            End If

                        Case 4 '''''''工件信号
                            If pvValues(0) = 1 Then
                                VPP_add("DT530")
                            End If
                            If pvValues(0) = 2 Then
                                VPP_add("DT350")
                            End If
                            If pvValues(0) = 11 Then
                                VPP_add("DT200")
                            End If
                        Case 13
                            If pvValues(0) = True Then
                                Label175.Text = ""
                                myJob4.Run()



                            End If
                    End Select

                End If
            End If
        Catch e As Exception
            Dim msg As String = e.Message
            MsgBox(msg, MsgBoxStyle.Critical)
        End Try
    End Sub
    'sends read complete notifications.
    Overridable Sub OnReadComplete( _
    ByVal dwTransid As Integer, _
    ByVal hGroup As Integer, _
    ByVal hrMasterquality As Integer, _
    ByVal hrMastererror As Integer, _
    ByVal dwCount As Integer, _
    ByVal phClientItems() As Integer, _
    ByVal pvValues() As Object, _
    ByVal pwQualities() As Short, _
    ByVal pftTimeStamps() As OpcRcw.Da.FILETIME, _
    ByVal pErrors() As Integer) Implements OpcRcw.Da.IOPCDataCallback.OnReadComplete
        Try
            If pErrors(0) = 0 Then

                If phClientItems(0) = 5 Then
                    If Not Convert.ToString(pvValues(0)) = "" Then
                        SN = Convert.ToString(pvValues(0))
                        Label173.Text = SN
                        ' Write_value(True, 10)
                    End If

                End If
                If phClientItems(0) = 12 Then
                    If pErrors(0) = 0 Then
                        If Not Convert.ToString(pvValues(0)) = "" Then
                            ID = Convert.ToString(pvValues(0))
                            Label174.Text = ID
                        End If
                    End If
                End If
            Else
                Dim ErrorString As String = "Read Error"
                MsgBox(ErrorString, MsgBoxStyle.Critical, "Error AsyncReadComplete()")
            End If


        Catch e As Exception
            Dim msg As String = e.Message
            MsgBox(msg, MsgBoxStyle.Critical)
        End Try
    End Sub
    ' handles asynchronous write complete events.
    Overridable Sub OnWriteComplete( _
    ByVal dwTransid As Integer, _
    ByVal hGroup As Integer, _
    ByVal hrMastererror As Integer, _
    ByVal dwCount As Integer, _
    ByVal phClientItems() As Integer, _
    ByVal pErrors() As Integer) Implements IOPCDataCallback.OnWriteComplete
        Try
            Dim ErrorStr As String = ""

            If dwCount = 1 Then
                m_server.GetErrorString(pErrors(0), LOCALE_ID, ErrorStr)
                UpdateTextbox(Edit_WriteRes, ErrorStr)
            End If
        Catch e As Exception
            Dim msg As String = e.Message
            MsgBox(msg, MsgBoxStyle.Critical)
        End Try
    End Sub
    ' handles asynchronous request cancel events.
    Public Sub OnCancelComplete( _
            ByVal dwTransid As Integer, _
            ByVal inthGroup As Integer) Implements IOPCDataCallback.OnCancelComplete
        ' This routine has to be present, although not used in this sample
    End Sub

    Private Function GetQualityText(ByVal qnr As Integer) As String
        'Const OPC_QUALITY_MASK As Short = &HC0
        'Const OPC_STATUS_MASK As Short = &HFC
        'Const OPC_LIMIT_MASK As Short = &H3
        Const OPC_QUALITY_BAD As Short = &H0
        Const OPC_QUALITY_UNCERTAIN As Short = &H40
        Const OPC_QUALITY_GOOD As Short = &HC0
        Const OPC_QUALITY_CONFIG_ERROR As Short = &H4
        Const OPC_QUALITY_NOT_CONNECTED As Short = &H8
        Const OPC_QUALITY_DEVICE_FAILURE As Short = &HC
        Const OPC_QUALITY_SENSOR_FAILURE As Short = &H10
        Const OPC_QUALITY_LAST_KNOWN As Short = &H14
        Const OPC_QUALITY_COMM_FAILURE As Short = &H18
        Const OPC_QUALITY_OUT_OF_SERVICE As Short = &H1C
        Const OPC_QUALITY_LAST_USABLE As Short = &H44
        Const OPC_QUALITY_SENSOR_CAL As Short = &H50
        Const OPC_QUALITY_EGU_EXCEEDED As Short = &H54
        Const OPC_QUALITY_SUB_NORMAL As Short = &H58
        Const OPC_QUALITY_LOCAL_OVERRIDE As Short = &HD8
        'Const OPC_LIMIT_OK As Short = &H0
        'Const OPC_LIMIT_LOW As Short = &H1
        'Const OPC_LIMIT_HIGH As Short = &H2
        'Const OPC_LIMIT_CONST As Short = &H3

        Dim qstr As String = ""

        Select Case qnr

            Case OPC_QUALITY_BAD
                qstr = "BAD"
            Case OPC_QUALITY_UNCERTAIN
                qstr = "UNCERTAIN"
            Case OPC_QUALITY_GOOD
                qstr = "GOOD"
            Case OPC_QUALITY_CONFIG_ERROR
                qstr = "CONFIG_ERROR"
            Case OPC_QUALITY_NOT_CONNECTED
                qstr = "NOT_CONNECTED"
            Case OPC_QUALITY_DEVICE_FAILURE
                qstr = "DEVICE_FAILURE"
            Case OPC_QUALITY_SENSOR_FAILURE
                qstr = "SENSOR_FAILURE"
            Case OPC_QUALITY_LAST_KNOWN
                qstr = "LAST_KNOWN"
            Case OPC_QUALITY_COMM_FAILURE
                qstr = "COMM_FAILURE"
            Case OPC_QUALITY_OUT_OF_SERVICE
                qstr = "OUT_OF_SERVICE"
            Case OPC_QUALITY_LAST_USABLE
                qstr = "LAST_USABLE"
            Case OPC_QUALITY_SENSOR_CAL
                qstr = "SENSOR_CAL"
            Case OPC_QUALITY_EGU_EXCEEDED
                qstr = "EGU_EXCEEDED"
            Case OPC_QUALITY_SUB_NORMAL
                qstr = "SUB_NORMAL"
            Case OPC_QUALITY_LOCAL_OVERRIDE
                qstr = "LOCAL_OVERRIDE"
            Case Else
                qstr = "UNKNOWN ERROR"
        End Select
        Return qstr
    End Function
#End Region
#Region "清除计数"
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Count_OK = 0
        Count_NG = 0
        TextBox_OK.Text = Count_OK
        TextBox_NG.Text = Count_NG
        TextBox_Total.Text = Count_OK + Count_NG
    End Sub
#End Region

#Region "保存图片"
    Private lock As New Object
  
#End Region
    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim fname As String
        fname = "G:\OP290\CCD1"
        If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
            MkDir(fname)                     '创建文件夹         
        End If
        Shell("Explorer G:\OP290\CCD1", vbNormalFocus)
    End Sub
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim fname As String
        fname = "G:\OP290\CCD2"
        If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
            MkDir(fname)                     '创建文件夹         
        End If
        Shell("Explorer G:\OP290\CCD2", vbNormalFocus)
    End Sub
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim fname As String
        fname = "G:\OP290\CCD3"
        If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
            MkDir(fname)                     '创建文件夹         
        End If
        Shell("Explorer G:\OP290\CCD3", vbNormalFocus)
    End Sub
    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim fname As String
        fname = "G:\OP290\CCD4"
        If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
            MkDir(fname)                     '创建文件夹         
        End If
        Shell("Explorer G:\OP290\CCD4", vbNormalFocus)
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Write_value(True, 6)
        Write_value(False, 6)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button15_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        myJobManager = CType(CogSerializer.LoadObjectFromFile("VPP\PartNum\" & ComboBox1.Text & ".vpp"), CogJobManager)
        myJob0 = myJobManager.Job(0)
        myIndependentJob0 = myJob0.OwnedIndependent
        myJob1 = myJobManager.Job(1)
        myIndependentJob1 = myJob1.OwnedIndependent
        myJob2 = myJobManager.Job(2)
        myIndependentJob2 = myJob2.OwnedIndependent
        myJob3 = myJobManager.Job(3)
        myIndependentJob3 = myJob3.OwnedIndependent
        myJobManager.UserQueueFlush()
        myJobManager.FailureQueueFlush()
        myJob0.ImageQueueFlush()
        myIndependentJob0.RealTimeQueueFlush()
        myJob1.ImageQueueFlush()
        myIndependentJob1.RealTimeQueueFlush()
        myJob2.ImageQueueFlush()
        myIndependentJob2.RealTimeQueueFlush()
        myJob3.ImageQueueFlush()
        myIndependentJob3.RealTimeQueueFlush()
        myJob4.ImageQueueFlush()
        myIndependentJob4.RealTimeQueueFlush()
        AddHandler myJobManager.Stopped, AddressOf myJobManager_Stopped
        AddHandler myJobManager.UserResultAvailable, AddressOf myJobManager_UserResultAvailable
        mGroup0 = CType(myJob0.VisionTool, CogToolGroup)
        mTool0 = CType(mGroup0.Tools("CogToolBlock1"), CogToolBlock)
        mGroup1 = CType(myJob1.VisionTool, CogToolGroup)
        mTool1 = CType(mGroup1.Tools("CogToolBlock1"), CogToolBlock)
        mGroup2 = CType(myJob2.VisionTool, CogToolGroup)
        mTool2 = CType(mGroup2.Tools("CogToolBlock1"), CogToolBlock)
        mGroup3 = CType(myJob3.VisionTool, CogToolGroup)
        mTool3 = CType(mGroup3.Tools("CogToolBlock1"), CogToolBlock)
        Add_data("CCD1")
        Add_data("CCD2")
        Add_data("CCD3")
        Add_data("CCD4")
        CogJobManagerEdit1.Subject = myJobManager
        Log_message("料号" & ComboBox1.Text & "加载成功")
    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Count_OK = 0
        Count_NG = 0
        TextBox_OK.Text = Count_OK
        TextBox_NG.Text = Count_NG
        TextBox_Total.Text = Count_OK + Count_NG
    End Sub

    Private Sub Button19_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim fname As String
        fname = "C:\图片\"
        If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
            MkDir(fname)                     '创建文件夹         
        End If
        Shell("Explorer C:\图片\", vbNormalFocus)
    End Sub

    Private Sub Button18_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Dim fname As String
        fname = "C:\图片\"
        If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
            MkDir(fname)                     '创建文件夹         
        End If
        Shell("Explorer C:\图片\", vbNormalFocus)
    End Sub

    Private Sub Button17_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Dim fname As String
        fname = "C:\图片\"
        If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
            MkDir(fname)                     '创建文件夹         
        End If
        Shell("Explorer C:\图片\", vbNormalFocus)
    End Sub

    Private Sub Button16_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Dim fname As String
        fname = "C:\图片\"
        If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
            MkDir(fname)                     '创建文件夹         
        End If
        Shell("Explorer C:\图片\", vbNormalFocus)
    End Sub


#Region "保存图片"
    Public Sub SaveImage()
        If CheckBox_Image.Checked = True Then
            Dim fname As String
            fname = "\\10.230.211.8\image\" + ComboBox1.Text + Format(Now, "yyyy年M月d日") + "\" + Label173.Text
            If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
                MkDir(fname)                     '创建文件夹         
            End If
            fname = fname + "\" + "1" + CCD3_Result.Text + "_" + Label174.Text + ".bmp"
            If Dir(fname, vbHidden) = "" Then   '查看是否有这个文件夹的文件
                Dim imagefileWrite As New CogImageFile
                imagefileWrite.Open(fname, CogImageFileModeConstants.Write)
                imagefileWrite.Append(CogRecordDisplay1.Image)
            End If
        End If
    End Sub
    Public Sub SaveImage1()

        If CheckBox_Image.Checked = True Then
            Dim fname As String
            fname = "\\10.230.211.8\image\" + ComboBox1.Text + Format(Now, "yyyy年M月d日") + "\" + Label173.Text
            If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
                MkDir(fname)                     '创建文件夹         
            End If
            fname = fname + "\" + "2" + CCD1_Result.Text + "_" + Label174.Text + ".bmp"
            If Dir(fname, vbHidden) = "" Then   '查看是否有这个文件夹的文件
                Dim imagefileWrite As New CogImageFile
                imagefileWrite.Open(fname, CogImageFileModeConstants.Write)
                imagefileWrite.Append(CogRecordDisplay2.Image)
            End If
        End If
    End Sub
    Private Sub SaveImage2()
        If CheckBox_Image.Checked = True Then
            Dim fname As String
            fname = "\\10.230.211.8\image\" + ComboBox1.Text + Format(Now, "yyyy年M月d日") + "\" + Label173.Text
            If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
                MkDir(fname)                     '创建文件夹         
            End If
            fname = fname + "\" + "3" + CCD4_Result.Text + "_" + Label174.Text + ".bmp"
            If Dir(fname, vbHidden) = "" Then   '查看是否有这个文件夹的文件
                Dim imagefileWrite As New CogImageFile
                imagefileWrite.Open(fname, CogImageFileModeConstants.Write)
                imagefileWrite.Append(CogRecordDisplay3.Image)
            End If
        End If
    End Sub
    Private Sub SaveImage3()
        If CheckBox_Image.Checked = True Then
            Dim fname As String
            fname = "\\10.230.211.8\image\" + ComboBox1.Text + Format(Now, "yyyy年M月d日") + "\" + Label173.Text
            If Dir(fname, vbDirectory) = "" Then '判断文件夹是否存在   
                MkDir(fname)                     '创建文件夹         
            End If
            fname = fname + "\" + "4" + CCD2_Result.Text + "_" + Label174.Text + ".bmp"
            If Dir(fname, vbHidden) = "" Then   '查看是否有这个文件夹的文件
                Dim imagefileWrite As New CogImageFile
                imagefileWrite.Open(fname, CogImageFileModeConstants.Write)
                imagefileWrite.Append(CogRecordDisplay4.Image)
            End If
        End If
    End Sub
#End Region

  

End Class
