Imports System.Threading

Public Class FormMyTasks
    Private myController As New ProductRequestController
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim MyTasksCriteria As String
    Dim HistoryCriteria As String
    Dim DS As New DataSet
    Dim MyTasksBS As BindingSource
    Dim HistoryBS As BindingSource
    Private Sub FormFindProductRequest_Load(sender As Object, e As EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        If User.can("ViewAllTx") Then
            MyTasksCriteria = " where status > 0 and status < 7"
            HistoryCriteria = " where status >= 7"
        ElseIf User.can("View Supply Chain HK") Then
            MyTasksCriteria = String.Format(" where (status = {0} or (status = {1} and mdapproval isnull))", Int(ProductRequestStatusEnum.StatusValidatedbyMDirector), Int(ProductRequestStatusEnum.StatusValidatedbyDirector))
            HistoryCriteria = String.Format(" where status >= {0}", Int(ProductRequestStatusEnum.StatusCompleted))
        ElseIf User.can("View Supply Chain TW") Then
        Else
            MyTasksCriteria = String.Format("where ((deptapproval = '{0}' and (status = {1} or status = {2})) or (mdapproval = '{0}' and status = 3)) ", DirectCast(User.identity, UserController).userid, Int(ProductRequestStatusEnum.StatusNew), Int(ProductRequestStatusEnum.StatusResubmit))
            HistoryCriteria = String.Format("where ((deptapproval = '{0}' and status > 1) or (mdapproval = '{0}' and status > 3) or(createdby = '{0}' and  status > 0)) ", DirectCast(User.identity, UserController).userid)
        End If
        Try
            DS = New DataSet
            If myController.MyTasksloaddata(DS, MyTasksCriteria, HistoryCriteria) Then
                ProgressReport(4, "InitData")
                ProgressReport(1, "Loading Data.Done!")
                ProgressReport(5, "Continuous")
            End If
        Catch ex As Exception
            ProgressReport(1, "Loading Data. Error::" & ex.Message)
            ProgressReport(5, "Continuous")
        End Try
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel1.Text = message

                    Case 4
                        MyTasksBS = New BindingSource
                        MyTasksBS.DataSource = DS.Tables(0)
                        HistoryBS = New BindingSource
                        HistoryBS.DataSource = DS.Tables(1)
                        DataGridView1.AutoGenerateColumns = False
                        DataGridView1.DataSource = MyTasksBS
                        DataGridView2.AutoGenerateColumns = False
                        DataGridView2.DataSource = HistoryBS
                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

    End Sub

    Private Sub RefreshToolStripButton_Click(sender As Object, e As EventArgs) Handles RefreshToolStripButton.Click
        loaddata()
    End Sub




    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim drv = MyTasksBS.Current
        Dim myform = New FormProductRequest(drv.row.item("id"), TxEnum.ValidateRecord)
        myform.ShowDialog()
        If myform.IsModified Then
            loaddata()
        End If
    End Sub
    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        Dim drv = HistoryBS.Current
        Dim myform = New FormProductRequest(drv.row.item("id"), TxEnum.HistoryRecord)
        myform.ShowDialog()
        If myform.IsModified Then
            loaddata()
        End If
    End Sub
End Class