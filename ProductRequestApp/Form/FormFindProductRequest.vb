Imports System.Threading

Public Class FormFindProductRequest
    Private myController As New ProductRequestController
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim Criteria As String
    Dim userid = DirectCast(User.identity, UserController).userid
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
            Criteria = ""
        Else
            Criteria = String.Format("where createdby = '{0}' and (status <= 1)", DirectCast(User.identity, UserController).userid)
        End If
        Try
            If myController.Findloaddata(Criteria) Then
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
                        DataGridView1.AutoGenerateColumns = False
                        DataGridView1.DataSource = myController.BS

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

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        loaddata()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim drv = myController.BS.Current
        Dim myform = New FormProductRequest(drv.row.item("id"), TxEnum.UpdateRecord)
        myform.ShowDialog()
        If myform.IsModified Then
            loaddata()
        End If
    End Sub

    'Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs)
    '    Dim drv As DataRowView = myController.BS.Current
    '    If IsNothing(drv) Then
    '        Exit Sub
    '    End If
    '    If drv.Row.Item("status") = ProductRequestStatusEnum.StatusCancelled Then
    '        MessageBox.Show("This record status is Cancelled.")
    '        Exit Sub
    '    End If

    '    If MessageBox.Show("Do you want to cancel this record?", "Cancel", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
    '        Dim remarks As String = InputBox("Please input some comment.")
    '        drv.Row.Item("status") = ProductRequestStatusEnum.StatusCancelled

    '        If drv.Row.Item("status") = ProductRequestStatusEnum.StatusNew Then
    '            Dim approvaldrv = myController.GetActionBS.AddNew
    '            approvaldrv.Row.Item("status") = drv.Row.Item("status")
    '            approvaldrv.Row.Item("statusname") = "Cancelled"
    '            approvaldrv.Row.Item("modifiedby") = userid
    '            approvaldrv.Row.Item("latestupdate") = Now
    '            approvaldrv.Row.Item("remarks") = remarks
    '            approvaldrv.Row.Item("prhdid") = drv.Item("id")

    '            approvaldrv.EndEdit()
    '            Logger.log(String.Format("** Submit {0}**", userid))
    '            If myController.save() Then
    '                'SendEmail()
    '                Dim sendto = drv.Row.Item("deptapproval")
    '                Dim sendtoName = drv.Row.Item("")
    '                Dim StatusName = "Cancelled"
    '                Dim cc As 
    '                Logger.log(String.Format("SendTo: {0}, SendTo Name: {1}, StatusName: {2}", sendto, sendtoName, StatusName))
    '                Dim myEmail As New PREmail
    '                If Not myEmail.Execute(sendto, sendtoName, StatusName, drv, CC) Then
    '                    Logger.log(String.Format("Error Message: {0}", myEmail.ErrorMessage))
    '                Else
    '                    Logger.log("Email Sent")
    '                End If
    '            End If
    '        Else
    '            myController.save()
    '        End If


    '    End If


    'End Sub
End Class