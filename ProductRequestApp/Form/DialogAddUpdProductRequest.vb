Imports System.Windows.Forms

Public Class DialogAddUpdProductRequest
    Dim drv As DataRowView
    Dim CMMFBS As BindingSource
    Dim ExpensesTypeBS As BindingSource

    Public Event RefreshDataGridView()

    Public Sub New(ByRef DRV As DataRowView, ByVal CMMFBS As BindingSource, ByVal ExpensesTypeBS As BindingSource)
        InitializeComponent()
        Me.drv = DRV
        Me.ExpensesTypeBS = ExpensesTypeBS
        Me.CMMFBS = CMMFBS
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.validate() Then
            drv.Row.Item("createddate") = Now
            drv.Row.Item("createdby") = DirectCast(User.identity, UserController).username
            drv.EndEdit()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If

    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(Button2, "")
        If TextBox4.Text = "" Then
            myret = False
            ErrorProvider1.SetError(Button2, "Please select from helper.")
        End If

        ErrorProvider1.SetError(Button1, "")
        If TextBox2.Text = "" Then
            myret = False
            ErrorProvider1.SetError(Button1, "Please select from helper.")
        End If

        ErrorProvider1.SetError(TextBox1, "")
        If TextBox1.Text = "" Then
            myret = False
            ErrorProvider1.SetError(TextBox1, "Quantity cannot be blank.")
        End If

        Return myret

    End Function

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        drv.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DialogAddUpdProjectRequest_Load(sender As Object, e As EventArgs) Handles Me.Load
        InitData()
    End Sub

    Private Sub InitData()
        Try
            TextBox1.DataBindings.Clear()
            TextBox2.DataBindings.Clear()
            TextBox3.DataBindings.Clear()
            TextBox4.DataBindings.Clear()

            TextBox1.DataBindings.Add(New Binding("Text", drv, "qty", True, DataSourceUpdateMode.OnPropertyChanged, ""))
            TextBox2.DataBindings.Add(New Binding("Text", drv, "localdescription", True, DataSourceUpdateMode.OnPropertyChanged, ""))
            TextBox3.DataBindings.Add(New Binding("Text", drv, "remarks", True, DataSourceUpdateMode.OnPropertyChanged, ""))
            TextBox4.DataBindings.Add(New Binding("Text", drv, "expensesname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim myhelper As New DialogHelper(ExpensesTypeBS)
        myhelper.DataGridView1.Columns(0).DataPropertyName = "expensesacc"
        myhelper.DataGridView1.Columns(0).HeaderText = "Expenses Account"
        myhelper.DataGridView1.Columns(1).DataPropertyName = "expensesname"
        myhelper.DataGridView1.Columns(1).HeaderText = "Expenses Name"
        myhelper.DataGridView1.Columns(1).Width = 300
        myhelper.DataGridView1.Columns(2).Visible = False
        myhelper.DataGridView1.Columns(3).Visible = False
        myhelper.Filter = "expensesacc like '%{0}%' or expensesname like '%{0}%'"
        myhelper.ShowDialog()
        Dim mydrv = ExpensesTypeBS.Current
        TextBox4.Text = mydrv.item("expensesname")
        drv.Item("expensestypeid") = mydrv.item("id")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myhelper As New DialogHelper(CMMFBS)
        myhelper.Size = New Point(720, 500)
        myhelper.DataGridView1.Columns(0).DataPropertyName = "cmmfstring"
        myhelper.DataGridView1.Columns(0).HeaderText = "CMMF"
        myhelper.DataGridView1.Columns(1).DataPropertyName = "commercialcode"
        myhelper.DataGridView1.Columns(1).HeaderText = "Commercial Code"
        myhelper.DataGridView1.Columns(2).DataPropertyName = "localdescription"
        myhelper.DataGridView1.Columns(2).HeaderText = "Description"
        myhelper.DataGridView1.Columns(2).Width = 400
        myhelper.DataGridView1.Columns(3).Visible = False
        myhelper.Filter = "cmmfstring like '%{0}%' or commercialcode like '%{0}%'  or localdescription like '%{0}%'"
        If myhelper.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim mydrv = CMMFBS.Current
            TextBox2.Text = mydrv.item("localdescription")
            drv.Item("localdescription") = mydrv.item("localdescription")
            drv.Item("commercialcode") = mydrv.item("commercialcode")
            drv.Item("price") = mydrv.item("price")
            drv.Item("cmmf") = mydrv.item("cmmf")
        End If

    End Sub


    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged, TextBox4.TextChanged
        RaiseEvent RefreshDataGridView()
    End Sub
End Class
