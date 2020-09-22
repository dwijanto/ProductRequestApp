Imports System.Windows.Forms

Public Class DialogAddUpdBPAddress
    Dim drv As DataRowView
    Dim BPartnerBS As BindingSource

    Public Event RefreshDataGridView()

    Public Sub New(ByRef DRV As DataRowView, ByVal BPartnerBS As BindingSource)
        InitializeComponent()
        Me.drv = DRV
        Me.BPartnerBS = BPartnerBS
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.validate() Then
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
            ComboBox1.DataBindings.Clear()

            TextBox1.DataBindings.Clear()
            TextBox2.DataBindings.Clear()
            TextBox3.DataBindings.Clear()
            TextBox4.DataBindings.Clear()
            TextBox5.DataBindings.Clear()
            TextBox6.DataBindings.Clear()
            TextBox7.DataBindings.Clear()
            TextBox8.DataBindings.Clear()
            TextBox9.DataBindings.Clear()

            ComboBox1.DataBindings.Add(New Binding("Text", drv, "addresstype", True, DataSourceUpdateMode.OnPropertyChanged))
            TextBox9.DataBindings.Add(New Binding("Text", drv, "addressid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
            TextBox1.DataBindings.Add(New Binding("Text", drv, "contactname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
            TextBox2.DataBindings.Add(New Binding("Text", drv, "line1", True, DataSourceUpdateMode.OnPropertyChanged, ""))
            TextBox5.DataBindings.Add(New Binding("Text", drv, "line2", True, DataSourceUpdateMode.OnPropertyChanged, ""))
            TextBox6.DataBindings.Add(New Binding("Text", drv, "line3", True, DataSourceUpdateMode.OnPropertyChanged, ""))
            TextBox3.DataBindings.Add(New Binding("Text", drv, "location", True, DataSourceUpdateMode.OnPropertyChanged, ""))
            TextBox7.DataBindings.Add(New Binding("Text", drv, "region", True, DataSourceUpdateMode.OnPropertyChanged, ""))
            TextBox8.DataBindings.Add(New Binding("Text", drv, "country", True, DataSourceUpdateMode.OnPropertyChanged, ""))
            TextBox4.DataBindings.Add(New Binding("Text", drv, "bpartnerfullname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim myhelper As New DialogHelper(BPartnerBS)
        myhelper.DataGridView1.Columns(0).DataPropertyName = "bpcode"
        myhelper.DataGridView1.Columns(0).HeaderText = "AR Code"
        myhelper.DataGridView1.Columns(1).DataPropertyName = "bpname"
        myhelper.DataGridView1.Columns(1).HeaderText = "Company Name"
        myhelper.DataGridView1.Columns(1).Width = 300
        myhelper.DataGridView1.Columns(2).Visible = False
        myhelper.DataGridView1.Columns(3).Visible = False
        myhelper.Filter = "bpcode like '%{0}%' or bpname like '%{0}%'"
        myhelper.ShowDialog()
        Dim mydrv = BPartnerBS.Current
        TextBox4.Text = mydrv.item("bpartnerfullname")
        drv.Item("bpid") = mydrv.item("id")
    End Sub

    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox4.TextChanged
        RaiseEvent RefreshDataGridView()
    End Sub
End Class
