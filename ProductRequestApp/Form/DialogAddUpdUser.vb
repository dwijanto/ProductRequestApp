Imports System.Windows.Forms

Public Class DialogAddUpdUser
    Dim DRV As DataRowView
    Dim ApprovalBS As BindingSource

    Public Sub New(ByVal drv As DataRowView, ApprovalBS As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.DRV = drv
        Me.ApprovalBS = ApprovalBS
    End Sub

    Public Shadows Function Validate() As Boolean
        Dim CBDrv As DataRowView
        CBDrv = ComboBox1.SelectedItem
        If Not IsNothing(CBDrv) Then
            If CBDrv.Row.Item("id") = 0 Then
                DRV.Row.Item("approvalid") = DBNull.Value
            End If
            DRV.Row.Item("approvaltype") = CBDrv.Row.Item("approvaltype")
        End If
        Return True
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click    
        Me.Validate()
        DRV.EndEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        DRV.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DialogAddUpdUser_Load(sender As Object, e As EventArgs) Handles Me.Load
        InitData()
    End Sub

    Private Sub InitData()
        TextBox1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()
        CheckBox1.DataBindings.Clear()

        ComboBox1.DataSource = ApprovalBS
        ComboBox1.DisplayMember = "approvaltype"
        ComboBox1.ValueMember = "id"

        TextBox1.DataBindings.Add(New Binding("Text", DRV, "userid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("Text", DRV, "username", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("Text", DRV, "email", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "approvalid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        CheckBox1.DataBindings.Add(New Binding("checked", DRV, "isactive", True, DataSourceUpdateMode.OnPropertyChanged))

    End Sub

End Class
