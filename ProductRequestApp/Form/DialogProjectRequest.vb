Imports System.Windows.Forms

Public Class DialogProjectRequest
    Dim DRV As DataRowView
    Dim CMMFBS As BindingSource
    Dim ExpensesTypeBS As BindingSource
    Public Sub New(drv, CMMFBS, ExpensesTypeBS)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Me.DRV = drv
        initdata()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DialogProjectRequest_Load(sender As Object, e As EventArgs) Handles Me.Load

    End Sub

    Private Sub initdata()
        TextBox1.DataBindings.Clear()

        TextBox1.DataBindings.Add(New Binding("Text", DRV, "qty", True, DataSourceUpdateMode.OnPropertyChanged, ""))
    End Sub

End Class
