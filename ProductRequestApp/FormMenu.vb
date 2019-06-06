Imports System.Reflection
Public Enum TxEnum
    NewRecord = 1
    CopyRecord = 2
    UpdateRecord = 3
End Enum
Public Class FormMenu

    Private UserInfo1 As UserInfo = UserInfo.getInstance
    Dim HasError As Boolean = True
    Private userid As String
    Private myuser As UserController

    Public Sub New()
        myuser = New UserController
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

        Try
            'userinfo1 = UserInfo.getInstance
            userinfo1.Userid = Environment.UserDomainName & "\" & Environment.UserName
            userinfo1.computerName = My.Computer.Name
            UserInfo1.ApplicationName = "Product Request Apps"
            userinfo1.Username = "N/A"
            userinfo1.isAuthenticate = False
            userinfo1.Role = 0
            UserInfo1.isAdmin = DataAccess.isAdmin(UserInfo1.Userid)
            HasError = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub FormMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If HasError Then
            Me.Close()
            Exit Sub
        End If

        Try
            userid = userinfo1.Userid 'Environment.UserDomainName & "\" & Environment.UserName

            Dim myAD = New ADPrincipalContext
            Dim UserInfo As List(Of ADPrincipalContext) = New List(Of ADPrincipalContext)
            If myAD.GetInfo(userid) Then
                myuser.Model.ADDUPDUserManager(ADPrincipalContext.ADPrincipalContexts)
            Else
                MessageBox.Show(myAD.ErrorMessage)
                Me.Close()
                Exit Sub
            End If


            Dim mydata As DataSet = myuser.findByUserid(userid.ToLower)
            If mydata.tables(0).rows.count > 0 Then
                Dim identity = myuser.findIdentity(mydata.Tables(0).rows(0).item("id"))
                User.setIdentity(identity)
                User.login(identity)
                User.IdentityClass = myuser
                DataAccess.LogLogin(UserInfo1.Userid)
                Me.Text = GetMenuDesc()
                Me.Location = New Point(300, 10)
                MenuHandles()
            Else
                'disable menubar
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
        End Try

    End Sub
   
    Public Function GetMenuDesc() As String
        Return "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & DataAccess.GetHostName & ", Database: " & DataAccess.GetDataBaseName & ", Userid: " & UserInfo1.Userid 'HelperClass1.UserId
    End Function

    Private Sub MenuHandles()

        ''Admin Function
        ''MasterToolStripMenuItem.Visible = userinfo1.isAdmin
        'AddHandler VendorAssignmentQEUserToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler FirstCmmfToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler MissingVendorToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler RBACToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler UserToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

        'Dim identity As UserController = User.getIdentity
        MasterToolStripMenuItem.Visible = User.can("View Master")
        AdminToolStripMenuItem.Visible = User.can("View Admin")
       
    End Sub

    Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ctrl As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim assembly1 As Assembly = Assembly.GetAssembly(GetType(FormMenu))
        Dim frm As Object = CType(assembly1.CreateInstance(assembly1.GetName.Name.ToString & "." & ctrl.Tag.ToString, True), Form)
        Dim myform = frm.GetInstance
        myform.show()
        myform.windowstate = FormWindowState.Normal
        myform.activate()
    End Sub



    Private Sub MyTasksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MyTasksToolStripMenuItem.Click

    End Sub

    Private Sub UserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UserToolStripMenuItem.Click

    End Sub

    Private Sub MasterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MasterToolStripMenuItem.Click

    End Sub
End Class
