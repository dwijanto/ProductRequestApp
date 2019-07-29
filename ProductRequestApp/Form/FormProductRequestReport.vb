Imports System.IO
Public Class FormProductRequestReport
    Dim SaveFileDialog1 As New SaveFileDialog
    Dim mycallback As FormatReportDelegate = AddressOf formatReport
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim criteria As String
        criteria = String.Format(" where hd.applicantdate >= '{0:yyyy-MM-dd}' and hd.applicantdate <= '{1:yyyy-MM-dd}'", DateTimePicker1.Value, DateTimePicker2.Value)
        Dim sqlstr As String = String.Format("select hd.*,marketing.getstatusname(hd.status) as status,dt.*,c.*,e.*,bpa.*,bp.* from marketing.prhd hd" &
                               " left join marketing.prdt dt on hd.id = dt.prhdid" &
                               " left join marketing.cmmf c on c.cmmf = dt.cmmf" &
                               " left join marketing.expensestype e on e.id = dt.expensestypeid" &
                               " left join marketing.bpaddress bpa on bpa.id = hd.bpartnerid" &
                               " left join marketing.bpartner bp on bp.id = bpa.bpid {0}", criteria)


        'Dim myreport As New ExportToExcelFile(Me, sqlstr, "")
        SaveFileDialog1.FileName = String.Format("{0:yyyyMMdd}_ProductRequest.xlsx", Date.Today)
        SaveFileDialog1.DefaultExt = "xlsx"

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim myform As ExportToExcelFile = New ExportToExcelFile(Me, sqlstr, IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), mycallback)
            myform.Run(sender, New System.EventArgs)
        End If
    End Sub

    Sub formatReport()

    End Sub

End Class