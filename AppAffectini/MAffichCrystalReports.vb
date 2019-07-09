Module MAffichCrystalReports

    Public Sub AffichCrystalReport()

        Dim report1 As New CrystalReport1

        Form2.Show()
        AffectINI.dtt.Rows.Clear()
        AffectINI.dtt.Columns.Clear()

        Dim req2 As String = "SELECT * FROM CrystalReport"
        Requette(req2)
        report1.SetDataSource(AffectINI.dtt)
        Form2.CrystalReportViewer1.ReportSource = report1
        Form2.CrystalReportViewer1.Refresh()
        Form2.CrystalReportViewer1.RefreshReport()

        AffectINI.dtt.Rows.Clear()
        AffectINI.dtt.Columns.Clear()
    End Sub
End Module
