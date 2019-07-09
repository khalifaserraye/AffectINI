Imports System.Data.OleDb

Module MAffichSale

    Public Sub AffichSalle(ByRef Salle As String)

        AffectINI.dtt.Rows.Clear()
        AffectINI.dtt.Columns.Clear()

        Dim da As New OleDbDataAdapter
        Dim dt As New DataTable

        Dim reqRecup As String
        Dim matricule As String
        Dim numeroGp As Integer
        Dim req As String

        reqRecup = "SELECT [affect].[matricule], [affect].[gp] FROM affect WHERE affect.local = '" + Salle + "'"

        da = New OleDbDataAdapter(reqRecup, AffectINI.connAccess)
        da.Fill(dt)

        For Each dtr As DataRow In dt.Rows
            matricule = dtr(0)
            numeroGp = dtr(1)

            req = "SELECT [ETUDIANTS].[Matricule], [NomEtud], [Prenoms], [Sect], [Gr], [affect].[local], [affect].[pos], [emploi].[module], [emploi].[date1], [emploi].[debut], [emploi].[fin] FROM ETUDIANTS, affect, emploi WHERE ETUDIANTS.Matricule = '" + matricule + "' AND ( affect.matricule = '" + matricule + "' AND affect.gp = " & numeroGp & ") AND emploi.gp = " & numeroGp

            Requette(req)
        Next
    End Sub
End Module
