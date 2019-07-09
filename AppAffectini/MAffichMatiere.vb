Imports System.Data.OleDb

Module MAffichMatiere

    Public Sub AffichMatiere(ByRef Matiere As String)

        AffectINI.dtt.Rows.Clear()
        AffectINI.dtt.Columns.Clear()

        Dim da As New OleDbDataAdapter
        Dim daa As New OleDbDataAdapter

        Dim dt1 As New DataTable
        Dim dt2 As New DataTable

        Dim matricule As String
        Dim req As String

        Dim reqRecup As String = "SELECT [emploi].[gp] FROM emploi WHERE [emploi].[module] = '" + Matiere + "'"

        da = New OleDbDataAdapter(reqRecup, AffectINI.connAccess)
        da.Fill(dt1)

        Dim numeroGp As Integer = dt1(0)(0)

        Dim req2 As String = "SELECT [affect].[matricule] FROM affect WHERE affect.gp = " & numeroGp
        daa = New OleDbDataAdapter(req2, AffectINI.connAccess)
        daa.Fill(dt2)

        For Each dtr As DataRow In dt2.Rows

            matricule = dtr(0)
            req = "SELECT [ETUDIANTS].[Matricule], [NomEtud], [Prenoms], [Sect], [Gr], [affect].[local], [affect].[pos], [emploi].[module], [emploi].[date1], [emploi].[debut], [emploi].[fin] FROM ETUDIANTS, affect, emploi WHERE ETUDIANTS.Matricule = '" + matricule + "' AND ( affect.matricule = '" + matricule + "' AND affect.gp = " & numeroGp & ") AND emploi.gp = " & numeroGp

            Requette(req)
        Next
    End Sub
End Module
