Imports System.Data.OleDb

Module MAffichEtudiant

    Public Sub AffichEtudiant(ByRef Matricule As String)

        AffectINI.dtt.Rows.Clear()
        AffectINI.dtt.Columns.Clear()

        Dim da As New OleDbDataAdapter
        Dim dt As New DataTable

        Dim reqRecup As String
        Dim numeroGp As UInteger
        Dim req As String

        reqRecup = "SELECT [affect].[gp] FROM affect WHERE affect.matricule = '" + Matricule + "'"

        da = New OleDbDataAdapter(reqRecup, AffectINI.connAccess)
        da.Fill(dt)

        For Each dtr As DataRow In dt.Rows
            numeroGp = dtr(0)

            req = "SELECT [ETUDIANTS].[Matricule], [NomEtud], [Prenoms], [Sect], [Gr], [affect].[local], [affect].[pos], [emploi].[module], [emploi].[date1], [emploi].[debut], [emploi].[fin] FROM ETUDIANTS, affect, emploi WHERE ETUDIANTS.Matricule = '" + Matricule + "' AND ( affect.matricule = '" + Matricule + "' AND affect.gp = " & numeroGp & ") AND emploi.gp = " & numeroGp

            Requette(req)
        Next

    End Sub
End Module
