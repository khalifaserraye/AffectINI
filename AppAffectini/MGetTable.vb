Imports System.Data.OleDb
Module MGetTable

    Function GetTable(ByVal selectCommand As String) As DataTable

        Dim tbl As New DataTable
        AffectINI.annee = AffectINI.anne.Text + ".accdb"
        AffectINI.cont = New OleDbConnection("Provider=Microsoft.ace.OleDb.12.0;DATA SOURCE = ..\..\Sauvegarde\" + AffectINI.annee)
        AffectINI.cmde = New OleDbCommand("", AffectINI.cont)
        AffectINI.annee = AffectINI.cmbRest1Annee.Text
        Try
            If AffectINI.cont.State = ConnectionState.Closed Then AffectINI.cont.Open()
            AffectINI.cmde.CommandText = selectCommand
            tbl.Load(AffectINI.cmde.ExecuteReader())

        Catch ex As Exception
        Finally
            If AffectINI.cont.State = ConnectionState.Open Then AffectINI.cont.Close()

        End Try

        Return tbl
    End Function
End Module
