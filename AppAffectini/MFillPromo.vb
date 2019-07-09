Imports System.Data.OleDb
Module MFillPromo

    Sub FillPromo()
        AffectINI.annee = AffectINI.cmbRest1Annee.Text + ".accdb"
        AffectINI.connexionNihad = "Provider=Microsoft.ace.OleDb.12.0;DATA SOURCE = ..\..\Sauvegarde\" + AffectINI.annee
        AffectINI.cont = New OleDbConnection(AffectINI.connexionNihad)

        If AffectINI.cont.State = ConnectionState.Closed Then AffectINI.cont.Open()
        AffectINI.cmbRes1Promo.DataSource = GetTable("SELECT DISTINCT  [ETUDIANTS].[Promo] FROM [ETUDIANTS] ,[affect] where [ETUDIANTS].[Matricule]=[affect].[matricule] ")
        AffectINI.cmbRes1Promo.DisplayMember = "Promo"
        If AffectINI.cont.State = ConnectionState.Open Then AffectINI.cont.Close()
    End Sub
End Module
