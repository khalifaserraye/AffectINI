Imports System.Data.OleDb
Module MFillSemestre

    Sub FillSemestre()
        AffectINI.annee = AffectINI.cmbRest1Annee.Text + ".accdb"
        AffectINI.connexionNihad = "Provider=Microsoft.ace.OleDb.12.0;DATA SOURCE = ..\..\Sauvegarde\" + AffectINI.annee
        AffectINI.cont = New OleDbConnection("Provider=Microsoft.ace.OleDb.12.0;DATA SOURCE = ..\..\Sauvegarde\" + AffectINI.annee)


        If AffectINI.cont.State = ConnectionState.Closed Then AffectINI.cont.Open()
        AffectINI.promo = AffectINI.cmbRes1Promo.Text
        AffectINI.cmbRes1Semestre.DataSource = GetTable("select DISTINCT  [exam].[semestre] FROM [exam], [ETUDIANTS] ,[affect] where [ETUDIANTS].[Matricule]=[affect].[matricule] AND [exam].[gp]=[affect].[gp] AND [ETUDIANTS].[Promo]='" + AffectINI.promo + "'")
        AffectINI.cmbRes1Semestre.DisplayMember = "semestre"
        If AffectINI.cont.State = ConnectionState.Open Then AffectINI.cont.Close()
    End Sub
End Module
