Imports System.Data.OleDb
Module MFillExam

    Sub FillExam()
        AffectINI.annee = AffectINI.cmbRest1Annee.Text + ".accdb"
        AffectINI.connexionNihad = "Provider=Microsoft.ace.OleDb.12.0;DATA SOURCE = ..\..\Sauvegarde\" + AffectINI.annee
        AffectINI.cont = New OleDbConnection("Provider=Microsoft.ace.OleDb.12.0;DATA SOURCE = ..\..\Sauvegarde\" + AffectINI.annee)
        AffectINI.promo = AffectINI.cmbRes1Promo.Text
        AffectINI.sem = AffectINI.cmbRes1Semestre.Text
        If AffectINI.cont.State = ConnectionState.Closed Then AffectINI.cont.Open()
        AffectINI.cmbRest1TpeExam.DataSource = GetTable("select DISTINCT  [exam].[type_exam] FROM [exam],[ETUDIANTS] ,[affect] where [ETUDIANTS].[Matricule]=[affect].[matricule] AND  [exam].[gp]=[affect].[gp] AND [ETUDIANTS].[Promo]='" + AffectINI.promo + "' And [exam].[semestre]='" + AffectINI.sem + "'")
        AffectINI.cmbRest1TpeExam.DisplayMember = "type_exam"
        If AffectINI.cont.State = ConnectionState.Open Then AffectINI.cont.Close()
    End Sub
End Module
