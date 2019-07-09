Module MFillModule

    Sub FillModule()

        AffectINI.promo = AffectINI.cmbRes1Promo.Text
        AffectINI.sem = AffectINI.cmbRes1Semestre.Text
        AffectINI.c = AffectINI.cmbRest1TpeExam.Text
        If AffectINI.cont.State = ConnectionState.Closed Then AffectINI.cont.Open()
        AffectINI.cmbRest1Module.DataSource = GetTable("select DISTINCT  [MODULES].[Code_Mat] FROM [MODULES] ,[exam] ,[emploi]  where [MODULES].[Code_Mat]=[emploi].[module] AND [MODULES].[Niveau]='" + AffectINI.promo + "' And [exam].[semestre]='" + AffectINI.sem + "' AND [exam].[type_exam]='" + AffectINI.c + "'  AND  [exam].[gp]=[emploi].[gp]")
        AffectINI.cmbRest1Module.DisplayMember = "Code_Mat"
        If AffectINI.cont.State = ConnectionState.Open Then AffectINI.cont.Close()
    End Sub
End Module
