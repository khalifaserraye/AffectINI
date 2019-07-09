Imports System.Data.OleDb
Imports System.IO

Public Class AffectINI

    Public moduleChoisit As String

    Public dtt As New DataTable

    'Creer la variable de connexion : 

    Public connAccess As OleDbConnection
    Public pathEtatSalle As String

    Dim daReq As New OleDbDataAdapter
    Dim dtReq As New DataTable

    Dim daSalle As New OleDbDataAdapter
    Dim dtSalle As New DataTable

    Dim DaMatiere As New OleDbDataAdapter
    Dim DtMatiere As New DataTable

    Public annee As String

    Public cont As New OleDbConnection
    Public cmde As New OleDbCommand
    Public promo As String
    Public sem As String
    Public bdd As String
    Public connexionNihad As String
    Public c As String

    Private Sub AffectINI_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        pnlToolbarAffect.Visible = False
        pnlToolbarMenu.Visible = True
        btnMenu.Enabled = False

        pnlRestauration2.Visible = False
        pnlRestauration1.Visible = False
        pnlEtatSortie.Visible = False
        pnlEtatSalle.Visible = False
        pnlSalleSuite.Visible = False
        pnlEtatSalleYes.Visible = False
        pnlMenu.Visible = True
        pnlLogin.Visible = False
        pnlAffectation1.Visible = False
        pnlAffectation2.Visible = False
        pnlAffectation3.Visible = False
        pnlEtatSalleNo.Visible = False

        defultpath.Text = "../../Affectation.accdb"
    End Sub

    Private Sub btnSeConnecter_Click(sender As Object, e As EventArgs) Handles btnSeConnecter.Click

        pnlRestauration2.Visible = False
        pnlRestauration1.Visible = False
        pnlEtatSortie.Visible = False
        pnlEtatSalle.Visible = False
        pnlSalleSuite.Visible = False
        pnlEtatSalleYes.Visible = False
        pnlMenu.Visible = False
        pnlLogin.Visible = True
        pnlAffectation1.Visible = False
        pnlAffectation2.Visible = False
        pnlAffectation3.Visible = False
        pnlEtatSalleNo.Visible = False

        pnlLeftSlide1.Visible = True
        pnlLeftSlide2.Visible = False
        pnlLeftSlide3.Visible = False
        pnlLeftSlide4.Visible = False
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click

        If (Me.txtLogIn.Text = "esi" And Me.txtPasswd.Text = "esi2019") Then

            pnlRestauration2.Visible = False
            pnlRestauration1.Visible = False
            pnlEtatSortie.Visible = False
            pnlEtatSalle.Visible = False
            pnlSalleSuite.Visible = False
            pnlEtatSalleYes.Visible = False
            pnlMenu.Visible = True
            pnlLogin.Visible = False
            pnlAffectation1.Visible = False
            pnlAffectation2.Visible = False
            pnlAffectation3.Visible = False
            pnlEtatSalleNo.Visible = False

            btnSeConnecter.Enabled = False
            btnMenu.Enabled = True
            pnlLoginSuccess.Visible = True
        Else
            Me.txtError.Visible = True
        End If
    End Sub


    Private Sub btnMenu_Click(sender As Object, e As EventArgs) Handles btnMenu.Click

        pnlToolbarMenu.Visible = False
        pnlToolbarAffect.Visible = True

        pnlRestauration2.Visible = False
        pnlRestauration1.Visible = False
        pnlEtatSortie.Visible = False
        pnlEtatSalle.Visible = False
        pnlSalleSuite.Visible = False
        pnlEtatSalleYes.Visible = False
        pnlMenu.Visible = False
        pnlLogin.Visible = False
        pnlAffectation1.Visible = True
        pnlAffectation2.Visible = False
        pnlAffectation3.Visible = False
        pnlEtatSalleNo.Visible = False

        pnlLeftSide1.Visible = True
        pnlLeftSide2.Visible = False
        pnlLeftSide3.Visible = False
        pnlLeftSide4.Visible = False
        pnlLeftSide5.Visible = False
        pnlLeftSide6.Visible = False
    End Sub

    Private Sub btnQuitter_Click_1(sender As Object, e As EventArgs) Handles btnQuitter.Click

        pnlLeftSlide1.Visible = False
        pnlLeftSlide2.Visible = False
        pnlLeftSlide3.Visible = True
        pnlLeftSlide4.Visible = False

        If MsgBox("Voulez vous vraiment quitter cette application ?", 36, "Quitter") = MsgBoxResult.Yes Then
            End
        End If
    End Sub

    Private Sub txtLogIn_TextChanged(sender As Object, e As EventArgs) Handles txtLogIn.Click

        txtLogIn.Clear()
    End Sub

    Private Sub txtPasswd_TextChanged(sender As Object, e As EventArgs) Handles txtPasswd.Click

        txtPasswd.Clear()
        txtPasswd.PasswordChar = "●"
    End Sub

    Private Sub btnExitLog_Click(sender As Object, e As EventArgs) Handles btnExitLog.Click

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub btnMinLog_Click(sender As Object, e As EventArgs) Handles btnMinLog.Click

        MMinFenetre.MinFenetre()
    End Sub

    Private Sub btnMaxLog_Click(sender As Object, e As EventArgs) Handles btnMaxLog.Click

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub btnMax_Click(sender As Object, e As EventArgs) Handles btnMax.Click

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub btnMin_Click(sender As Object, e As EventArgs) Handles btnMin.Click

        MMinFenetre.MinFenetre()
    End Sub

    Private Sub btnAide_Click(sender As Object, e As EventArgs) Handles btnAide.Click

        pnlLeftSlide1.Visible = False
        pnlLeftSlide2.Visible = False
        pnlLeftSlide3.Visible = False
        pnlLeftSlide4.Visible = True

        Process.Start("..\..\Page_D'aide\index.html")
    End Sub

    Private Sub btnRestauration_Click(sender As Object, e As EventArgs) Handles btnRestauration.Click

        pnlRestauration2.Visible = False
        pnlRestauration1.Visible = True
        pnlEtatSortie.Visible = False
        pnlEtatSalle.Visible = False
        pnlSalleSuite.Visible = False
        pnlEtatSalleYes.Visible = False
        pnlMenu.Visible = False
        pnlLogin.Visible = False
        pnlAffectation1.Visible = False
        pnlAffectation2.Visible = False
        pnlAffectation3.Visible = False
        pnlEtatSalleNo.Visible = False

        pnlLeftSide1.Visible = False
        pnlLeftSide2.Visible = True
        pnlLeftSide3.Visible = False
        pnlLeftSide4.Visible = False
        pnlLeftSide5.Visible = False
        pnlLeftSide6.Visible = False
    End Sub

    Private Sub BtnQuit1_Click(sender As Object, e As EventArgs) Handles BtnQuit1.Click

        pnlLeftSide1.Visible = False
        pnlLeftSide2.Visible = False
        pnlLeftSide3.Visible = False
        pnlLeftSide4.Visible = False
        pnlLeftSide5.Visible = True
        pnlLeftSide6.Visible = False

        If MsgBox("Voulez vous vraiment quitter cette application ?", 36, "Quitter") = MsgBoxResult.Yes Then
            End
        End If
    End Sub

    Private Sub btnAffect_Click(sender As Object, e As EventArgs) Handles btnAffect.Click

        pnlRestauration2.Visible = False
        pnlRestauration1.Visible = False
        pnlEtatSortie.Visible = False
        pnlEtatSalle.Visible = False
        pnlSalleSuite.Visible = False
        pnlEtatSalleYes.Visible = False
        pnlMenu.Visible = False
        pnlLogin.Visible = False
        pnlAffectation1.Visible = True
        pnlAffectation2.Visible = False
        pnlAffectation3.Visible = False
        pnlEtatSalleNo.Visible = False

        pnlLeftSide1.Visible = True
        pnlLeftSide2.Visible = False
        pnlLeftSide3.Visible = False
        pnlLeftSide4.Visible = False
        pnlLeftSide5.Visible = False
        pnlLeftSide6.Visible = False
    End Sub

    Private Sub btnEtatSalle_Click(sender As Object, e As EventArgs) Handles btnEtatSalle.Click

        dtSalle.Rows.Clear()
        dtSalle.Columns.Clear()

        dtt.Rows.Clear()
        dtt.Columns.Clear()

        pnlRestauration2.Visible = False
        pnlRestauration1.Visible = False
        pnlEtatSortie.Visible = False
        pnlEtatSalle.Visible = True
        pnlSalleSuite.Visible = False
        pnlEtatSalleYes.Visible = False
        pnlMenu.Visible = False
        pnlLogin.Visible = False
        pnlAffectation1.Visible = False
        pnlAffectation2.Visible = False
        pnlAffectation3.Visible = False
        pnlEtatSalleNo.Visible = False

        pnlLeftSide1.Visible = False
        pnlLeftSide2.Visible = False
        pnlLeftSide3.Visible = True
        pnlLeftSide4.Visible = False
        pnlLeftSide5.Visible = False
        pnlLeftSide6.Visible = False

        Dim sauv As String = dtpAnneEtatSalle.Value
        dtpAnneEtatSalle.Value = CDate(sauv)

    End Sub

    Private Sub btnMinAffect_Click(sender As Object, e As EventArgs)

        MMinFenetre.MinFenetre()
    End Sub

    Private Sub btnMaxAffect_Click(sender As Object, e As EventArgs)

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub btnExitAffect_Click(sender As Object, e As EventArgs)

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub btnMinRest_Click(sender As Object, e As EventArgs)

        MMinFenetre.MinFenetre()
    End Sub

    Private Sub btnMaxRest_Click(sender As Object, e As EventArgs)

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub btnExitRest_Click(sender As Object, e As EventArgs)

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub btnExitEtatSortie_Click(sender As Object, e As EventArgs)

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub btnMaxEtatSortie_Click(sender As Object, e As EventArgs)

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub btnSuivantAffect_Click(sender As Object, e As EventArgs)

        pnlRestauration2.Visible = False
        pnlRestauration1.Visible = False
        pnlEtatSortie.Visible = False
        pnlEtatSalle.Visible = False
        pnlSalleSuite.Visible = False
        pnlEtatSalleYes.Visible = False
        pnlMenu.Visible = False
        pnlLogin.Visible = False
        pnlEtatSalleNo.Visible = False
    End Sub

    Private Sub btnExitSaff_Click(sender As Object, e As EventArgs)

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub btnMaxSaff_Click(sender As Object, e As EventArgs)

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub btnMinEtatSortie_Click(sender As Object, e As EventArgs)

        MMinFenetre.MinFenetre()
    End Sub

    Private Sub btnYes_Click(sender As Object, e As EventArgs) Handles btnYes.Click

        pnlRestauration2.Visible = False
        pnlRestauration1.Visible = False
        pnlEtatSortie.Visible = False
        pnlEtatSalle.Visible = False
        pnlSalleSuite.Visible = False
        pnlEtatSalleYes.Visible = True
        pnlMenu.Visible = False
        pnlLogin.Visible = False
        pnlAffectation1.Visible = False
        pnlAffectation2.Visible = False
        pnlAffectation3.Visible = False
        pnlEtatSalleNo.Visible = False
    End Sub

    Private Sub btnSuite_Click(sender As Object, e As EventArgs) Handles btnSuite.Click

        'Effacer l'ancien contenu du DataGridView
        dtReq.Rows.Clear()
        dtReq.Columns.Clear()


        Dim dtDebut As New DateTime
        Dim dtFin As New DateTime

        dtDebut = dteDebut.Value.ToShortTimeString
        dtFin = dteFin.Value.ToShortTimeString

        daReq = New OleDbDataAdapter("SELECT DISTINCT [local].[salle], [emploi].[module], [emploi].[date1], [emploi].[debut], [emploi].[fin] FROM [local], [emploi] WHERE  [local].[gp] = [emploi].[gp]  AND [local].[salle] = '" + cmbSalle.SelectedValue + "' AND [emploi].[date1] = #" & dteDate.Value.ToString("MM/dd/yyyy") & "# AND ( ( [emploi].[debut] >= #" & dtDebut & "# AND [emploi].[debut] < #" & dtFin & "# ) OR ( [emploi].[debut] <= #" & dtDebut & " # AND [emploi].[fin] >= #" & dtFin & "# ) OR ( [emploi].[fin] > #" & dtDebut & "# AND [emploi].[fin] < #" & dtFin & "# ) )", connAccess)
        daReq.Fill(dtReq)
        dtgvEtatSalleYes.Visible = True
        lblAffichSalle.Text = "Eetat de la salle " + cmbSalle.SelectedValue + " : "
        lblAffichSalle.Visible = True
        dtgvEtatSalleYes.DataSource = dtReq

        If (dtReq.Rows.Count = 0) Then
            dtgvEtatSalleNo.Visible = False
            lblAfficheNo.Text = lblAfficheNo.Text + "VIDE !"
        End If
    End Sub

    Private Sub btnNo_Click(sender As Object, e As EventArgs) Handles btnNo.Click

        pnlRestauration2.Visible = False
        pnlRestauration1.Visible = False
        pnlEtatSortie.Visible = False
        pnlEtatSalle.Visible = False
        pnlSalleSuite.Visible = False
        pnlEtatSalleYes.Visible = False
        pnlMenu.Visible = False
        pnlLogin.Visible = False
        pnlAffectation1.Visible = False
        pnlAffectation2.Visible = False
        pnlAffectation3.Visible = False
        pnlEtatSalleNo.Visible = True

        dtgvEtatSalleNo.Visible = True

        'Effacer l'ancien contenu du DataGridView

        dtReq.Rows.Clear()
        dtReq.Columns.Clear()

        dtt.Rows.Clear()
        dtt.Columns.Clear()

        lblAfficheNo.Text = "Etat de la salle " + cmbSalle.SelectedValue + " : "

        'La requete : 

        'Dim daDate = New OleDbDataAdapter("Select gp from [emploi] where [emploi].[date1] = #" & dteDate.Value.ToString("MM/dd/yyyy") & "#", connAccess)
        'Dim dtDate As New DataTable
        'daDate.Fill(dtDate)

        Dim daLocal = New OleDbDataAdapter("Select [local].gp from [local] where [local].[salle] ='" + cmbSalle.SelectedValue + "'", connAccess)
        Dim dtLocal As New DataTable
        daLocal.Fill(dtLocal)

        Dim grLocal As Integer
        Dim req As String

        For Each dtrLocal As DataRow In dtLocal.Rows

            grLocal = dtrLocal(0)

            req = "SELECT [local].[salle], [emploi].[module], [emploi].[date1], [emploi].[debut], [emploi].[fin] FROM [local], [emploi] WHERE ( [local].[salle] = '" + cmbSalle.SelectedValue + "' AND [local].[gp] = " & grLocal & ") AND ([emploi].[date1] = #" & CDate(dteDate.Value.ToString("MM/dd/yyyy")) & "# AND [emploi].gp = [local].gp)" ' & grLocal & ")"
            Requette(req)
        Next

        'Regler le probleme de l'heure lors de l'affichage
        Dim dtHeure As New DataTable
        Dim val As String

        dtHeure = dtt.Copy()

        Dim i As Integer = 0
        Dim j As Integer = 0

        For Each row As DataRow In dtt.Rows

            For j = 0 To 4

                val = row(j).ToString
                If (j = 3 Or j = 4) Then
                    dtHeure(i)(j) = TimeValue(val).ToShortTimeString
                Else
                    dtHeure(i)(j) = val
                End If

            Next

            i = i + 1
        Next

        dtgvEtatSalleNo.DataSource = dtHeure

        If (dtt.Rows.Count = 0) Then
            dtgvEtatSalleNo.Visible = False
            lblAfficheNo.Text = lblAfficheNo.Text + "VIDE"
        End If

    End Sub

    Private Sub btnMinEtatSalle_Click(sender As Object, e As EventArgs) Handles btnMinEtatSalle.Click

        MMinFenetre.MinFenetre()
    End Sub

    Private Sub btnMaxEtatSalle_Click(sender As Object, e As EventArgs) Handles btnMaxEtatSalle.Click

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub btnExitEtatSalle_Click(sender As Object, e As EventArgs) Handles btnExitEtatSalle.Click

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub btnMinEtatSalleNo_Click(sender As Object, e As EventArgs) Handles btnMinEtatSalleNo.Click

        MMinFenetre.MinFenetre()
    End Sub

    Private Sub btnMaxEtatSalleNo_Click(sender As Object, e As EventArgs) Handles btnMaxEtatSalleNo.Click

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub btnExitEtatSalleNo_Click(sender As Object, e As EventArgs) Handles btnExitEtatSalleNo.Click

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub btnMinEtatSalleYes_Click(sender As Object, e As EventArgs) Handles btnMinEtatSalleYes.Click

        MMinFenetre.MinFenetre()
    End Sub

    Private Sub btnMaxEtatSalleYes_Click(sender As Object, e As EventArgs) Handles btnMaxEtatSalleYes.Click

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub btnExitEtatSalleYes_Click(sender As Object, e As EventArgs) Handles btnExitEtatSalleYes.Click

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub btnParticulier_Click(sender As Object, e As EventArgs) Handles btnParticulier.Click

        pnlIfGeneral.Visible = True
        btnGeneral.Enabled = False


        Dim sauv As String = dteYear.Value
        dteYear.Value = CDate(sauv)
    End Sub

    Private Sub btnGeneral_Click(sender As Object, e As EventArgs) Handles btnGeneral.Click

        If MsgBox("Si vous continuez l'ancienne modification dans la table 'CrystalReport' de la BDD sera effacee", 36, "Annuler") = MsgBoxResult.Yes Then

            Dim str1 As String
            str1 = "Delete from [CrystalReport]"
            Dim cmd As OleDbCommand = New OleDbCommand(str1, connAccess)
            Try
                connAccess.Open()
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connAccess.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            GoTo fin2
        End If

        Dim matricule As String

        dtt.Rows.Clear()
        dtt.Columns.Clear()

        DtMatiere.Rows.Clear()
        DtMatiere.Columns.Clear()

        Dim req As String = "SELECT DISTINCT [affect].[matricule] FROM affect"
        DaMatiere = New OleDbDataAdapter(req, connAccess)
        DaMatiere.Fill(DtMatiere)

        Dim da As New OleDbDataAdapter
        Dim dt As New DataTable

        dt.Rows.Clear()
        dt.Columns.Clear()

        Dim reqRecup As String
        Dim numeroGp As UInteger
        Dim req1 As String

        'Generer la table 'dtt' contenant les informations necessaires obtenu en faisant des requettes.

        For Each dtr1 As DataRow In DtMatiere.Rows
            matricule = dtr1(0)

            reqRecup = "SELECT [affect].[gp] FROM affect WHERE [affect].[matricule] = '" + matricule + "'"

            da = New OleDbDataAdapter(reqRecup, connAccess)
            da.Fill(dt)

            For Each dtr2 As DataRow In dt.Rows
                numeroGp = dtr2(0)

                req1 = "SELECT [ETUDIANTS].[Matricule], [NomEtud], [Prenoms], [Sect], [Gr], [affect].[local], [affect].[pos], [emploi].[module], [emploi].[date1], [emploi].[debut], [emploi].[fin] FROM ETUDIANTS, affect, emploi WHERE ETUDIANTS.Matricule = '" + matricule + "' AND ( affect.matricule = '" + matricule + "' AND affect.gp = " & numeroGp & ") AND emploi.gp = " & numeroGp

                Requette(req1)
            Next

            dt.Rows.Clear()
            dt.Columns.Clear()
        Next

        'Remplir la table 'CrystalReport' de la base de donnees, a partir de la table 'dtt'.
        FillAccessTable()

        'Affichage et sauvegarde des donnnees sous Crystal Reports.
        AffichCrystalReport()

fin2:
    End Sub

    Private Sub btnCrystalReports_Click(sender As Object, e As EventArgs) Handles btnCrystalReports.Click

        'Effacer le contenu de la table CrystalReport de la BDD

        If MsgBox("Si vous continuez l'ancienne modification dans la table 'CrystalReport' de la BDD sera effacee", 36, "Annuler") = MsgBoxResult.Yes Then

            Dim str As String
            str = "Delete from [CrystalReport]"
            Dim cmd As OleDbCommand = New OleDbCommand(str, connAccess)
            Try
                connAccess.Open()
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connAccess.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            GoTo Arret
        End If

        pnlEtatSortie.Visible = False

        'Generer la table 'dtt' contenant les informations necessaires obtenu en faisant des requettes.
        AffichEtudiant(cmbEtudiant.SelectedValue)

        'Remplir la table 'CrystalReport' de la base de donnees, a partir de la table 'dtt'.
        FillAccessTable()

        'Affichage et sauvegarde des donnnees sous Crystal Reports.
        AffichCrystalReport()


        GoTo fin
Arret:
        If MsgBox("Voulez vous donc sauvegarder l'ancienne modification. Si vous cliquer sur 'OUI' vous serez deriger vers Crystal Reports pour la sauvegarde !", 36, "Annuler") = MsgBoxResult.Yes Then

            AffichCrystalReport()
        End If
fin:
    End Sub

    Private Sub btnSalleSuite_Click(sender As Object, e As EventArgs) Handles btnSalleSuite.Click

        moduleChoisit = cmbMatiereSalle.SelectedValue

        pnlSalleSuite.Visible = True
        pnlEtatSortie.Visible = False

        Dim daa11 As OleDbDataAdapter
        Dim Dtt11 As New DataTable

        Dim reqRecup As String = "SELECT [emploi].[gp] FROM [emploi] WHERE [emploi].[module] = '" + moduleChoisit + "'"
        daa11 = New OleDbDataAdapter(reqRecup, connAccess)
        daa11.Fill(Dtt11)

        Dim numeroGp1 As Integer = Dtt11(0)(0)

        Dim daS As OleDbDataAdapter
        Dim dtS As New DataTable

        Dim req6 As String = "SELECT DISTINCT [affect].[local] FROM [affect] WHERE [affect].[gp] = " & numeroGp1
        daS = New OleDbDataAdapter(req6, connAccess)
        daS.Fill(dtS)

        cmbSalleSuite.DataSource = dtS
    End Sub

    Private Sub btnCrstlSuite_Click(sender As Object, e As EventArgs) Handles btnCrstlSuite.Click

        'Effacer le contenu de la table CrystalReport de la BDD

        If MsgBox("Si vous continuez l'ancienne modification dans la table 'CrystalReport' de la BDD sera effacee", 36, "Annuler") = MsgBoxResult.Yes Then

            Dim str1 As String
            str1 = "Delete from [CrystalReport]"
            Dim cmd As OleDbCommand = New OleDbCommand(str1, connAccess)
            Try
                connAccess.Open()
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connAccess.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            GoTo fin1
        End If

        Dim daa1 As OleDbDataAdapter
        Dim daa2 As OleDbDataAdapter

        Dim dtt1 As New DataTable
        Dim dtt2 As New DataTable

        Dim matricule As String
        Dim req As String

        Dim reqRecup As String = "SELECT [emploi].[date1], [emploi].[gp] FROM [emploi] WHERE [emploi].[module] = '" + moduleChoisit + "'"
        daa1 = New OleDbDataAdapter(reqRecup, connAccess)
        daa1.Fill(dtt1)

        Dim dateExamen As String = dtt1(0)(0)
        Dim numeroGp As Integer = dtt1(0)(1)

        dtt.Rows.Clear()
        dtt.Columns.Clear()

        Dim req5 As String = "SELECT [affect].[matricule] FROM [affect] WHERE [affect].[gp] = " & numeroGp
        daa2 = New OleDbDataAdapter(req5, connAccess)
        daa2.Fill(dtt2)

        For Each dtr1 As DataRow In dtt2.Rows

            matricule = dtr1(0)
            req = "SELECT [ETUDIANTS].[Matricule], [NomEtud], [Prenoms], [Sect], [Gr], [affect].[local], [affect].[pos] FROM [ETUDIANTS], [affect] WHERE [ETUDIANTS].[Matricule] = '" + matricule + "' AND ( [affect].[matricule] = '" + matricule + "' AND [affect].[gp] = " & numeroGp & ")"

            Requette(req)
        Next

        'Remplir la table 'CrystalReport' de la base de donnees, a partir de la table 'dtt'.

        Dim str As String
        Dim i As Integer
        Dim cpt As Integer = 1

        str = "Insert into CrystalReport([N], [Matricule], [Nom], [Prenom], [Section], [Gr], [Local], [Position]) values (?,?,?,?,?,?,?,?)"

        For i = 0 To dtt.Rows.Count - 1

            Dim cmd As OleDbCommand = New OleDbCommand(str, connAccess)

            connAccess.Close()

            'Toujours on verifie si notre chaine de caractere n'est pas vide, car il se peut qu'un champ de la bDD est vide !

            cmd.Parameters.Add(New OleDbParameter("N", CType(cpt, Integer)))
            If (dtt(i)(0).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Matricule", CType(dtt(i)(0), String)))
            End If
            If (dtt(i)(1).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Nom", CType(dtt(i)(1), String)))
            End If
            If (dtt(i)(2).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Prenom", CType(dtt(i)(2), String)))
            End If
            If (dtt(i)(3).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Section", CType(dtt(i)(3), String)))
            End If
            If (dtt(i)(4).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Gr", CType(dtt(i)(4), String)))
            End If
            If (dtt(i)(5).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Local", CType(dtt(i)(5), String)))
            End If
            If (IsNumeric(dtt(i)(6))) Then
                cmd.Parameters.Add(New OleDbParameter("Position", CType(dtt(i)(6), Integer)))
            End If

            cpt = cpt + 1

            Try
                connAccess.Open()
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connAccess.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Next

        Dim report2 As New CrystalReport2

        Form2.Show()
        dtt.Rows.Clear()
        dtt.Columns.Clear()

        'report2.SetDataSource()

        Dim req3 As String = "SELECT [Matricule], [Nom], [Prenom], [Section], [Gr], [Position] FROM CrystalReport WHERE [CrystalReport].[local] = '" + cmbSalleSuite.SelectedValue + "'"

        Requette(req3)
        report2.SetDataSource(dtt)
        Form2.CrystalReportViewer1.ReportSource = report2
        Form2.CrystalReportViewer1.Refresh()
        Form2.CrystalReportViewer1.RefreshReport()

        dtt.Rows.Clear()
        dtt.Columns.Clear()
fin1:
    End Sub

    Private Sub btnRetourner_Click(sender As Object, e As EventArgs) Handles btnRetourner.Click

        pnlEtatSortie.Visible = True
        pnlSalleSuite.Visible = False
    End Sub
    Function test_bdd() As Integer
        Dim chemin As String
        bdd = cmbRest1Annee.Text.ToString + ".accdb"

        chemin = "..\..\Sauvegarde\" + bdd

        Dim file As File
        If File.Exists(chemin) Then
            MsgBox("il ya une affectation dans cette année")
            Return 1
            Exit Function
        Else
            MsgBox("il y'a pas une affectation dans cette année ")
            Return 0
            Exit Function
        End If
    End Function

    Private Sub btnExitEtatSortie_Click_1(sender As Object, e As EventArgs) Handles btnExitEtatSortie.Click

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub btnMaxEtatSortie_Click_1(sender As Object, e As EventArgs) Handles btnMaxEtatSortie.Click

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub btnMinEtatSortie_Click_1(sender As Object, e As EventArgs) Handles btnMinEtatSortie.Click

        MMinFenetre.MinFenetre()
    End Sub

    Private Sub btnRest1Annee_Click(sender As Object, e As EventArgs)

        annee = cmbRest1Annee.Text
        FillPromo()
    End Sub


    Function GetTable(ByVal selectCommand As String) As DataTable
        Dim tbl As New DataTable
        annee = anne.Text + ".accdb"
        cont = New OleDbConnection("Provider=Microsoft.ace.OleDb.12.0;DATA SOURCE = ..\..\Sauvegarde\" + annee)
        cmd = New OleDbCommand("", cont)
        Try

            If cont.State = ConnectionState.Closed Then cont.Open()
            cmd.CommandText = selectCommand
            tbl.Load(cmd.ExecuteReader())

        Catch ex As Exception
        Finally
            If cont.State = ConnectionState.Open Then cont.Close()
        End Try
        Return tbl
    End Function

    Private Sub btnEntrerRetauration1_Click(sender As Object, e As EventArgs) Handles btnEntrerRetauration1.Click

        pnlRestauration1.Visible = False
        pnlRestauration2.Visible = True
        Dim modu As String

        promo = cmbRes1Promo.Text
        sem = cmbRes1Semestre.Text
        c = cmbRest1TpeExam.Text
        modu = cmbRest1Module.Text
        If cont.State = ConnectionState.Closed Then cont.Open()
        dgvRest2.DataSource = GetTable("SELECT [affect].[matricule] , [ETUDIANTS].[NomEtud] ,[ETUDIANTS].[Prenoms],[ETUDIANTS].[Sect],[ETUDIANTS].[Gr] FROM [ETUDIANTS] , [affect] ,[emploi] ,[exam] WHERE [ETUDIANTS].[Matricule]=[affect].[matricule] AND  [emploi].[gp]=[affect].[gp] AND [ETUDIANTS].[Promo]='" + promo + "' AND [emploi].[module]='" + modu + "' AND [exam].[semestre]='" + sem + "'  AND  [exam].[gp]=[affect].[gp] AND [exam].[type_exam]='" + c + "'")
        If cont.State = ConnectionState.Open Then cont.Close()
    End Sub

    Private Sub grpBoxRest2_Enter(sender As Object, e As EventArgs) Handles grpBoxRest2.Enter

        If rdbtnMatriculeRest2.Checked = True Then
            cmbRest2Matricule.Visible = True
        End If
    End Sub

    Private Sub btnRest2Chercher_Click(sender As Object, e As EventArgs) Handles btnRest2Chercher.Click

        Dim ser As String
        c = cmbRest1TpeExam.Text
        promo = cmbRes1Promo.Text
        sem = cmbRes1Semestre.Text
        ser = "SELECT  [affect].[matricule] ,[ETUDIANTS].[NomEtud],[ETUDIANTS].[Prenoms] ,[ETUDIANTS].[Promo],[ETUDIANTS].[Sect],[ETUDIANTS].[Gr],[emploi].[module],[affect].[local],[affect].[pos]  FROM [ETUDIANTS]  ,[affect ],[emploi] ,[MODULES] ,[exam] WHERE [ETUDIANTS].[Matricule] =[affect].[matricule] and [affect].[gp]=[emploi].[gp] AND [MODULES].[Code_Mat]=[emploi].[module] and [exam].[semestre]='" + sem + "' AND [MODULES].[Niveau]='" + promo + "' and  [exam].[type_exam]='" + c + "'   AND  [exam].[gp]=[affect].[gp] AND  "
        If rdbtnMatriculeRest2.Checked Then
            ser += "[ETUDIANTS].[Matricule] =   '" + cmbRest2Matricule.Text + "'"
        End If
        dgvRest2.DataSource = GetTable(ser)
    End Sub

    Private Sub btnRest2Precedent_Click(sender As Object, e As EventArgs) Handles btnRest2Precedent.Click

        pnlRestauration2.Visible = False
        pnlRestauration1.Visible = True
        pnlEtatSortie.Visible = False
        pnlEtatSalle.Visible = False
        pnlSalleSuite.Visible = False
        pnlEtatSalleYes.Visible = False
        pnlMenu.Visible = False
        pnlLogin.Visible = False
        pnlAffectation1.Visible = False
        pnlAffectation2.Visible = False
        pnlAffectation3.Visible = False
        pnlEtatSalleNo.Visible = False
    End Sub

    Private Sub btnEtatSortie_Click(sender As Object, e As EventArgs) Handles btnEtatSortie.Click

        pnlRestauration2.Visible = False
        pnlRestauration1.Visible = False
        pnlEtatSortie.Visible = True
        pnlEtatSalle.Visible = False
        pnlSalleSuite.Visible = False
        pnlEtatSalleYes.Visible = False
        pnlMenu.Visible = False
        pnlLogin.Visible = False
        pnlAffectation1.Visible = False
        pnlAffectation2.Visible = False
        pnlAffectation3.Visible = False
        pnlEtatSalleNo.Visible = False

        pnlLeftSide1.Visible = False
        pnlLeftSide2.Visible = False
        pnlLeftSide3.Visible = False
        pnlLeftSide4.Visible = True
        pnlLeftSide5.Visible = False
        pnlLeftSide6.Visible = False

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles btnMinRestauration.Click

        MMinFenetre.MinFenetre()
    End Sub

    Private Sub btnMaxRestauration_Click(sender As Object, e As EventArgs) Handles btnMaxRestauration.Click

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles btnExitRestauration.Click

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles btnRest2Min.Click

        MMinFenetre.MinFenetre()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles btnRest2Max.Click

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles btnRest2Exit.Click

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles btnSSuiteMin.Click

        MMinFenetre.MinFenetre()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles btnSSuiteMax.Click

        MMaxFenetre.MaxFenetre()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles btnSSuiteExit.Click

        MExitFenetre.ExitFenetre()
    End Sub

    Private Sub btnAcceuil_Click(sender As Object, e As EventArgs) Handles btnAcceuil.Click

        pnlToolbarAffect.Visible = False
        pnlToolbarMenu.Visible = True

        pnlRestauration2.Visible = False
        pnlRestauration1.Visible = False
        pnlEtatSortie.Visible = False
        pnlEtatSalle.Visible = False
        pnlSalleSuite.Visible = False
        pnlEtatSalleYes.Visible = False
        pnlMenu.Visible = True
        pnlLogin.Visible = False
        pnlAffectation1.Visible = False
        pnlAffectation2.Visible = False
        pnlAffectation3.Visible = False
        pnlEtatSalleNo.Visible = False

        pnlLeftSide1.Visible = False
        pnlLeftSide2.Visible = False
        pnlLeftSide3.Visible = False
        pnlLeftSide4.Visible = False
        pnlLeftSide5.Visible = False
        pnlLeftSide6.Visible = False
    End Sub


    '      AFFECTATION AUTOMATIQUE - MADJI WALID -
    '*********************************************************

    Dim con As New OleDbConnection
    Dim arr As New ArrayList
    Dim dt As New DataTable
    Dim dt2 As New DataTable
    Dim maxgp As Integer = 0
    Dim da As OleDbDataAdapter
    Dim cmd As OleDbCommand
    Dim dt0 As New DataTable
    Dim dt00 As New DataTable
    Dim dt3 As DataTable
    Dim dt4 As New DataTable
    Dim dt6 As New DataTable
    Dim pos As Integer = 0
    Dim path As String
    Dim path2 As String
    Dim nameBdd As String

    '//////////////////////////////////////////////////////////////////////////////////////////
    Function executer_cmd(ByVal comd As String) As Integer ''''''''''''''
        Dim cmd As New OleDbCommand(comd, con)
        Try
            con.Open()
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "attention")
            Return 0
            Exit Function
        Finally
            con.Close()
        End Try
        Return 1
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////
    Sub headers_dgvsalles() '''''''''''''''''
        dgvsalles.DataSource = Nothing
        dgvsalles.Columns.Add("a", "MODULEs")
        dgvsalles.Columns.Add("b", "salle")
        dgvsalles.Columns.Add("c", "max")
        dgvsalles.Columns.Add("d", "gp")
        dgvsalles.Columns(3).Visible = False
    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////
    Sub filldgv2() ''''''''''''
        dt2.Columns.Clear()
        dt2.Rows.Clear()
        dt2.Columns.Add("MODULEs")
        dt2.Columns.Add("salle")
        dt2.Columns.Add("max")
        dt2.Columns.Add("gp")
        'dgvsalles.DataSource = dt2
        'dgvsalles.Columns(3).Visible = False

    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////
    Sub remplir_dgvsalles() ''''''''''''''
        dgvsalles.DataSource = Nothing
        For i As Integer = 0 To dgvsalles.Rows.Count - 1
            dgvsalles.Rows.RemoveAt(0)
        Next

        Dim ro() As DataRow = dt2.Select("gp ='" + gp.Text + "'")

        For i As Integer = 0 To ro.Count - 1
            dgvsalles.Rows.Add()
            dgvsalles.Rows(i).Cells(0).Value = ro(i)(0).ToString
            dgvsalles.Rows(i).Cells(1).Value = ro(i)(1).ToString
            dgvsalles.Rows(i).Cells(2).Value = ro(i)(2).ToString
            dgvsalles.Rows(i).Cells(3).Value = ro(i)(3).ToString

        Next
    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////
    Function remplir_dt(ByVal comd As String, ByVal dt As DataTable) '''''''''''''''''''
        dt.Columns.Clear()
        dt.Rows.Clear()

        Dim da As New OleDbDataAdapter(comd, con)
        Try
            da.Fill(dt)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "attention")
            Return 0
            Exit Function
        End Try
        Return 1
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////
    Sub filldgv() ''''''''''''''''''''''''
        dt.Columns.Clear()
        dt.Rows.Clear()
        dt.Columns.Add("MODULE")
        dt.Columns.Add("date")
        dt.Columns.Add("debut")
        dt.Columns.Add("fin")
        dgvmodules.DataSource = dt

    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////
    Sub net() '''''''''''''''''
        record.Text = String.Empty
        record.Items.Clear()

    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////
    Sub fill_groupes() ''''''''''''''''''''
        dt4.Columns.Clear()
        dt4.Rows.Clear()
        dt4.Columns.Add("MODULE")
        dt4.Columns.Add("gp")
    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////
    Sub add_groupe(ByVal record As String, ByVal gp As String) '''''''''''''''''''

        Dim row As DataRow = dt4.NewRow
        row(0) = record
        row(1) = gp
        dt4.Rows.Add(row)
    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////
    Sub update_emploi() ''''''''''''''''''''''''''''''''''''''
        For Each row In dt4.Rows


            Try
                cmd = New OleDbCommand("update emploi set gp =@gp where emploi.module=@module and emploi.gp= 0 ", con)
                With cmd.Parameters
                    .AddWithValue("@gp", row(1)).DbType = DbType.Int32
                    .AddWithValue("@module", row(0)).DbType = DbType.String
                End With
                con.Open()
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "attention")
                Exit Sub
            Finally
                con.Close()

            End Try

        Next
    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles parcourir.Click '''''''''''''''''
        'selectionner la bdd de l'utilisateur
        With OpenFileDialog1
            .AddExtension = True
            .CheckFileExists = True
            .CheckPathExists = True
            .InitialDirectory = "C://"
            .FileName = ""
            .Title = "choisir une base de donnee!!!"
            .Filter = "Microsoft Access |*.accdb|Microsoft Access Databases|*.accdb"
        End With
        'copier la bdd dans l'espace de travail

        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            path = OpenFileDialog1.FileName
            defultpath.Text = path
        End If


    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////
    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click ''''''''''''

        If (tester_view1() = 0) Then
            Exit Sub
        End If
        If (testerbdd_et_con() = 0) Then
            Exit Sub
        End If
        allerde1a2()
        getting_modules(nompromo.Text, semestre.Text)
        If (nompromo.Text = "1CP") And (semestre.Text <> "anu") Then
            ajoutBW.Visible = True
            If (semestre.Text = "S1") Then
                ajoutBW.Text = "ajouter eng"
            Else
                ajoutBW.Text = "ajouter bweb"
            End If
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles addmodule.Click '''''''''''''''''
        'test des champs

        newgp.Visible = True
        If (addmodule.Text = "modifier") Then
            record.Enabled = True
            addmodule.Text = "add"
        End If
        If (tester_existance_module() = 0) Then
            Exit Sub
        End If

        For Each row In dt.Rows
            If Not ((record.Text = "BW" And row(0).ToString = "ANG1") Or (record.Text = "ANG1" And row(0).ToString = "BW")) Then
                If (CDate(dateexam.Text) = CDate(row(1))) And ((TimeValue(row(2)) >= TimeValue(debut.Value) And TimeValue(row(2)) <= TimeValue(fin.Value)) Or (TimeValue(row(3)) >= TimeValue(debut.Value) And TimeValue(row(3)) <= TimeValue(fin.Value))) Then
                    MsgBox("attention, il ya deja un examen de " & row(0) & " qui ce deroule pendant la periode choisie", MsgBoxStyle.Critical, "periode occupée")
                    debut.Focus()
                    Exit Sub
                End If
            End If
        Next
        If (TimeValue(debut.Value) >= TimeValue(fin.Value)) Then
            MsgBox("attention,l'heure de debut ne peut pas etre superieure a l'heure de la fin", MsgBoxStyle.Critical, "horaire")
            Exit Sub
        End If
        If (view3.Visible = True) Then
            modules.Items.Add(record.Text)
            ajouter_module()
            cmd = New OleDbCommand("insert into emploi values(@module,@date1,@debut,@fin,@gp)", con)
            Dim i As Integer = dgvmodules.Rows.Count - 1
            With cmd.Parameters
                .AddWithValue("@module", dgvmodules.Rows(i).Cells(0).Value).DbType = DbType.String
                .AddWithValue("@date1", dgvmodules.Rows(i).Cells(1).Value).DbType = DbType.Date
                .AddWithValue("@debut", dgvmodules.Rows(i).Cells(2).Value).DbType = DbType.String
                .AddWithValue("@fin", dgvmodules.Rows(i).Cells(3).Value).DbType = DbType.String

                .AddWithValue("@gp", 0).DbType = DbType.Int32
            End With
            Try
                con.Open()
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "attention ")
                Exit Sub
            Finally
                con.Close()

            End Try
        Else
            ajouter_module()
        End If
        If (view3.Visible = False) Then

            allerausalles.Enabled = True
        End If
    End Sub

    Function tester_existance_module() As Integer ''''''''''''''''''
        If (record.Text = String.Empty) Then
            MsgBox("il faut choisir un module", MsgBoxStyle.Critical, "attention")
            record.Focus()
            Return 0
            Exit Function
        Else
            For i As Integer = 0 To dgvmodules.RowCount - 1

                If (dgvmodules.Rows(i).Cells(0).Value = record.Text.ToString) Or (dgvmodules.Rows(i).Cells(1).Value = dateexam.Value And dgvmodules.Rows(i).Cells(3).Value > debut.Value) Then
                    MsgBox("le module existe deja ou erreur de l'horraire du module", MsgBoxStyle.Critical, "attention")
                    record.Focus()
                    Return 0
                    Exit Function
                End If

            Next

        End If
        Return 1
    End Function

    Sub ajouter_module() ''''''''''
        Dim ro As DataRow = dt.NewRow
        ro(0) = record.Text.ToString
        ro(1) = dateexam.Text
        ro(2) = debut.Text
        ro(3) = fin.Text

        dt.Rows.Add(ro)
        dgvmodules.DataSource = dt
        record.Text = String.Empty
    End Sub

    Sub net2() '''''''''''''
        For i As Integer = 0 To modules.Items.Count - 1
            modules.SetItemCheckState(i, CheckState.Unchecked)
            local.Text = String.Empty
            max.ResetText()

        Next
    End Sub

    Function getlocalaux() As Integer ''''''''''''''''

        Dim str As String = "select * from salles "
        If (remplir_dt(str, dt6) = 0) Then
            Return 0
            Exit Function
        End If
        For i As Integer = 0 To dt6.Rows.Count - 1
            local.Items.Add(dt6.Rows(i)(0))
        Next
        Return 1
    End Function

    Sub remplir_emploi() ''''''''''''''''''''''''''
        For i As Integer = 0 To dgvmodules.Rows.Count - 1
            cmd = New OleDbCommand("insert into emploi values(@module,@date1,@debut,@fin,@gp)", con)
            With cmd.Parameters
                .AddWithValue("@module", dgvmodules.Rows(i).Cells(0).Value).DbType = DbType.String
                .AddWithValue("@date1", dgvmodules.Rows(i).Cells(1).Value).DbType = DbType.Date
                .AddWithValue("@debut", dgvmodules.Rows(i).Cells(2).Value).DbType = DbType.String
                .AddWithValue("@fin", dgvmodules.Rows(i).Cells(3).Value).DbType = DbType.String

                .AddWithValue("@gp", 0).DbType = DbType.Int32
            End With
            Try
                con.Open()
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "attention ")
                Exit Sub
            Finally
                con.Close()

            End Try
        Next
    End Sub

    Sub aller2a3() ''''''''''''''''''''
        allerausalles.Enabled = False
        update.Enabled = False
        modules.Enabled = True
        view3.Visible = True
        view4.Visible = False
        local.Text = String.Empty
        max.Value = 0
        For K As Integer = 0 To dgvmodules.Rows.Count - 1
            modules.Items.Add(dgvmodules.Rows(K).Cells(0).Value.ToString)
        Next
        dgvsalles.Visible = True
        fill_groupes()
        getlocalaux()

        filldgv2()

    End Sub

    Private Sub ajoutBW_Click(sender As Object, e As EventArgs) Handles ajoutBW.Click
        If (ajoutBW.Text = "ajouter eng") Or (ajoutBW.Text = "ajouter bweb") Then
            PANEL1.Visible = True
            If (view3.Visible = True) Then
                For i As Integer = 0 To dgvmodules.Rows.Count - 1
                    If (dgvmodules.Rows(i).Cells(0).Value = "ANG1") Or (dgvmodules.Rows(i).Cells(0).Value = "BW") Then
                        dgvmodules.Rows(i).Selected = True
                        supp_mod_affect(i)
                        If (ajoutBW.Text = "ajouter eng") Then
                            modules.Items.Add("BW")
                        Else
                            modules.Items.Add("ANG1")
                        End If
                        Exit Sub
                    End If

                Next
            End If


        Else
            Dim cond As String
            Dim mod_aff As String
            If (ajoutBW.Text = "supprimer bweb") Then
                cond = "BW"
                mod_aff = "ANG1"
                ajoutBW.Text = "ajouter bweb"
            Else
                mod_aff = "BW"
                cond = "ANG1"
                ajoutBW.Text = "ajouter eng"
            End If
            If MsgBox("voulez vous supprimer " & cond & " ?", MsgBoxStyle.YesNo, "confirmation") = MsgBoxResult.Yes Then
                record.Items.Remove(cond)
                For i As Integer = 0 To dgvmodules.Rows.Count - 1
                    If (i > dgvmodules.Rows.Count - 1) Then  'a cause du supression 
                        Exit Sub
                    End If
                    If (dgvmodules.Rows(i).Cells(0).Value = cond.ToString) Then
                        dgvmodules.Rows(i).Selected = True
                        suppr_mod(i)
                    ElseIf (dgvmodules.Rows(i).Cells(0).Value = mod_aff.ToString) Then
                        dgvmodules.Rows(i).Selected = True
                        supp_mod_affect(i)
                        modules.Items.Add(mod_aff)
                    End If
                Next

            End If
        End If
    End Sub

    Private Sub BOK_Click(sender As Object, e As EventArgs) Handles BOK.Click
        Dim modd As String
        If (sections.CheckedItems.Count = 0) Then
            MsgBox("attention il faut choisir au moins une section")
        Else
            PANEL1.Visible = False
            If (ajoutBW.Text = "ajouter bweb") Then
                ajoutBW.Text = "supprimer bweb"
                modd = "BW"

            Else
                ajoutBW.Text = "supprimer eng"
                modd = "ANG1"
            End If
            record.Items.Add(modd)
        End If
    End Sub

    Private Sub banu_Click(sender As Object, e As EventArgs) Handles banu.Click
        sections.Items.Clear()
        sections.Items.Add("A")
        sections.Items.Add("B")
        sections.Items.Add("C")
        sections.Items.Add("D")
        PANEL1.Visible = False
    End Sub

    Private Sub returntoview2_Click(sender As Object, e As EventArgs) Handles returntoview2.Click
        If MsgBox("vous aller perdre les informations pour cette promo, continue?", MsgBoxStyle.YesNo, "confirmation") = MsgBoxResult.Yes Then


            Dim prec_promo As String
            For i As Integer = 0 To listepromo.CheckedItems.Count - 1
                If (listepromo.CheckedItems(i).ToString = nompromo.Text) Then
                    If (i > 0) Then
                        prec_promo = listepromo.CheckedItems(i - 1).ToString
                        Exit For
                    End If
                End If

            Next
            If (prec_promo = String.Empty) Then
                pnlAffectation2.Visible = False
                pnlAffectation1.Visible = True
                reinitialiser_var_glob()
                net()
                reremplir_liste_promo()
            Else
                reinitialiser_var_glob()
                remplir_var_glob(prec_promo)

            End If
        End If
    End Sub

    Sub reremplir_liste_promo()
        listepromo.Items.Clear()
        listepromo.Items.Add("1CP")
        listepromo.Items.Add("2CP")
        listepromo.Items.Add("1CS")
        listepromo.Items.Add("2ST")
        listepromo.Items.Add("2SL")
        listepromo.Items.Add("2SQ")
    End Sub

    Private Sub update_Click(sender As Object, e As EventArgs) Handles update.Click
        Dim ind As Integer
        If (dgvmodules.SelectedRows.Count > 0) Then
            ind = dgvmodules.SelectedRows.Item(0).Index
            record.Text = dt.Rows(ind)(0).ToString
            dateexam.Text = dt.Rows(ind)(1).ToString
            debut.Text = dt.Rows(ind)(2).ToString
            fin.Text = dt.Rows(ind)(3).ToString
            dt.Rows.RemoveAt(dgvmodules.SelectedRows.Item(0).Index)
            addmodule.Text = "modifier"
            record.Enabled = False

        End If
    End Sub

    Private Sub delete_Click(sender As Object, e As EventArgs) Handles delete.Click

        If dgvmodules.SelectedRows.Count <= 0 Then
            Exit Sub
        End If
        If MsgBox("voulez vous supprimer cette ligne?", MsgBoxStyle.YesNo, "confirmation") = MsgBoxResult.Yes Then
            For i As Integer = 0 To dgvmodules.SelectedRows.Count - 1
                suppr_mod(i)
            Next
        End If
    End Sub

    Sub supp_mod_affect(ByVal i As Integer)
        If (view3.Visible = True) Then
            Dim name = dt.Rows(dgvmodules.SelectedRows.Item(0).Index)(0).ToString
            If (modules.Items.Contains(name)) Then
                modules.Items.Remove(name)

            End If
            For j As Integer = 0 To dt4.Rows.Count - 1
                If dt4.Rows(j)(0).ToString = name Then

                    dt4.Rows.RemoveAt(j)
                    Exit For

                End If

            Next


            Dim newrow As String = Nothing

            For Each row In dt2.Rows
                'dt2.Rows.Remove(row)
                newrow = String.Empty


                For Each ele In row(0).ToString.Split("/")
                    newrow &= ele & "/"

                    If (ele = name) Then
                        newrow = newrow.Remove(newrow.Length - ele.Length - 1)


                    End If

                Next
                newrow = newrow.Remove(newrow.Length - 1)


                row(0) = newrow


            Next
            Dim ro() As DataRow = dt2.Select("MODULEs=''")
            If (ro.Count > 0) Then
                Dim f As Integer = CType(ro(0)(3).ToString, Int32)

                If (arr.Contains(f)) Then
                    arr.Remove(f)


                End If


            End If
            For Each row In ro

                dt2.Rows.Remove(row)
            Next
            'update_emploi
            remplir_dgvsalles()
            cmd = New OleDbCommand("delete emploi.* from [emploi] inner join [MODULES] on emploi.module=MODULES.Code_Mat where [MODULES.Niveau]='" + nompromo.Text.ToString + "'", con)
            Try
                con.Open()
                cmd.ExecuteNonQuery()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "erreur de connexion")
                Exit Sub
            Finally
                con.Close()
            End Try
            If (modules.CheckedItems.Count <= 0) Then
                view4.Visible = False
            End If
            If (dgvsalles.Rows.Count <= 0) Then
                If (arr.Count > 0) Then
                    Dim k As Integer = 0
                    While (k < modules.Items.Count)

                        If (modules.GetItemCheckState(k) = CheckState.Checked) Then

                            modules.Items.RemoveAt(k)
                            k = k - 1
                        End If
                        k = k + 1
                    End While

                    Dim nxt As Integer = arr.Item(0)
                    gp.Text = nxt.ToString
                    remplir_dgvsalles()
                    'dgvsalles.DataSource = dt2


                    For Each item In dt4.Rows
                        If item(1) = gp.Text Then
                            modules.Items.Add(item(0))
                            modules.SetItemChecked(modules.Items.Count - 1, True)
                        End If

                    Next
                End If
            End If
            If (dt.Rows.Count <= 0) Then

                view4.Visible = False
                view3.Visible = False
                modules.Items.Clear()
                local.Items.Clear()
                arr.Clear()
                allerausalles.Enabled = True
            Else
                remplir_emploi()

            End If
        End If
    End Sub

    Sub suppr_mod(ByVal i As Integer)

        supp_mod_affect(i)
        dt.Rows.RemoveAt(dgvmodules.SelectedRows.Item(0).Index)


    End Sub

    Private Sub deleteall_Click(sender As Object, e As EventArgs) Handles deleteall.Click
        If MsgBox("voulez vous supprimer tout?", MsgBoxStyle.YesNo, "confirmation") = MsgBoxResult.Yes Then
            dt2.Rows.Clear()
            For i As Integer = 0 To dgvsalles.Rows.Count - 1
                dgvsalles.Rows.RemoveAt(0)
            Next
            modules.Items.Clear()
            local.Items.Clear()
            arr.Clear()
            view4.Visible = False
            view3.Visible = False
            allerausalles.Visible = True
            dt.Rows.Clear()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles allerausalles.Click ''''''**************''''''''
        If (dgvmodules.Rows.Count <= 0) Then
            MsgBox("l'emploi est vide, il faut au moins un module", MsgBoxStyle.Critical, "emploi vide !")
            Exit Sub
        End If
        Dim promo As String
        modules.Items.Clear()
        aller2a3()


        dgvsalles.Columns.Clear()
        dgvsalles.Rows.Clear()
        remplir_emploi()
        'remplissage de modules
        promo = nompromo.Text.ToString
        tri_crit(critere.Text, promo)
        headers_dgvsalles()

    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click ''''''''''''''''
        If (test_gp_modules() = 0) Then
            Exit Sub
        End If
        'update_emploi()
        alleraulocalaux()

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles addlocal.Click '''''''''''''
        If (addlocal.Text = "modifier") Then
            addlocal.Text = "add"
        End If
        If (modules.CheckedItems.Count <= 0) Then
            MsgBox("il faut choisir au moins un module", MsgBoxStyle.Critical, "aucun module!")
            Exit Sub
        End If
        If (test_gp_modules() = 0) Then
            Exit Sub
        End If
        If (tester_local() <> 0 And tester_local2() <> 0) Then


            Dim row As DataRow = dt2.NewRow
            If (dgvsalles.Rows.Count > 0) Then
                row(0) = dgvsalles.Rows(0).Cells(0).Value
            Else
                For Each item In modules.CheckedItems

                    row(0) += item + "/"

                Next
                row(0).ToString.Remove(row(0).ToString.Length - 1)
            End If

            row(1) = local.Text
            row(2) = max.Value.ToString
            row(3) = gp.Text
            dt2.Rows.Add(row)
            'dgvsalles.DataSource = dt2

            'remplir_dgvsalles()
            Dim i = dgvsalles.Rows.Add()
            dgvsalles.Rows(i).Cells(0).Value = row(0).ToString
            dgvsalles.Rows(i).Cells(1).Value = row(1).ToString
            dgvsalles.Rows(i).Cells(2).Value = row(2).ToString
            dgvsalles.Rows(i).Cells(3).Value = row(3).ToString



            nextpromo.Enabled = True
        End If

    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles newgp.Click ''''''''''
        If modules.Items.Count = modules.CheckedItems.Count Then
            MsgBox("pas de modules !")
            Exit Sub
        End If

        Dim k As Integer = 0
        If (dgvsalles.Rows.Count <= 0) Then
            MsgBox("ce groupe n'est pas correct, il y a aucun local associé", MsgBoxStyle.Critical, "pas de locaux!")
            Exit Sub
        End If
        If (tester_nb_places() = 0) Then
            Exit Sub
        End If

        modules.Enabled = True
        view4.Visible = False
        While (k < modules.Items.Count)

            If (modules.GetItemCheckState(k) = CheckState.Checked) Then
                add_groupe(modules.Items(k), gp.Text)
                modules.Items.RemoveAt(k)
                k = k - 1
            End If
            k = k + 1
        End While
        Dim f As Integer = CType(gp.Text, Int32)
        arr.Add(f)

        arr.Sort()
        Dim nxt As Integer = arr.Item(arr.Count - 1) + 1
        gp.Text = nxt.ToString

        If (modules.Items.Count <= 0) Then
            newgp.Enabled = False

        End If
        For i As Integer = 0 To dgvsalles.Rows.Count - 1
            dgvsalles.Rows.RemoveAt(0)
        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click '''''''''''''''''''''''

        Dim f As Integer = CType(gp.Text, Int32)

        If (arr.Count <= 0) Then
            MsgBox("pas de groupes !!")
            Exit Sub

        ElseIf (arr.Contains(f) And (arr.IndexOf(f) = 0)) Then
            MsgBox("pas de groupes !!")
            Exit Sub
        End If

        If (tester_nb_places() = 0) Then
            Exit Sub
        Else
            Dim k As Integer = 0
            While (k < modules.Items.Count)

                If (modules.GetItemCheckState(k) = CheckState.Checked) Then
                    If (Not (arr.Contains(f))) Then
                        add_groupe(modules.Items(k), gp.Text)
                    End If
                    modules.Items.RemoveAt(k)
                    k = k - 1
                End If
                k = k + 1
            End While

        End If
        Dim nxt As Integer
        If (arr.Contains(f)) Then
            nxt = arr.Item(arr.IndexOf(f) - 1)
        Else
            arr.Add(f)

            update_emploi()
            nxt = arr.Item(arr.Count - 2)
        End If
        gp.Text = nxt.ToString
        remplir_dgvsalles()
        'dgvsalles.DataSource = dt2
        Dim i As Integer = 0

        For Each item In dt4.Rows

            If item(1) = gp.Text Then
                modules.Items.Add(item(0))
                modules.SetItemChecked(modules.Items.Count - 1, True)
            End If

        Next

    End Sub

    Private Sub gpsuiv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gpsuiv.Click '''''''''''''''''''




        If (arr.Count <= 0) Then
            MsgBox("pas de groupes !!")
            Exit Sub
        End If

        Dim f As Integer = CType(gp.Text, Int32)
        If Not (arr.Contains(f)) Then
            MsgBox("pas de groupes !!")
            Exit Sub
        End If

        If arr.IndexOf(f) = arr.Count - 1 Then

            MsgBox("c'est le dernier groupe  !!")
            Exit Sub

        Else
            If (tester_nb_places() = 0) Then
                Exit Sub
            End If

            Dim k As Integer = 0
            While (k < modules.Items.Count)

                If (modules.GetItemCheckState(k) = CheckState.Checked) Then

                    modules.Items.RemoveAt(k)
                    k = k - 1
                End If
                k = k + 1
            End While
        End If
        Dim nxt As Integer = arr.Item(arr.IndexOf(f) + 1)

        gp.Text = nxt.ToString
        remplir_dgvsalles()
        'dgvsalles.DataSource = dt2



        For Each item In dt4.Rows

            If item(1) = gp.Text Then
                modules.Items.Add(item(0))
                modules.SetItemChecked(modules.Items.Count - 1, True)
            End If

        Next

    End Sub

    Private Sub supprimer2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles supprimer2.Click ''''''''''''''''''''
        If dgvsalles.SelectedRows.Count <= 0 Then
            Exit Sub
        End If
        If (dgvsalles.SelectedRows.Count <= 0) Then
            MsgBox("veuillez d'abord selectionner des ligne a supprimer")
        End If
        If MsgBox("voulez vous supprimer cette ligne?", MsgBoxStyle.YesNo, "confirmation") = MsgBoxResult.Yes Then
            For i As Integer = 0 To dgvsalles.SelectedRows.Count - 1
                Dim arra As New ArrayList
                For j As Integer = 0 To dt2.Rows.Count - 1
                    If (dt2.Rows(j)(1) = dgvsalles.SelectedRows.Item(0).Cells(1).Value) Then

                        'dt2.Rows.RemoveAt(j)
                        arra.Add(j)
                    End If
                Next
                Dim k As Integer = 0
                For Each pos As Integer In arra

                    dt2.Rows.RemoveAt(pos - k)
                    k += 1
                Next

                dgvsalles.Rows.RemoveAt(dgvsalles.SelectedRows.Item(0).Index)

            Next
        End If
    End Sub

    Private Sub supp_tt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles supp_tt.Click ''''''''''''''''''
        If dgvsalles.Rows.Count <= 0 Then
            Exit Sub
        End If
        If MsgBox("voulez vous supprimer tous les locaux pour ce groupe de modules", MsgBoxStyle.YesNo, "confirmation") = MsgBoxResult.Yes Then
            For i As Integer = 0 To dgvsalles.Rows.Count - 1
                dgvsalles.Rows.RemoveAt(0)
            Next

            Dim arra As New ArrayList
            For j As Integer = 0 To dt2.Rows.Count - 1
                If (dt2.Rows(j)(3) = gp.Text) Then

                    'dt2.Rows.RemoveAt(j)
                    arra.Add(j)
                End If
            Next
            Dim k As Integer = 0
            For Each pos As Integer In arra

                dt2.Rows.RemoveAt(pos - k)
                k += 1
            Next


        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nextpromo.Click ''''''''''''''''''''
        If (dgvsalles.Rows.Count <= 0) Then
            MsgBox("aucun local n'est associé a ce groupe", MsgBoxStyle.Critical, "pas de locaux!")
            Exit Sub
        End If
        If (modules.Items.Count <> modules.CheckedItems.Count) Then
            MsgBox("attention ,il faut associé des locaux pour tous les modules", MsgBoxStyle.Critical, "module non traité")
            Exit Sub
        End If
        If (tester_nb_places() = 0) Then
            Exit Sub
        End If
        Dim promo As String

        Dim f As Integer = CType(gp.Text, Int32)

        If Not (arr.Contains(f)) Then
            arr.Add(f)


            Dim k As Integer = 0
            While (k < modules.Items.Count)

                If (modules.GetItemCheckState(k) = CheckState.Checked) Then
                    add_groupe(modules.Items(k), gp.Text)
                End If
                k = k + 1
            End While
            update_emploi()
        End If
        view3.Visible = False
        dgvsalles.Visible = False


        For i As Integer = 0 To dt2.Rows.Count - 1 Step 1
            cmd = New OleDbCommand("insert into [local] values (@locall,@max,@gp)", con)
            With cmd.Parameters

                .AddWithValue("@locall", dt2.Rows(i)(1).ToString).DbType = DbType.String
                .AddWithValue("@max", dt2.Rows(i)(2).ToString).DbType = DbType.Int32
                .AddWithValue("@gp", dt2.Rows(i)(3).ToString).DbType = DbType.Int32
            End With
            Try
                con.Open()
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "attention!!!")
                Exit Sub
            Finally
                con.Close()

            End Try
        Next
        promo = nompromo.Text.ToString
        ajoutBW.Visible = False
        For i As Integer = 0 To listepromo.CheckedItems.Count - 1
            If (listepromo.CheckedItems(i).ToString = promo) Then
                If (i < listepromo.CheckedItems.Count - 1) Then
                    nompromo.Text = listepromo.CheckedItems(i + 1)
                    Exit For
                Else
                    If (RadioButton1.Checked = False) Then
                        remplir_exam()
                        pnlAffectation2.Visible = False
                        Dim frm1 As New Form3
                        For Indice As Integer = 0 To listepromo.CheckedItems.Count - 1
                            If listepromo.CheckedItems(Indice).ToString = "1CP" Then
                                frm1.CP1.Enabled = True
                            End If
                            If listepromo.CheckedItems(Indice).ToString = "2CP" Then
                                frm1.CP2.Enabled = True
                            End If
                            If listepromo.CheckedItems(Indice).ToString = "1CS" Then
                                frm1.CS1.Enabled = True
                            End If
                            If listepromo.CheckedItems(Indice).ToString = "2SQ" Then
                                frm1.SQ2.Enabled = True
                            End If
                            If listepromo.CheckedItems(Indice).ToString = "2ST" Then
                                frm1.ST2.Enabled = True
                            End If
                            If listepromo.CheckedItems(Indice).ToString = "2SL" Then
                                frm1.SL2.Enabled = True
                            End If
                        Next
                        frm1.namb.Text = nameBdd
                        frm1.sem.Text = semestre.Text
                        frm1.exm.Text = examende.Text
                        frm1.Show()
                        Me.Close()
                        Exit Sub

                    End If
                    pnlAffectation2.Visible = False
                    view3.Visible = False
                    nompromo.Visible = False
                    ' traitement.Visible = True
                    pnlAffectation3.Visible = True

                    'trait.Visible = True
                    dgvfianle.Visible = True
                    aff_auto(promo)

                    ListBox_prm.SetSelected(0, True)

                    ' dgvfianle.Visible = True
                    pnlAffectation3.Visible = False
                    Buttonatt.Visible = False
                    affichage()
                    Exit Sub
                End If
            End If
        Next
        modules.Items.Clear()
        net()
        net2()
        For i As Integer = 0 To dgvmodules.Rows.Count - 1
            dgvmodules.Rows.RemoveAt(0)

        Next
        filldgv()
        allerausalles.Enabled = True
        update.Enabled = True
        nextpromo.Visible = True

        valider_view2()
        gp.Text = (arr.Item(arr.Count - 1) + 1).ToString
        dgvsalles.Columns.Remove("a")
        dgvsalles.Columns.Remove("b")
        dgvsalles.Columns.Remove("c")
        dgvsalles.Columns.Remove("d")
        For i As Integer = 0 To dgvsalles.Rows.Count - 1
            dgvsalles.Rows.RemoveAt(0)
        Next

        addmodule.Enabled = False

        arr.Clear()
        aff_auto(promo)
        getting_modules(nompromo.Text, semestre.Text)
        addmodule.Enabled = True
        maxgp = f

    End Sub

    Private Sub ListBox_prm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox_prm.SelectedIndexChanged '''''''''''''''''
        select_prom()
        ' select_gp()

    End Sub

    Private Sub ListBox_gp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox_gp.SelectedIndexChanged '''''''''
        select_gp()
        Try
            Dim da As New OleDbDataAdapter("select [emploi.module] from [emploi] where [emploi].gp = " & ListBox_gp.SelectedItem(0).ToString, con)
            Dim newdt As New DataTable
            da.Fill(newdt)
            Label12.Text = "Les modules : "
            For Each row In newdt.Rows
                Label12.Text &= row(0).ToString & " / "
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

    End Sub

    Private Sub ListBox_salle_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox_salle.SelectedIndexChanged ''''''''''''''''''''''

        affichage()

    End Sub

    Sub select_gp() ''''''''''''''''''''''
        Dim querry As String = Nothing
        If (ListBox_gp.SelectedItems.Count = 0) Then
            dt6.Rows.Clear()
            dt6.Columns.Clear()
            For i As Integer = 0 To dgvfianle.Rows.Count - 1
                dgvfianle.Rows.RemoveAt(0)
            Next
            Exit Sub
        End If
        querry = ListBox_gp.SelectedItems(0).ToString

        Dim str2 As String = "select [local.salle] from [local] where [local].gp = " & querry
        remplir_dt(str2, dt3)
        ListBox_salle.Items.Clear()
        ' checkmodule.Items.Clear()
        For Each item In dt3.Rows
            ListBox_salle.Items.Add(item(0).ToString)
            ' checkmodule.Items.Add(item(0).ToString)
            'checkmodule.SetItemChecked(checkmodule.Items.Count - 1, True)
        Next
        ListBox_salle.SetSelected(0, True)

    End Sub

    Sub select_prom() ''''''''''''''''''''''''''
        Dim querry As String = Nothing
        If (ListBox_prm.SelectedItems.Count = 0) Then
            dt6.Rows.Clear()
            dt6.Columns.Clear()
            For i As Integer = 0 To dgvfianle.Rows.Count - 1
                dgvfianle.Rows.RemoveAt(0)
            Next
            Exit Sub
        End If
        querry = "'" & ListBox_prm.SelectedItems(0).ToString & "'"
        Dim str As String = "select distinct(emploi.gp) from [Salles],[local],[emploi],[MODULES],[exam] where [local].gp=exam.gp and [exam.type_exam]='" + examende.Text.ToString + "' and [exam.semestre]='" + semestre.Text.ToString + "' and [local].gp=emploi.gp and [local.salle]=[Salles.Code_Sal] and [emploi.module]=[MODULES.Code_Mat] and ([MODULES.Niveau]= " & querry & ")"
        remplir_dt(str, dt6)
        ListBox_gp.Items.Clear()
        'checkgp.Items.Clear()
        For Each item In dt6.Rows
            ListBox_gp.Items.Add(item(0).ToString)
            'checkgp.Items.Add(item(0).ToString)

            'checkgp.SetItemChecked(checkgp.Items.Count - 1, True)
        Next
        ListBox_gp.SetSelected(0, True)

    End Sub

    Private Sub change_gp() ''''''''''''''''
        If (modules.CheckedItems.Count <= 0) Then
            For i As Integer = 0 To dgvsalles.Rows.Count - 1
                dgvsalles.Rows.RemoveAt(0)

            Next
            Dim arra As New ArrayList
            For j As Integer = 0 To dt2.Rows.Count - 1
                If (dt2.Rows(j)(3) = gp.Text) Then

                    'dt2.Rows.RemoveAt(j)
                    arra.Add(j)
                End If
            Next
            Dim k As Integer = 0
            For Each pos As Integer In arra

                dt2.Rows.RemoveAt(pos - k)
                k += 1
            Next
        End If
        For Each row As DataGridViewRow In dgvsalles.Rows
            row.Cells(0).Value = ""
            For Each item In modules.CheckedItems
                row.Cells(0).Value += item + "/"

            Next
            row.Cells(0).ToString.Remove(row.Cells(0).ToString.Length - 1)
        Next
    End Sub

    Function tester_nb_places() As Integer ''''''''''''''''''''''''
        Dim nb As Integer = 0
        Dim nb2 As Integer = 0
        Dim num As Integer = 0

        For Each row In dt2.Rows
            If (row(3) = gp.Text.ToString) Then
                nb += CType(row(2).ToString, Int32)
            End If
        Next
        If (semestre.Text = "S1") Then
            If (modules.CheckedItems(0).ToString = "BW") And (sections.CheckedItems.Count > 0) Then

                For Each item In sections.CheckedItems
                    Dim ro() As DataRow = dt3.Select("Sect='" + item.ToString + "'")
                    If ro.Count > 0 Then
                        num += ro.Count
                    End If
                Next
            ElseIf (modules.CheckedItems(0) = "ANG1") And (sections.CheckedItems.Count > 0) Then
                For Each item In sections.CheckedItems
                    Dim ro() As DataRow = dt3.Select("Sect='" + item.ToString + "'")
                    If ro.Count > 0 Then
                        nb2 += ro.Count
                    End If
                Next
            End If
        ElseIf (semestre.Text = "S2") Then
            If (modules.CheckedItems(0).ToString = "ANG1") And (sections.CheckedItems.Count > 0) Then
                For Each item In sections.CheckedItems
                    Dim ro() As DataRow = dt3.Select("Sect='" + item.ToString + "'")
                    If ro.Count > 0 Then
                        num += ro.Count
                    End If
                Next
            ElseIf (modules.CheckedItems(0) = "BW") And (sections.CheckedItems.Count > 0) Then
                For Each item In sections.CheckedItems
                    Dim ro() As DataRow = dt3.Select("Sect='" + item.ToString + "'")
                    If ro.Count > 0 Then
                        nb2 += ro.Count
                    End If
                Next
            End If

        End If
        'If (nb2 > 0) And (nb <> nb2) Then
        '    MsgBox("attention le nombre de places n'egale pas au nombre d'étudiants,il y a " & nb2 & " étudiants et vous avez reserver  " & nb & " places", MsgBoxStyle.Critical, "nombre de places insuffisants")
        '    local.Focus()
        '    Return 0
        '    Exit Function
        'ElseIf (num > 0) And (dt3.Rows.Count - num <> nb) Then
        '    MsgBox("attention le nombre de places n'egale pas au nombre d'étudiants ,il y a " & dt3.Rows.Count - num & " étudiants et vous avez reserver  " & nb & " places", MsgBoxStyle.Critical, "nombre de places insuffisants")
        '    local.Focus()
        '    Return 0
        '    Exit Function
        'ElseIf (nb <> dt3.Rows.Count) And (num = 0) And (nb2 = 0) Then
        '    MsgBox("attention le nombre de places n'egale pas au nombre d'étudiants ,il y a " & dt3.Rows.Count & " étudiants et vous avez reserver " & nb & " places", MsgBoxStyle.Critical, "nombre de places insuffisants")
        '    local.Focus()
        '    Return 0
        '    Exit Function
        'End If
        Return 1
    End Function

    Function test_gp_modules() As Integer ''''''''''''''''''''''
        For Each item In modules.CheckedItems
            If (sections.CheckedItems.Count > 0) And (item = "BW" Or item = "ANG1") And (modules.CheckedItems.Count > 1) Then
                MsgBox("attention le module : " & item & " doit ètre seul dans un groupe ")
                Return 0
                Exit Function
            End If
        Next
        If (modules.Items.Count = modules.CheckedItems.Count) Then
            newgp.Visible = False
            nextpromo.Visible = True
        Else
            newgp.Visible = True
        End If
        If (modules.CheckedItems.Count <= 0) Then
            MsgBox("il faut choisir au moins un module !", MsgBoxStyle.Critical, "attention")
            Return 0
            Exit Function
        End If
        If (dgvsalles.Rows.Count > 0) Then
            change_gp()
        End If
        Return 1
    End Function

    Sub alleraulocalaux() '''''''''''''''''''''''''
        view4.Visible = True
        'modules.Enabled = False
    End Sub

    Sub getting_modules(ByVal promo As String, ByVal sem As String) '''''''''''''''''
        dt3 = New DataTable
        Dim strr As String
        If (semestre.Text = "anu") Then
            strr = "SELECT MODULES.Code_Mat FROM MODULES WHERE MODULES.Niveau = '" + promo.ToString + "'"
        Else
            strr = "SELECT MODULES.Code_Mat FROM MODULES WHERE MODULES.Niveau = '" + promo.ToString + "'and Sem ='" + sem.ToString + "'"
        End If

        remplir_dt(strr, dt3)

        For i As Integer = 0 To dt3.Rows.Count - 1
            record.Items.Add(dt3.Rows(i)(0))
        Next i
    End Sub

    Function conn_bdd(ByVal path2 As String) '''''''''''''''''''
        Try

            con = New OleDbConnection("Provider=Microsoft.ace.OleDb.12.0;Data Source=" + path2)
            cmd = New OleDbCommand("create table emploi ([module] varchar(20) ,[date1] datetime,[debut] datetime,[fin] datetime,gp int ,primary key([module],date1))", con)
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            cmd = New OleDbCommand("create table [local] ([salle] varchar(20) ,max1 int,gp int ,primary key([salle],gp))", con)
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            cmd = New OleDbCommand("create table exam (gp int primary key,[type_exam] varchar(20),[semestre] varchar(20))", con)
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            cmd = New OleDbCommand("create table affect ([matricule] varchar(20) ,gp int ,[local] varchar(20),pos int,primary key([matricule],gp))", con)
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()

            cmd = New OleDbCommand("create table CrystalReport (N int primary key, [Matricule] varchar(20), [Nom] varchar(20), [Prenom] varchar(20), [Section] varchar(20), [Gr] varchar(20), [Local] varchar(20), [Position] varchar(20), [Module] varchar(20), [DateExam] datetime, [HDebut] datetime, [HFin] datetime )", con)
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "echec")
            Return 0
            Exit Function
        Finally
            con.Close()
        End Try
        Return 1
    End Function

    Function tester_promo(ByVal path As String) '''''''''''''''''''''''''''
        Try
            con = New OleDbConnection("Provider=Microsoft.ace.OleDb.12.0;Data Source=" + path)
            Dim da As New OleDbDataAdapter("select distinct([MODULES.Niveau]) from [MODULES],[emploi],[exam] where  [MODULES.Code_Mat]=[emploi.module] and exam.gp=emploi.gp and [exam.type_exam]='" & examende.Text.ToString & "' and [exam.semestre]='" & semestre.Text.ToString & "'", con)
            da.Fill(dt)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return 0
            Exit Function
        Finally
            con.Close()
        End Try
        Dim msg As String = "les promos: "
        Dim prm As String
        Dim cpt As Integer = 0
        Dim ro() As DataRow
        Dim diaform As New Dialog1
        For Each item In listepromo.CheckedItems
            ro = dt.Select("[MODULES.Niveau]='" + item.ToString + "'")

            If ro.Count > 0 Then
                diaform.CheckedListBox1.Items.Add(item)
                ' MsgBox("l'affectation de la promo : " & item & " est deja faite", MsgBoxStyle.Critical, "attention")
                cpt = cpt + 1
                msg += item + ", "
                prm = item


                'Return 0
                'Exit Function

            End If
        Next
        If (cpt > 0) Then

            If cpt = 1 Then

                msg += " est deja affectée pour cet examen, voulez vous une nouvelle affectation?"
                If (MsgBox(msg, MsgBoxStyle.YesNo, "promo deja affectée") = MsgBoxResult.Yes) Then

                    Try
                        'cmd = New OleDbCommand("delete * from Emploi,loc,affect where MODULES.Code_Mat=Emploi.module", con)

                        'cmd = New OleDbCommand("DELETE affect.*, loc.*, Emploi.* FROM ((affect INNER JOIN Emploi ON affect.gp = Emploi.gp) INNER JOIN loc ON Emploi.gp = loc.gp) INNER JOIN MODULES ON Emploi.module = MODULES.Code_Mat where MODULES.Niveau='" + diaform.CheckedListBox1.Items(0) + "'", con)
                        dt.Columns.Clear()
                        dt.Rows.Clear()
                        da = New OleDbDataAdapter("select emploi.gp from [emploi],[MODULES],[exam] where [emploi.module]= [MODULES.Code_Mat] and exam.gp=emploi.gp  and [MODULES.Niveau]='" + diaform.CheckedListBox1.Items(0).ToString + "'  and [exam.type_exam]='" + examende.Text.ToString + "' and [exam.semestre]='" + semestre.Text.ToString + "'", con)
                        da.Fill(dt)
                        'cmd = New OleDbCommand("delete emploi.* from [emploi],[MODULES],[exam] where [emploi.module]= [MODULES.Code_Mat] and exam.gp=emploi.gp  and [MODULES.Niveau]='" + diaform.CheckedListBox1.Items(0) + "'  and [exam.type_exam]='" + examende.Text + "' and [exam.semestre]='" + semestre.Text.ToString + "'", con)

                        'con.Open()
                        'cmd.ExecuteNonQuery()
                        'con.Close()

                        Dim query As String = Nothing
                        Dim query2 As String = Nothing
                        Dim query3 As String = Nothing
                        Dim query4 As String = Nothing
                        For Each row In dt.Rows
                            If (query = String.Empty) Then
                                query = "[local].gp = " & row(0).ToString
                                query2 = "[emploi].gp = " & row(0).ToString
                                query3 = "[exam].gp = " & row(0).ToString
                                query4 = "[affect].gp = " & row(0).ToString
                            Else
                                query += " or [local].gp= " & row(0).ToString
                                query2 += " or [emploi].gp= " & row(0).ToString
                                query3 += " or [exam].gp= " & row(0).ToString
                                query4 += " or [affect].gp= " & row(0).ToString
                            End If

                        Next

                        'cmd = New OleDbCommand("delete [local].*,affect.*,exam.*,emploi.* from [local],[affect],[exam],[emploi] where emploi.gp= exam.gp and [local].gp=affect.gp and ( " + query + ") and [local].gp=exam.gp", con)
                        cmd = New OleDbCommand("delete * from [local] where  " + query, con)
                        con.Open()
                        cmd.ExecuteNonQuery()
                        con.Close()
                        cmd = New OleDbCommand("delete * from [emploi] where  " + query2, con)
                        con.Open()
                        cmd.ExecuteNonQuery()
                        con.Close()
                        cmd = New OleDbCommand("delete * from [exam] where  " + query3, con)
                        con.Open()
                        cmd.ExecuteNonQuery()
                        con.Close()
                        cmd = New OleDbCommand("delete * from [affect] where  " + query4, con)
                        con.Open()
                        cmd.ExecuteNonQuery()


                        Return 1
                        Exit Function
                    Catch ex As Exception
                        MsgBox(ex.Message)
                        Return 0
                        Exit Function
                    Finally
                        con.Close()
                    End Try
                Else

                    Dim a As Integer = -1
                    For i As Integer = 0 To listepromo.CheckedItems.Count

                        If (listepromo.CheckedItems(i) = prm) Then
                            ' listepromo.Items.Remove(item)
                            a = i
                            Exit For

                        End If
                    Next
                    If (a >= 0) Then
                        listepromo.SetItemChecked(a, False)
                    End If
                    diaform.CheckedListBox1.Items.Clear()
                    If (listepromo.CheckedItems.Count <= 0) Then
                        Return 0
                        Exit Function
                    End If
                    Return 1
                    Exit Function
                End If
            Else



                msg += " sont deja affectée pour cet examen, voulez vous faire une nouvelle affectation pour ces promos ?"
                If (MsgBox(msg, MsgBoxStyle.YesNo, "promo deja affectée") = MsgBoxResult.Yes) Then
                    Try
                        diaform.ShowDialog()
                        For Each item In diaform.CheckedListBox1.CheckedItems
                            dt.Rows.Clear()
                            dt.Columns.Clear()

                            da = New OleDbDataAdapter("select emploi.gp from [emploi],[MODULES],[exam] where emploi.gp=exam.gp and [exam.type_exam]='" + examende.Text.ToString + "' and [exam.semestre]='" + semestre.Text.ToString + "' and [emploi.module]= [MODULES.Code_Mat]  and [MODULES.Niveau]='" + item.ToString + "'", con)
                            da.Fill(dt)
                            Dim query As String = Nothing
                            Dim query2 As String = Nothing
                            Dim query3 As String = Nothing
                            Dim query4 As String = Nothing
                            For Each row In dt.Rows
                                If (query = String.Empty) Then
                                    query = "[local].gp = " & row(0).ToString
                                    query2 = "[emploi].gp = " & row(0).ToString
                                    query3 = "[exam].gp = " & row(0).ToString
                                    query4 = "[affect].gp = " & row(0).ToString
                                Else
                                    query += " or [local].gp= " & row(0).ToString
                                    query2 += " or [emploi].gp= " & row(0).ToString
                                    query3 += " or [exam].gp= " & row(0).ToString
                                    query4 += " or [affect].gp= " & row(0).ToString
                                End If

                            Next

                            'cmd = New OleDbCommand("delete [local].*,affect.*,exam.*,emploi.* from [local],[affect],[exam],[emploi] where emploi.gp= exam.gp and [local].gp=affect.gp and ( " + query + ") and [local].gp=exam.gp", con)
                            cmd = New OleDbCommand("delete * from [local] where  " + query, con)
                            con.Open()
                            cmd.ExecuteNonQuery()
                            con.Close()
                            cmd = New OleDbCommand("delete * from [emploi] where  " + query2, con)
                            con.Open()
                            cmd.ExecuteNonQuery()
                            con.Close()
                            cmd = New OleDbCommand("delete * from [exam] where  " + query3, con)
                            con.Open()
                            cmd.ExecuteNonQuery()
                            con.Close()
                            cmd = New OleDbCommand("delete * from [affect] where  " + query4, con)
                            con.Open()
                            cmd.ExecuteNonQuery()

                        Next
                    Catch ex As Exception
                        MsgBox(ex.Message)
                        Return 0
                        Exit Function
                    Finally
                        con.Close()
                    End Try

                End If

            End If

            For Each item In diaform.CheckedListBox1.Items
                If (diaform.CheckedListBox1.Items.Count > 1) Then
                    If (Not (diaform.CheckedListBox1.CheckedItems.Contains(item))) Then
                        ' listepromo.Items.Remove(item)
                        listepromo.SetItemChecked(listepromo.Items.IndexOf(item), False)
                    End If
                End If

            Next
        End If
        If (listepromo.CheckedItems.Count <= 0) Then
            diaform.CheckedListBox1.Items.Clear()
            Return 0
            Exit Function
        End If
        Return 1
    End Function

    Function testerbdd_et_con() As Integer ''''''''''''''''''''''
        nameBdd = anne.Text.ToString + ".accdb"
        Try
            path2 = "../../Sauvegarde/" + nameBdd


            Dim file As System.IO.File
            If File.Exists(path2) Then
                If (tester_promo(path2) = 0) Then
                    Return 0
                    Exit Function
                Else
                    Return 1
                    Exit Function
                End If
            Else

                My.Computer.FileSystem.CopyFile(defultpath.Text, path2, Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "fail")
            Return 0
            Exit Function
        End Try
        If (conn_bdd(path2) = 0) Then
            Return 0
            Exit Function
        End If
        Return 1
    End Function

    Function tester_view1() As Integer ''''''''''''''''
        If (semestre.Text <> "S1" And semestre.Text <> "S2" And semestre.Text <> "anu") Then
            MsgBox("le semestre n'est pas valide", MsgBoxStyle.Critical, "semestre")
            Return 0
            Exit Function
        End If
        If (examende.Text = String.Empty Or semestre.Text = String.Empty Or anne.Text = String.Empty Or critere.Text = String.Empty) Then
            MsgBox("attention il faut remplir tout les champs", MsgBoxStyle.Critical, "attention")
            Return 0
            Exit Function
        End If
        If (RadioButton1.Checked = False And RadioButton2.Checked = False) Then
            MsgBox("veuillez selectionner un mode d'affectation")
            Return 0
            Exit Function
        End If

        If listepromo.CheckedItems.Count <= 0 Then
            MsgBox("il faut choisir au moins une promo", MsgBoxStyle.Critical, "attention")
            Return 0
            Exit Function


        End If

        Return 1
    End Function

    Sub allerde1a2() ''''''''''''''''''
        ListBox_prm.Items.Clear()

        For Each item In listepromo.CheckedItems

            ListBox_prm.Items.Add(item)
        Next
        pnlAffectation1.Visible = False
        pnlAffectation2.Visible = True
        view3.Visible = False
        Dim dd As New DataTable
        Try
            Dim str As String = "select max(exam.gp) from exam"
            Try
                remplir_dt(str, dd)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            If (dd.Rows.Count > 0) And (dd.Rows(0)(0).ToString <> String.Empty) Then
                maxgp = CType(dd.Rows(0)(0).ToString, Int32)
                gp.Text = (maxgp + 1).ToString

            Else

                gp.Text = "1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "attention")
            Exit Sub
        End Try
        filldgv()
        net()
        nompromo.Text = listepromo.CheckedItems(0).ToString
    End Sub

    Sub tri_crit(ByVal critere As String, ByVal promo As String) '''''''''''''''''
        '
        Dim querry As String
        If (listepromo.CheckedItems.Contains("1CP")) Then
            critere = "Matricule"
        End If


        If (critere = "Rang" Or critere = "MoyAnu") Then
            querry = "select ETUDIANTS.Matricule,ETUDIANTS.Sect from [ETUDIANTS] inner join [moyennes] on ETUDIANTS.Matricule= moyennes.Matricule where [ETUDIANTS.Promo]= '" + promo + "' and [ETUDIANTS.Sect] <> '" + String.Empty + "' order by [moyennes." + critere + "]"
        Else
            querry = "select ETUDIANTS.Matricule,ETUDIANTS.Sect from [ETUDIANTS] where [ETUDIANTS.Promo]= '" + promo + "' and [ETUDIANTS.Sect] <> '" + String.Empty + "' order by [ETUDIANTS." + critere + "]"
        End If
        remplir_dt(querry, dt3)
    End Sub

    Sub affecter_local(ByVal local As String, ByVal capacite As String, ByVal max As Integer, ByVal gp As String) '''''''''''''
        Dim j As Integer = 0
        Dim tab As ArrayList = New ArrayList
        Dim pos As Integer = 1
        Dim pas As Integer = (dt0.Rows.Count / max)
        If (pas * max) > dt0.Rows.Count Then
            pas = pas - 1
        End If

        If (pas Mod 2 <> 0) Then
            pas = pas - 1
            If (pas <= 0) Then
                pas = 1
            End If
        End If
        For i As Integer = 0 To max - 1
            Try
                cmd = New OleDbCommand("insert into affect values(@mat,@gp,@locall,@pos)", con)
                With cmd.Parameters

                    .AddWithValue("@mat", dt0.Rows(j)(0).ToString).DbType = DbType.String
                    .AddWithValue("@gp", gp).DbType = DbType.Int32
                    .AddWithValue("@locall", local).DbType = DbType.String
                    .AddWithValue("@pos", pos).DbType = DbType.Int32

                End With

                con.Open()
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "affect_local")
                Exit Sub
            Finally
                con.Close()
            End Try
            tab.Add(j)
            j = j + pas
            pos = pos + 2
            If (pos > capacite) Then
                pos = 2
            End If
        Next
        For k As Integer = 0 To tab.Count - 1
            If (k = 0) Then
                dt0.Rows(tab.Item(k)).Delete()
            Else
                dt0.Rows(tab.Item(k) - k).Delete()
            End If

        Next
    End Sub

    Sub affecter_promo(ByVal promo As String) '''''''''''''''''''''''''''


        Dim str As String = "select emploi.gp,[local.salle],Salles.Cont_Exam,[local].max1 from [Salles],[local],[emploi],[MODULES],[exam] where [local].gp=exam.gp and [exam.type_exam]='" + examende.Text.ToString + "' and [exam.semestre]='" + semestre.Text.ToString + "' and [local].gp=emploi.gp and [local.salle]=[Salles.Code_Sal] and [emploi.module]=[MODULES.Code_Mat] and [MODULES.Niveau]= '" & promo & "'"

        remplir_dt(str, dt6)
    End Sub

    Sub aff_auto(ByVal promo As String) '''''''''''''''''''''''''''
        'view3.Visible = False
        If (RadioButton1.Checked = True) Then
            remplir_exam()
        Else
            remplir_exam()
            Exit Sub
        End If

        Try
            'tri_crit(critere.Text, promo)
            affecter_promo(promo)
            dt0.Columns.Clear()

            dt0.Rows.Clear()
            dt0.Columns.Add("mat")

            For Each item In dt3.Rows
                Dim ro As DataRow = dt0.NewRow
                ro(0) = item(0).ToString
                dt0.Rows.Add(ro)
                'MsgBox(dt0.Rows(0)(0).ToString)
            Next
            Dim grp As String = dt6.Rows(0)(0).ToString
            Dim tes As New DataTable
            Dim querry As String = Nothing
            Dim selecti As String = Nothing
            Dim cpt As Integer = 0
            For Each row In dt6.Rows

                querry = "select [emploi.module] from [emploi] where [emploi].gp = " & row(0).ToString
                tes.Rows.Clear()
                tes.Columns.Clear()
                remplir_dt(querry, tes)
                If (row(0).ToString <> grp) Or (selecti = String.Empty) Then

                    If (tes(0)(0).ToString = "BW" Or tes(0)(0).ToString = "ANG1") And (sections.CheckedItems.Count > 0) Then

                        If (semestre.Text = "S1") Then
                            If (tes(0)(0).ToString = "BW") Then

                                For Each item In sections.CheckedItems

                                    If (selecti = String.Empty) Then
                                        selecti = "Sect <> '" & item.ToString & "'"
                                    Else
                                        selecti &= " and Sect <> '" & item.ToString & "'"
                                    End If
                                Next
                            Else
                                For Each item In sections.CheckedItems
                                    If (selecti = String.Empty) Then
                                        selecti = "Sect = '" & item.ToString & "'"
                                    Else
                                        selecti &= " and Sect = '" & item.ToString & "'"
                                    End If

                                Next
                            End If

                        Else
                            If (tes(0)(0).ToString = "BW") Then

                                For Each item In sections.CheckedItems

                                    If (selecti = String.Empty) Then
                                        selecti = "Sect = '" & item.ToString & "'"
                                    Else
                                        selecti &= " and Sect = '" & item.ToString & "'"
                                    End If
                                Next
                            Else
                                For Each item In sections.CheckedItems
                                    If (selecti = String.Empty) Then
                                        selecti = "Sect <> '" & item.ToString & "'"
                                    Else
                                        selecti &= " and Sect <> '" & item.ToString & "'"
                                    End If

                                Next
                            End If
                        End If

                        Dim ro() As DataRow = dt3.Select(selecti)
                        If ro.Count > 0 Then
                            dt0.Columns.Clear()
                            dt0.Rows.Clear()
                            dt0.Columns.Add("mat")
                            For Each it In ro
                                Dim ro1 As DataRow = dt0.NewRow
                                ro1(0) = it(0).ToString
                                dt0.Rows.Add(ro1)
                            Next
                        End If
                    End If
                ElseIf row(0).ToString <> grp Then
                    dt0.Columns.Clear()

                    dt0.Rows.Clear()
                    dt0.Columns.Add("mat")
                    For Each item In dt3.Rows
                        Dim ro As DataRow = dt0.NewRow
                        ro(0) = item(0).ToString
                        dt0.Rows.Add(ro)

                    Next

                End If
                If (row(0).ToString <> grp) Then
                    grp = row(0).ToString
                    selecti = ""
                End If
                If cpt <> 0 Then
                    MsgBox(cpt)
                    If (row(0).ToString = dt6.Rows(cpt - 1)(0).ToString) Then
                        cpt = cpt + 1
                        Exit Sub
                    End If
                End If
                affecter_local(row(1), row(2), row(3), row(0))
                ' Me.BindingContext(dt).Position
                cpt = cpt + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "aff_auto")
            Exit Sub
        End Try
        'dgvsalles.DataSource = dt6
        'dgvsalles.Visible = True

    End Sub

    Sub affichage() '''''''''''''''''''''
        dt6.Rows.Clear()
        dt6.Columns.Clear()
        For i As Integer = 0 To dgvfianle.Rows.Count - 1
            dgvfianle.Rows.RemoveAt(0)
        Next
        If (pnlAffectation3.Visible = False) Then
            pnlAffectation3.Visible = True
            select_prom()
            'select_gp()

            Exit Sub
        End If
        If (ListBox_prm.SelectedItems.Count = 0) Then
            Exit Sub
        End If
        If (pnlAffectation3.Visible = True) And ((ListBox_prm.SelectedItems.Count = 0) Or (ListBox_gp.SelectedItems.Count = 0) Or (ListBox_salle.SelectedItems.Count = 0)) Then
            Exit Sub
        End If
        Dim condit As String = "and ( "

        Dim p As Integer = 0
        condit &= " [ETUDIANTS.Promo]= '" & ListBox_prm.SelectedItems(0).ToString & "' "
        Dim querry As String = " "
        If (ListBox_gp.SelectedItems.Count > 0) Then
            querry = " and ( [emploi].gp= " & ListBox_gp.SelectedItems(0).ToString & " )"
        End If



        Dim querry2 As String = " "
        If (ListBox_salle.SelectedItems.Count > 0) Then
            querry2 = " and ( [local.salle]= '" & ListBox_salle.SelectedItems(0).ToString & "' )"
        End If



        condit &= ")"

        Dim str As String = "select distinct( [ETUDIANTS.NomEtud]),[affect.matricule],[affect.local],affect.pos,[ETUDIANTS.Sect] ,[ETUDIANTS.Prenoms],[ETUDIANTS.Gr]" &
                                       "from [ETUDIANTS],[affect],[emploi],[local] " &
                                       "where [ETUDIANTS.Matricule]=[affect.matricule]" &
                                       " and affect.gp=[local].gp " &
                                       "and [local].gp=emploi.gp " & condit & querry & querry2 & " order by affect.pos "

        If (remplir_dt(str, dt6) = 0) Then
            MsgBox("erreur dans l'affichage", MsgBoxStyle.Critical, "erreur ")
            Exit Sub
        End If

        pnlAffectation2.Visible = False
        pnlAffectation3.Visible = True
        dgvfianle.DataSource = dt6
        dgvfianle.Columns(2).Visible = False
        dgvfianle.Columns(5).Visible = False

    End Sub

    Sub valider_view2() '''''''''''''''''''''''''
        record.Enabled = True
        addmodule.Enabled = True
        modifButtons.Enabled = True
        allerausalles.Enabled = True

    End Sub

    Sub remplir_exam() ''''''''''''''''''''''''''''

        Dim str As String = "select distinct(emploi.gp) from emploi where emploi.gp >" & maxgp
        Dim ddd As New DataTable
        remplir_dt(str, ddd)
        For Each row In ddd.Rows
            cmd = New OleDbCommand("insert into exam values(@gp,@typ_ex,@sem)", con)
            With cmd.Parameters

                .AddWithValue("@gp", row(0).ToString).DbType = DbType.Int32
                .AddWithValue("@typ_ex", examende.Text).DbType = DbType.String
                .AddWithValue("@sem", semestre.Text).DbType = DbType.String

            End With
            Try
                con.Open()
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "remplir_exam")
                Exit Sub
            Finally
                con.Close()
            End Try

        Next
    End Sub

    Dim tabl_aff(,) As String = New String(51, 5) {}
    Dim buttons As New List(Of System.Windows.Forms.Button)
    Dim buttutil As New List(Of System.Windows.Forms.Button)

    Function tester_local2() As Integer ''''''''''''''''''''''''''''
        dt6.Rows.Clear()
        dt6.Columns.Clear()
        local.Items.Clear()
        getlocalaux()
        If local.Text <> String.Empty Then
            Dim ro() As DataRow = dt6.Select("Code_Sal='" + local.Text + "'")
            If ro.Count = 0 Then
                MsgBox("le local n'est pas valide ", MsgBoxStyle.Critical, "attention")
                local.Focus()
                Return 0
                Exit Function
            End If

            If (max.Value = 0) Then
                MsgBox("le local ne peut pas etre vide", MsgBoxStyle.Critical, "local vide")
                Return 0
                Exit Function
            End If

            Dim max1 As Integer = CType(ro(0)(1), Int32)
            If max.Value = 0 Or max.Value > max1 Then
                MsgBox("vous avez dépacer la capacité du local ; le max est :" & max1, MsgBoxStyle.Critical, "attention")
                max.Focus()
                Return 0
                Exit Function
            End If
        End If
        Return 1
    End Function

    Function tester_local() As Integer ''''''''''''''''''''''''''''
        Dim query As String
        Dim dt5 As New DataTable

        If (local.Text = String.Empty) Then
            MsgBox("il faut choisir un local", MsgBoxStyle.Critical, "attention")
            Return 0
            Exit Function
        End If
        For i As Integer = 0 To dgvsalles.Rows.Count - 1
            If (dgvsalles.Rows(i).Cells(1).Value = local.Text) Then
                MsgBox("attention, un local ne peut pas etre utilise 2 fois pour le meme groupe de modules", MsgBoxStyle.Critical, "local exist!")
                Return 0
                Exit Function
            End If
        Next
        Try

            For Each item In modules.CheckedItems
                Dim ro() As DataRow = dt.Select("module='" + item + "'")

                If ro.Count > 0 Then
                    Dim modd As String = ""
                    If (ro(0)(0).ToString = "BW" And sections.CheckedItems.Count > 0) Then
                        modd = "ANG1"
                    ElseIf (ro(0)(0).ToString = "ANG1" And sections.CheckedItems.Count > 0) Then
                        modd = "BW"
                    End If
                    Dim dt As String = CDate(ro(0)(1))
                    query = "select emploi.*,[local].salle from emploi,[local] where [emploi.date1] = #" + dt + "# and emploi.gp=[local].gp"
                    'query = "select * from [emploi] inner join [local] on  emploi.gp=local.gp where [emploi.date1] = #" + ro(0)(1).ToString + "# and [local.salle]= '" + local.Text.ToString + "'"
                    If (remplir_dt(query, dt5) = 0) Then
                        Return 0
                        Exit Function
                    End If

                    For Each row In dt5.Rows

                        If (row(5).ToString = local.Text.ToString And row(0) <> item And row(0) <> modd And (TimeValue(row(2)) <= TimeValue((ro(0)(2))) And TimeValue(row(3)) >= TimeValue((ro(0)(2))))) Then
                            MsgBox("ce local est occupe pour examen de " + row(0), MsgBoxStyle.Critical, "local occupé")
                            Return 0
                            Exit Function
                        End If
                    Next
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "attention")
            Return 0
            Exit Function
        End Try

        Return 1
    End Function

    Function remplir_var_glob(ByVal promo As String) As Integer '''''''''''''''''''''''
        nompromo.Text = promo.ToString
        If (nompromo.Text = "1CP") Then
            ajoutBW.Visible = True
        End If
        Dim str As String = "select distinct(emploi.gp) from [Salles],[local],[emploi],[MODULES],[exam] where [local].gp=exam.gp and [exam.type_exam]='" + examende.Text.ToString + "' and [exam.semestre]='" + semestre.Text.ToString + "' and [local].gp=emploi.gp and [local.salle]=[Salles.Code_Sal] and [emploi.module]=[MODULES.Code_Mat] and [MODULES.Niveau]= '" & promo & "'"
        remplir_dt(str, dt6)
        modules.Items.Clear()

        For Each row In dt6.Rows
            arr.Add(CType(row(0).ToString, Int32))


        Next
        maxgp = arr(0) - 1
        Dim str2 As String = "select [emploi.module],[emploi.date1],[emploi.debut],[emploi.fin],[emploi].gp from [Salles],[local],[emploi],[MODULES],[exam] where [local].gp=exam.gp and [exam.type_exam]='" + examende.Text.ToString + "' and [exam.semestre]='" + semestre.Text.ToString + "' and [local].gp=emploi.gp and [local.salle]=[Salles.Code_Sal] and [emploi.module]=[MODULES.Code_Mat] and [MODULES.Niveau]= '" & promo & "'"
        remplir_dt(str2, dt3)

        Dim str3 As String = "select [local].* from [Salles],[local],[emploi],[MODULES],[exam] where [local].gp=exam.gp and [exam.type_exam]='" + examende.Text.ToString + "' and [exam.semestre]='" + semestre.Text.ToString + "' and [local].gp=emploi.gp and [local.salle]=[Salles.Code_Sal] and [emploi.module]=[MODULES.Code_Mat] and [MODULES.Niveau]= '" & promo & "'"
        remplir_dt(str3, dt6)

        fill_groupes()
        filldgv()
        filldgv2()
        Dim querry As String = Nothing
        Dim querry2 As String = Nothing
        Dim querry3 As String = Nothing
        For Each item In arr
            If (querry = String.Empty) Then
                querry = " where affect.gp= " & item
                querry2 = " where [local].gp= " & item
                querry3 = " where [exam].gp= " & item
            Else
                querry += " or affect.gp= " & item
                querry2 += " or [local].gp= " & item
                querry3 += " or [exam].gp= " & item
            End If
            Dim modules As String = Nothing
            Dim ro() As DataRow = dt3.Select("gp= " & item)
            For Each item2 In ro

                Dim roww As DataRow = dt.NewRow
                roww(0) = item2(0).ToString
                roww(1) = CDate(item2(1).ToString)
                roww(2) = Hour(TimeValue(item2(2).ToString)) & ":" & Minute(TimeValue(item2(2).ToString))
                roww(3) = Hour(TimeValue(item2(3).ToString)) & ": " & Minute(TimeValue(item2(3).ToString))
                dt.Rows.Add(roww)
                modules += item2(0).ToString + "/"
                add_groupe(item2(0).ToString, item)
            Next
            modules.ToString.Remove(modules.ToString.Length - 1)
            Dim ro2() As DataRow = dt6.Select("gp= " & item)
            For Each item3 In ro2
                Dim row As DataRow = dt2.NewRow
                row(0) = modules.ToString
                row(1) = item3(0).ToString
                row(2) = item3(1).ToString
                row(3) = item.ToString
                dt2.Rows.Add(row)
            Next

        Next
        dgvmodules.DataSource = dt
        Try
            cmd = New OleDbCommand("delete * from [affect] " & querry, con)
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            cmd = New OleDbCommand("delete * from [local] " & querry2, con)
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            cmd = New OleDbCommand("delete * from [exam] " & querry3, con)
            con.Open()
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
            Return 0
            Exit Function
        Finally
            con.Close()
        End Try

        gp.Text = arr.Item(0).ToString
        'headers_dgvsalles()
        remplir_dgvsalles()

        For Each item In dt4.Rows

            If item(1) = gp.Text Then
                modules.Items.Add(item(0))
                modules.SetItemChecked(modules.Items.Count - 1, True)
            End If
        Next
        allerausalles.Enabled = False
        update.Enabled = False
        modules.Enabled = True
        view3.Visible = True
        view4.Visible = False
        local.Text = String.Empty
        max.Value = 0
        view4.Visible = True
        dgvsalles.Visible = True
        getting_modules(nompromo.Text, semestre.Text)
        tri_crit(critere.Text, nompromo.Text)


        Return 1
    End Function

    Sub reinitialiser_var_glob() '''''''''''''''''''''''''''''

        Dim query As String = Nothing
        Dim query2 As String = Nothing
        Dim query3 As String = Nothing
        Dim query4 As String = Nothing

        If (arr.Count = 0) Then
            Try
                'cmd = New OleDbCommand("delete [local].*,affect.*,exam.*,emploi.* from [local],[affect],[exam],[emploi] where emploi.gp= exam.gp and [local].gp=affect.gp and ( " + query + ") and [local].gp=exam.gp", con)

                cmd = New OleDbCommand("delete * from [emploi] where [emploi].gp =0 ", con)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()

            Catch ex As Exception
                MsgBox(ex.Message)

                Exit Sub
            Finally
                con.Close()
            End Try
            Exit Sub
        End If
        modules.Items.Clear()

        For Each item In arr

            If (query = String.Empty) Then
                query = "where [local].gp = " & item.ToString
                query2 = " where [emploi].gp = " & item.ToString
                query3 = "where [exam].gp = " & item.ToString
                query4 = "where [affect].gp = " & item.ToString
            Else
                query += " or [local].gp= " & item.ToString
                query2 += " or [emploi].gp= " & item.ToString
                query3 += " or [exam].gp= " & item.ToString
                query4 += " or [affect].gp= " & item.ToString
            End If

        Next

        Try
            'cmd = New OleDbCommand("delete [local].*,affect.*,exam.*,emploi.* from [local],[affect],[exam],[emploi] where emploi.gp= exam.gp and [local].gp=affect.gp and ( " + query + ") and [local].gp=exam.gp", con)
            cmd = New OleDbCommand("delete * from [local]  " + query, con)
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            cmd = New OleDbCommand("delete * from [emploi]  " + query2, con)
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            cmd = New OleDbCommand("delete * from [exam]  " + query3, con)
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            cmd = New OleDbCommand("delete * from [affect]   " + query4, con)
            con.Open()
            cmd.ExecuteNonQuery()


        Catch ex As Exception
            MsgBox(ex.Message)

            Exit Sub
        Finally
            con.Close()
        End Try

        arr.Clear()
        dt.Rows.Clear()
        dt.Columns.Clear()
        dt2.Rows.Clear()
        dt2.Columns.Clear()
        dt3.Rows.Clear()
        dt3.Columns.Clear()
        dt4.Rows.Clear()
        dt4.Columns.Clear()
        dt6.Rows.Clear()
        dt6.Columns.Clear()
        modules.Items.Clear()
        dgvsalles.Columns.Clear()
        dgvsalles.Rows.Clear()


    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) ''''''''''*******************'''''''''''''''''''''

        pnlAffectation2.Visible = False
        pnlAffectation2.Visible = True

        net()
        filldgv()
    End Sub

    Private Sub form1_closing() Handles Me.FormClosing ''''''*******************'''''''''''''''
        If (pnlAffectation1.Visible = False And pnlAffectation3.Visible = False And pnlAffectation2.Visible = True) Then
            Try
                ' My.Computer.FileSystem.DeleteFile("../../BDD/" + nameBdd)
                reinitialiser_var_glob()
                For Each item In listepromo.CheckedItems
                    dt.Rows.Clear()
                    dt.Columns.Clear()

                    da = New OleDbDataAdapter("select emploi.gp from [emploi],[MODULES],[exam] where emploi.gp=exam.gp and [exam.type_exam]='" + examende.Text.ToString + "' and [exam.semestre]='" + semestre.Text.ToString + "' and [emploi.module]= [MODULES.Code_Mat]  and [MODULES.Niveau]='" + item.ToString + "'", con)
                    da.Fill(dt)
                    If (dt.Rows.Count <= 0) Then
                        Exit Sub
                    End If
                    Dim query As String = Nothing
                    Dim query2 As String = Nothing
                    Dim query3 As String = Nothing
                    Dim query4 As String = Nothing
                    For Each row In dt.Rows
                        If (query = String.Empty) Then
                            query = "[local].gp = " & row(0).ToString
                            query2 = "[emploi].gp = " & row(0).ToString
                            query3 = "[exam].gp = " & row(0).ToString
                            query4 = "[affect].gp = " & row(0).ToString
                        Else
                            query += " or [local].gp= " & row(0).ToString
                            query2 += " or [emploi].gp= " & row(0).ToString
                            query3 += " or [exam].gp= " & row(0).ToString
                            query4 += " or [affect].gp= " & row(0).ToString
                        End If

                    Next

                    'cmd = New OleDbCommand("delete [local].*,affect.*,exam.*,emploi.* from [local],[affect],[exam],[emploi] where emploi.gp= exam.gp and [local].gp=affect.gp and ( " + query + ") and [local].gp=exam.gp", con)
                    cmd = New OleDbCommand("delete * from [local] where  " + query, con)
                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()
                    cmd = New OleDbCommand("delete * from [emploi] where  " + query2, con)
                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()
                    cmd = New OleDbCommand("delete * from [exam] where  " + query3, con)
                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()
                    cmd = New OleDbCommand("delete * from [affect] where  " + query4, con)
                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()
                Next
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub
            Finally
                con.Close()
            End Try
        End If


    End Sub

    Private Sub Button7_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles parcourir.MouseMove ''''*******'''''''
        parcourir.BackColor = Color.Yellow
    End Sub

    Private Sub Button7_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles parcourir.MouseLeave '''''******''''
        parcourir.BackColor = Color.Snow

    End Sub

    Private Sub dgvfianle_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvfianle.RowHeaderMouseClick ''''''''''''*******''''''''''''

        For Each item In buttutil
            If (item.BackColor = Color.Blue) Then
                item.BackColor = Color.Red
            End If

        Next
        For Each but In buttutil

            If (but.Text = dgvfianle.SelectedRows(0).Cells(3).Value.ToString) Then

                but.BackColor = Color.Blue


                dgvfianle.CurrentRow.Selected = False
                Exit Sub
            End If
        Next

    End Sub

    Private Sub cmbRes1Promo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRes1Promo.SelectedIndexChanged


        FillSemestre()
    End Sub

    Private Sub cmbRes1Semestre_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRes1Semestre.SelectedIndexChanged
        annee = cmbRest1Annee.Text + ".accdb"
        connexionNihad = "Provider=Microsoft.ace.OleDb.12.0;DATA SOURCE = ..\..\Sauvegarde\" + annee
        annee = cmbRest1Annee.Text
        FillExam()
    End Sub

    Private Sub cmbRest1TpeExam_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRest1TpeExam.SelectedIndexChanged

        annee = cmbRest1Annee.Text + ".accdb"
        connexionNihad = "Provider=Microsoft.ace.OleDb.12.0;DATA SOURCE = ..\..\Sauvegarde\" + annee
        cont = New OleDbConnection(connexionNihad)
        FillModule()
    End Sub

    Private Sub dtpAnneEtatSalle_ValueChanged(sender As Object, e As EventArgs) Handles dtpAnneEtatSalle.ValueChanged

        dtSalle.Rows.Clear()
        dtSalle.Columns.Clear()

        pathEtatSalle = "../../Sauvegarde/" & dtpAnneEtatSalle.Text.ToString & ".accdb"
        If (File.Exists(pathEtatSalle)) Then

            'Connexion vers la base de donnee pour recuperer les salles.

            connAccess = New OleDbConnection("provider = Microsoft.Ace.OLEDB.12.0; DATA SOURCE =" & pathEtatSalle)
            daSalle = New OleDbDataAdapter("SELECT DISTINCT [local].[salle] FROM [local]", connAccess)
            daSalle.Fill(dtSalle)
            cmbSalle.DataSource = dtSalle
            connAccess.Close()
        Else
            MsgBox("La base de donnee n'existe pas")
        End If
    End Sub

    Private Sub btnSuivEtat_Click(sender As Object, e As EventArgs) Handles btnSuivEtat.Click

        If (cmbCritere.Text = "Matiere/Salle") Then
            btnSalleSuite.Visible = True
            btnCrystalReports.Visible = False
        Else
            btnSalleSuite.Visible = False
            btnCrystalReports.Visible = True
        End If

        If (cmbCritere.Text = "Matiere/Salle") Then

            pnlIfMatiereSalle.Visible = True
            pnlMatricule.Visible = False

            dtt.Rows.Clear()
            dtt.Columns.Clear()
            Dim req As String = "SELECT DISTINCT [emploi].[module] FROM emploi"
            Requette(req)
            cmbMatiereSalle.DataSource = dtt
        Else
            If (cmbCritere.Text = "Matricule") Then

                pnlIfMatiereSalle.Visible = False
                pnlMatricule.Visible = True

                dtt.Rows.Clear()
                dtt.Columns.Clear()

                DtMatiere.Rows.Clear()
                DtMatiere.Columns.Clear()

                Dim req As String = "SELECT DISTINCT [affect].[matricule] FROM affect"
                DaMatiere = New OleDbDataAdapter(req, connAccess)
                DaMatiere.Fill(DtMatiere)
                cmbEtudiant.DataSource = DtMatiere
            End If
        End If

        btnCrystalReports.Enabled = True
    End Sub

    Private Sub cmbCritere_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCritere.SelectedIndexChanged

        btnSuivEtat.Visible = True
    End Sub

    Private Sub Button4_Click_2(sender As Object, e As EventArgs) Handles Button4.Click

        annee = cmbRest1Annee.Text + ".accdb"
        If (test_bdd() = 0) Then
            cmbRes1Promo.Text = ""
            cmbRes1Semestre.Text = ""
            cmbRest1TpeExam.Text = ""
            cmbRest1Module.Text = ""
            Exit Sub
        End If

        connexionNihad = "Provider=Microsoft.ace.OleDb.12.0;DATA SOURCE = ..\..\Sauvegarde\" + annee

        cont = New OleDbConnection("Provider=Microsoft.ace.OleDb.12.0;DATA SOURCE = ..\..\Sauvegarde\" + annee)
        FillPromo()
    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        MinFenetre()
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        MaxFenetre()
    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        ExitFenetre()
    End Sub

    Private Sub Button12_Click_1(sender As Object, e As EventArgs) Handles Button12.Click
        MinFenetre()
    End Sub

    Private Sub Button11_Click_1(sender As Object, e As EventArgs) Handles Button11.Click
        MaxFenetre()
    End Sub

    Private Sub Button13_Click_1(sender As Object, e As EventArgs) Handles Button13.Click
        ExitFenetre()
    End Sub

    Private Sub Button15_Click_1(sender As Object, e As EventArgs) Handles Button15.Click
        MinFenetre()
    End Sub

    Private Sub Button14_Click_1(sender As Object, e As EventArgs) Handles Button14.Click
        MaxFenetre()
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        ExitFenetre()
    End Sub

    Private Sub dteYear_ValueChanged(sender As Object, e As EventArgs) Handles dteYear.ValueChanged
        pathEtatSalle = "../../Sauvegarde/" & dteYear.Text.ToString & ".accdb"
        If (File.Exists(pathEtatSalle)) Then

            'Connexion vers la base de donnee pour recuperer les salles.

            connAccess = New OleDbConnection("provider = Microsoft.Ace.OLEDB.12.0; DATA SOURCE =" & pathEtatSalle)
        Else
            MsgBox("La base de donnee n'existe pas")
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Label2.Visible = False
        critere.Visible = False
        critere.Text = "Promo"
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Label2.Visible = True
        critere.Visible = True
    End Sub

    Private Sub modules_MouseHover(sender As Object, e As EventArgs) Handles modules.MouseHover
        modules.Height = modules.Size.Height * modules.Items.Count
    End Sub

    Private Sub modules_MouseLeave(sender As Object, e As EventArgs) Handles modules.MouseLeave
        modules.Height = 25
    End Sub
End Class
