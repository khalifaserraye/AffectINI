Imports System.ComponentModel
Imports System.Data.OleDb

Public Class Form3
    Dim nameBDD As String
    Private cnxstr As String
    Dim con As New OleDbConnection
    Dim i As Integer
    Dim rowIndex As Integer = 0
    Dim ab, cd, ef, gh, ij, kl As Boolean
    Dim Varc() As String
    Dim chaine, matricule, nome, pre, groupe, promo As String
    Dim dtTemp As New DataTable
    Dim c1 As Boolean = False
    Dim c2 As Boolean = False
    Dim c3 As Boolean = False
    Dim c4 As Boolean = False
    Dim c5 As Boolean = False
    Dim c6 As Boolean = False
    Dim LastClicked As Button
    Dim BeforeLastClicked As Button
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        nameBDD = namb.Text.ToString
        cnxstr = "Provider=Microsoft.ace.OleDb.12.0; Data Source=../../Sauvegarde/" + nameBDD
        con = New OleDbConnection(cnxstr)
        BeforeLastClicked = Button1
        Panel1.AllowDrop = True
        Panel2.AllowDrop = True
        Panel3.AllowDrop = True
        Panel4.AllowDrop = True
        Panel5.AllowDrop = True
        ComboBox2.Enabled = False
        ComboBox1.Enabled = False
        Button8011.Visible = False
        Label166.Visible = False
        Label165.Visible = False
        Label164.Visible = False
        Label163.Visible = False
        Button8012.Visible = False
        Button8014.Visible = False
        Button8013.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel1.Visible = False
        Panel3.Visible = False
        Panel11.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        Panel2.Visible = False
        Panel6.Visible = False
        Panel7.Visible = False
        Panel8.Visible = False
        Panel10.Visible = False
        Panel9.Visible = False
    End Sub

    Private Sub DataGridView1_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDown
        Dim selecttext As String
        i = e.RowIndex
        Panel10.AllowDrop = True
        Panel1.AllowDrop = True
        Panel2.AllowDrop = True
        Panel3.AllowDrop = True
        Panel4.AllowDrop = True
        Panel5.AllowDrop = True
        Panel6.AllowDrop = True
        Panel7.AllowDrop = True
        Panel9.AllowDrop = True
        Panel8.AllowDrop = True
        Panel12.AllowDrop = True
        Panel13.AllowDrop = True
        Panel11.AllowDrop = True
        tester()
        If i <> -1 Then
            selecttext = DataGridView1.Rows(i).Cells(0).Value + "-" + DataGridView1.Rows(i).Cells(1).Value + "-" + DataGridView1.Rows(i).Cells(2).Value + "-" + DataGridView1.Rows(i).Cells(3).Value
            DataGridView1.DoDragDrop(selecttext, DragDropEffects.Copy)
            ChangerModules.Visible = True
        End If
        Dim attmax As Boolean = False
        If attmax = True Then
        End If
        If MaxSalle() = True Then
            sauvegarde()
            ComboBox1.Items.Remove(ComboBox1.Text)
            ComboBox1.Text = ""
            DataGridView1.Enabled = False
            If ComboBox1.Items.Count = 0 Then
                ComboBox2.Items.Remove(ComboBox2.Text)
                ComboBox2.Text = ""
                If ComboBox2.Items.Count = 0 Then
                    MsgBox("Vous avez terminé avec cette promo.")
                    LastClicked.Enabled = False
                    LastClicked.Font = New Font("Copperplate Gothic Bold", 12.5, FontStyle.Bold)
                Else
                    MsgBox("Vous avez terminé avec ce groupe de modules.")
                End If
            Else
                MsgBox("Vous avez atteint le max, veuillez choisir un autre local.")
            End If
            viderpannaux()
        End If
    End Sub

    Private Function DonnerPos(ByVal b As Button) As String
        Dim Pos As String = ""
        Dim TabBtn() As Button = {Button145, Button146, Button147, Button148, Button139, Button140, Button141, Button142, Button143, Button144, Button71, Button72, Button73, Button74, Button75, Button70, Button76, Button77, Button78, Button79, Button8006, Button8010, Button8009, Button8008, Button8007, Button8005, Button8001, Button8002, Button8003, Button8004, Button63, Button64, Button65, Button66, Button81, Button82, Button260, Button27, Button62, Button80, Button83, Button30, Button84, Button67, Button68, Button61, Button85, Button86, Button6006, Button6007, Button6002, Button6001, Button6000, Button6003, Button6004, Button6005, Button1997, Button1998, Button1995, Button1996, Button1993, Button1994, Button1999, Button1992, Button1000, Button29, Button28, Button25, Button240, Button300, Button200, Button69, Button109, Button110, Button106, Button105, Button104, Button103, Button101, Button102, Button113, Button114, Button115, Button116, Button117, Button118, Button119, Button120, Button121, Button122, Button8014, Button8011, Button8013, Button8012, Button87, Button88, Button89, Button90, Button91, Button92, Button93, Button94, Button95, Button96, Button97, Button98, Button99, Button100, Button107, Button108, Button111, Button112, Button123, Button124, Button125, Button126, Button127, Button128, Button129, Button130, Button131, Button132, Button133, Button134, Button135, Button136, Button137, Button138, button52, button53, button54, button55, button56, button57, button58, button59, button60, button51, button42, button43, button44, button45, button46, button47, button48, button49, button50, button41, button32, button33, button34, button35, button36, button37, button38, button39, button40, button31, button15, button16, button17, button18, button19, button20, button21, button22, button23, button14, button5, button6, button7, button8, button9, button10, button11, button12, button13, Button4}
        Dim tabLab() As Label = {Label127, Label127, Label128, Label128, Label124, Label125, Label125, Label126, Label126, Label124, Label43, Label41, Label44, Label42, Label45, Label40, Label36, Label38, Label37, Label39, Label157, Label155, Label156, Label151, Label152, Label158, Label161, Label162, Label159, Label160, Label66, Label64, Label63, Label62, Label68, Label65, Label77, Label72, Label59, Label61, Label58, Label54, Label60, Label53, Label67, Label55, Label56, Label57, Label154, Label153, Label147, Label148, Label141, Label144, Label142, Label145, Label150, Label149, Label133, Label139, Label137, Label138, Label132, Label140, Label134, Label136, Label76, Label78, Label73, Label74, Label135, Label75, Label71, Label70, Label90, Label91, Label93, Label69, Label86, Label92, Label98, Label107, Label121, Label108, Label106, Label104, Label102, Label120, Label103, Label113, Label163, Label166, Label164, Label165, Label51, Label52, Label46, Label80, Label85, Label81, Label79, Label50, Label48, Label84, Label49, Label82, Label47, Label83, Label88, Label96, Label87, Label95, Label122, Label101, Label119, Label100, Label109, Label115, Label99, Label118, Label116, Label110, Label117, Label111, Label114, Label105, Label112, Label123, label4, label3, label3, label2, label2, label15, label15, label14, label14, label4, label9, label8, label8, label7, label7, label6, label6, label5, label5, label9, label1, label13, label13, label12, label12, label11, label11, label10, label10, label1, Label26, Label23, Label27, Label24, Label28, Label25, Label29, label20, Label30, Label22, Label35, label18, Label31, label17, Label32, label16, Label33, Label21, Label34, label19}
        For indice = 0 To 177
            If b.Name = TabBtn(indice).Name Then
                Pos = tabLab(indice).Text
                Exit For
            End If
        Next
        Return Pos
    End Function

    Private Sub AjouterAlaBDD(ByVal m As String, ByVal p As Integer)
        con.Close()
        con.ConnectionString = cnxstr
        con.Open()
        Dim ch As String
        ch = "insert into affect([matricule],[gp],[local],[pos]) values (?,?,?,?)"
        Dim cmd As OleDbCommand = New OleDbCommand(ch, con)
        cmd.Parameters.Add(New OleDbParameter("matricule", CType(m, String)))
        cmd.Parameters.Add(New OleDbParameter("gp", CType(Modules2mmGpToNbrGp(ComboBox2.Text), Integer)))
        cmd.Parameters.Add(New OleDbParameter("local", CType(ComboBox1.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("pos", CType(p, Integer)))
        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            con.Close()
        Catch ex As Exception
            MsgBox("Sauvegarde deux fois!!")
        End Try
    End Sub

    Private Sub sauvegarde()
        Dim pan As Panel
        Dim but As Button
        Dim indice1 As Integer = 0
        Dim TabPan124() As Panel = {Panel1, Panel2, Panel4}
        For indice1 = 0 To 2
            For Each pan In TabPan124(indice1).Controls.OfType(Of Panel)
                If pan.Visible = True Then
                    For Each but In pan.Controls.OfType(Of Button)
                        If but.Text <> "" Then
                            AjouterAlaBDD(but.Text.Substring(0, 7), DonnerPos(but))
                        End If
                    Next
                End If
            Next
        Next
        For Each pan In Me.Controls.OfType(Of Panel)
            If (pan.Name <> "'Panel14") And (pan.Visible = True) Then
                For Each but In pan.Controls.OfType(Of Button)
                    If but.Text <> "" Then
                        AjouterAlaBDD(but.Text.Substring(0, 7), DonnerPos(but))
                    End If
                Next
            End If
        Next
    End Sub

    Private Function MaxSalle() As Boolean
        Dim daMax As New OleDbDataAdapter("SELECT distinct * from [local] where [local].[salle] = '" + ComboBox1.Text + "' ", con)
        Dim dtMax As New DataTable
        daMax.Fill(dtMax)
        Dim pan As Panel
        Dim but As Button
        Dim bool As Boolean = False
        Dim indice1 As Integer = 0
        Dim compteur As Integer = 0
        Dim TabPan124() As Panel = {Panel1, Panel2, Panel4}
        For indice1 = 0 To 2
            For Each pan In TabPan124(indice1).Controls.OfType(Of Panel)
                If pan.Visible = True Then
                    For Each but In pan.Controls.OfType(Of Button)
                        If but.Text <> "" Then
                            compteur = compteur + 1
                        End If
                    Next
                End If
            Next
        Next
        For Each pan In Me.Controls.OfType(Of Panel)
            If (pan.Name <> "'Panel14") And (pan.Visible = True) Then
                For Each but In pan.Controls.OfType(Of Button)
                    If but.Text <> "" Then
                        compteur = compteur + 1
                    End If
                Next
            End If
        Next
        For indice1 = 0 To dtMax.Rows.Count - 1
            If dtMax.Rows(indice1)(2) = Modules2mmGpToNbrGp(ComboBox2.Text) Then
                If compteur >= dtMax.Rows(indice1)(1) Then
                    bool = True
                End If
                Exit For
            End If
        Next
        Return bool
    End Function

    Private Function Nbr2GpToModules2mmGp(ByVal Nbr2Gp As Integer) As String
        Dim Modules2mmGp As String = ""
        Dim CommandString As String = "SELECT distinct [emploi.module], [emploi.gp] from emploi"
        Dim da As New OleDbDataAdapter(CommandString, con)
        Dim dt As New DataTable
        da.Fill(dt)
        For i = 0 To dt.Rows.Count - 1
            If dt.Rows(i).Item(1) = Nbr2Gp Then
                Modules2mmGp = Modules2mmGp & dt.Rows(i).Item(0) & ", "
            End If
        Next
        If Modules2mmGp <> "" Then
            Modules2mmGp = Modules2mmGp.Substring(0, Modules2mmGp.Length - 2)
        End If
        Return Modules2mmGp
    End Function

    Private Function Modules2mmGpToNbrGp(ByVal Modules2mmGp As String) As Integer
        Dim daGp1 As New OleDbDataAdapter("select [emploi.module], [emploi.gp] from [emploi], [exam] where [emploi].[gp] = [exam].[gp] and [exam].[type_exam] ='" + exm.Text.ToString + "' and [exam].[semestre] = '" + sem.Text.ToString + "' ", con)
        Dim dtGp1 As New DataTable
        Dim daGp2 As New OleDbDataAdapter("SELECT [local].[salle], [local].[gp] from [local]", con)
        Dim dtGp2 As New DataTable
        daGp1.Fill(dtGp1)
        daGp2.Fill(dtGp2)
        Dim cpt, NbrGp As Integer
        cpt = 0
        For i = 0 To Modules2mmGp.Length - 1
            If Modules2mmGp(i) = "," Then
                Exit For
            End If
            cpt = cpt + 1
        Next
        If Modules2mmGp <> Nothing Then
            Modules2mmGp = Modules2mmGp.Substring(0, cpt)
            For i = 0 To dtGp2.Rows.Count - 1
                If Modules2mmGp = dtGp1.Rows(i).Item(0) Then
                    NbrGp = dtGp1.Rows(i).Item(1)
                    Exit For
                End If
            Next
        End If
        Return NbrGp
    End Function

    Private Sub ClickPromo(sender As Object, e As EventArgs) Handles CP1.Click, CP2.Click, CS1.Click, SQ2.Click, SL2.Click, ST2.Click
        Dim Niv As Button = CType(sender, Button)
        Dim daPromo As New OleDbDataAdapter("SELECT Matricule,NomEtud,Prenoms,Gr from ETUDIANTS where Gr <> '""' and  Promo = '" + Niv.Text + "' order by [Matricule] ", con)
        Dim daGp As New OleDbDataAdapter("select distinct [emploi.gp] from [emploi], [MODULES], [exam] where [emploi].[gp] = [exam].[gp] and [exam].[type_exam] ='" + exm.Text.ToString + "' and [exam].[semestre] = '" + sem.Text.ToString + "'  and [MODULES].[Niveau] = '" + Niv.Text + "'  and [emploi].[module] = [MODULES].[Code_Mat]  ", con)
        Dim dtPromo As New DataTable
        Dim dtGp As New DataTable
        daPromo.Fill(dtPromo)
        daGp.Fill(dtGp)
        BeforeLastClicked.Font = New Font("Copperplate Gothic Bold", 12.5)
        BeforeLastClicked.ForeColor = Color.FromArgb(14, 28, 53)
        LastClicked = Niv
        Niv.Font = New Font("Copperplate Gothic Bold", 12.5, FontStyle.Underline Or FontStyle.Bold)
        Niv.ForeColor = Color.Red
        BeforeLastClicked = LastClicked
        DataGridView1.DataSource = dtPromo
        ComboBox1.Items.Clear()
        ComboBox2.Text = Nothing
        ComboBox2.Items.Clear()
        ComboBox2.Enabled = True
        ComboBox2.Text = Nothing
        viderpannaux()
        Dim iGp, s As Integer
        For iGp = 0 To dtGp.Rows.Count - 1
            s = dtGp.Rows(iGp)(0)
            ComboBox2.Items.Add(Nbr2GpToModules2mmGp(s))
        Next
    End Sub

    Sub viderpannaux()
        Dim pan As Panel
        Dim but As Button
        Dim indice1 As Integer = 0
        Dim TabPan124() As Panel = {Panel1, Panel2, Panel4}
        For indice1 = 0 To 2
            For Each pan In TabPan124(indice1).Controls.OfType(Of Panel)
                For Each but In pan.Controls.OfType(Of Button)
                    but.Text = ""
                    but.BackColor = Color.Maroon
                    but.AllowDrop = True
                Next
            Next
        Next
        For Each pan In Me.Controls.OfType(Of Panel)
            If pan.Name <> "'Panel14" Then
                For Each but In pan.Controls.OfType(Of Button)
                    but.Text = ""
                    but.BackColor = Color.Maroon
                    but.AllowDrop = True
                Next
            End If
        Next
    End Sub

    Sub Ajouter_etudiant()
        Dim ro As DataRow = dtTemp.NewRow
        ro(0) = matricule
        ro(1) = nome
        ro(2) = pre
        ro(3) = promo
        dtTemp.Rows.Add(ro)
        dtTemp = DataGridView1.DataSource
        DataGridView1.DataSource = dtTemp
    End Sub

    Private Sub tester()
        If (ComboBox2.Text = String.Empty) Then
            MsgBox("il faut choisir un groupe de modules", MsgBoxStyle.Critical, "attention")
        Else
            If (ComboBox1.Text = String.Empty) Then
                MsgBox("il faut choisir au moins une salle", MsgBoxStyle.Critical, "attention")
            End If
        End If
    End Sub

    Private Sub ComboBox1AND2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress, ComboBox2.KeyPress
        e.KeyChar = String.Empty
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        ComboBox1.Enabled = True
        ChangerModules.Enabled = True
        Dim daLocal As New OleDbDataAdapter("SELECT distinct [salle], gp from [local]", con)
        Dim dtLocal As New DataTable
        daLocal.Fill(dtLocal)
        Dim iLocal As Integer
        ComboBox1.Text = ""
        ComboBox1.Items.Clear()
        For iLocal = 0 To dtLocal.Rows.Count - 1
            If dtLocal.Rows(iLocal).Item(1).ToString = Modules2mmGpToNbrGp(ComboBox2.Text) Then
                ComboBox1.Items.Add(dtLocal.Rows(iLocal).Item(0))
            End If
        Next
    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        viderpannaux()
        DataGridView1.Enabled = True
        ChangerLocal.Enabled = True
        If ComboBox1.Text = "CP2" Or ComboBox1.Text = "S4B" Or ComboBox1.Text = "DPG" Or ComboBox1.Text = "SIS" Or ComboBox1.Text = "BP" Or ComboBox1.Text = "CP1" Or ComboBox1.Text = "CP3" Or ComboBox1.Text = "CP4" Or ComboBox1.Text = "CP5" Or ComboBox1.Text = "CP6" Or ComboBox1.Text = "CP7" Or ComboBox1.Text = "CP8" Or ComboBox1.Text = "CP9" Or ComboBox1.Text = "M1" Or ComboBox1.Text = "M2" Or ComboBox1.Text = "M3" Or ComboBox1.Text = "M4" Or ComboBox1.Text = "M5" Or ComboBox1.Text = "M6" Or ComboBox1.Text = "M7" Or ComboBox1.Text = "M8" Or ComboBox1.Text = "Me" Or ComboBox1.Text = "MH" Then
            Panel1.Visible = True
            Panel2.Visible = False
            Panel3.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel13.Visible = False
            Panel12.Visible = False
            Button8011.Visible = False
            Button8012.Visible = False
            Button8014.Visible = False
            Button8013.Visible = False
            Label166.Visible = False
            Label165.Visible = False
            Label164.Visible = False
            Label163.Visible = False
        ElseIf ComboBox1.Text = "A3" Then
            Panel13.Visible = False
            Panel12.Visible = False
            Button8011.Visible = False
            Button8012.Visible = False
            Button8014.Visible = False
            Button8013.Visible = False
            Label166.Visible = False
            Label165.Visible = False
            Label164.Visible = False
            Label163.Visible = False
            Panel4.Visible = True
            Panel11.Visible = True
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel1.Visible = False
            Panel10.Visible = False
            Label47.Text = 33
            Label83.Text = 34
            Label84.Text = 31
            Label48.Text = 32
            Label82.Text = 30
            Label49.Text = 29
            Label85.Text = 28
            Label50.Text = 27
            Label79.Text = 26
            Label81.Text = 25
            Label80.Text = 24
            Label46.Text = 23
            Label51.Text = 22
            Label52.Text = 21
            Label95.Text = 20
            Label87.Text = 19
            Label96.Text = 18
            Label88.Text = 17
            Label58.Text = 35
            Label60.Text = 36
            Label59.Text = 37
            Label61.Text = 38
            Label68.Text = 39
            Label65.Text = 40
            Label62.Text = 41
            Label66.Text = 42
            Label64.Text = 43
            Label63.Text = 44
            Label67.Text = 45
            Label55.Text = 46
            Label53.Text = 47
            Label54.Text = 48
            Label56.Text = 49
            Label57.Text = 50
            Label77.Text = 51
            Label72.Text = 52
            Label114.Text = 1
            Label105.Text = 2
            Label112.Text = 3
            Label123.Text = 4
            Label111.Text = 5
            Label116.Text = 6
            Label110.Text = 7
            Label117.Text = 8
            Label109.Text = 9
            Label115.Text = 10
            Label99.Text = 11
            Label118.Text = 12
            Label100.Text = 13
            Label122.Text = 14
            Label101.Text = 15
            Label119.Text = 16
        ElseIf ComboBox1.Text = "A4" Then
            Panel4.Visible = True
            Panel10.Visible = False
            Panel9.Visible = False
            Panel6.Visible = True
            Panel7.Visible = True
            Panel8.Visible = True
            Panel11.Visible = True
            Panel13.Visible = False
            Panel12.Visible = False
            Button8011.Visible = False
            Button8012.Visible = False
            Button8014.Visible = False
            Button8013.Visible = False
            Label166.Visible = False
            Label165.Visible = False
            Label164.Visible = False
            Label163.Visible = False
            Label114.Text = 1
            Label105.Text = 2
            Label112.Text = 3
            Label123.Text = 4
            Label111.Text = 5
            Label116.Text = 6
            Label110.Text = 7
            Label117.Text = 8
            Label109.Text = 9
            Label115.Text = 10
            Label99.Text = 11
            Label118.Text = 12
            Label100.Text = 13
            Label122.Text = 14
            Label101.Text = 15
            Label119.Text = 16
            Label102.Text = 17
            Label120.Text = 18
            Label103.Text = 19
            Label113.Text = 20
            Label104.Text = 21
            Label121.Text = 22
            Label108.Text = 23
            Label106.Text = 24
            Label98.Text = 25
            Label107.Text = 26
            Label47.Text = 27
            Label83.Text = 28
            Label48.Text = 29
            Label84.Text = 30
            Label49.Text = 31
            Label82.Text = 32
            Label50.Text = 33
            Label85.Text = 34
            Label81.Text = 35
            Label79.Text = 36
            Label46.Text = 37
            Label80.Text = 38
            Label52.Text = 39
            Label51.Text = 40
            Label87.Text = 41
            Label95.Text = 42
            Label88.Text = 43
            Label96.Text = 44
            Label71.Text = 45
            Label70.Text = 46
            Label90.Text = 47
            Label69.Text = 48
            Label93.Text = 49
            Label91.Text = 50
            Label86.Text = 51
            Label92.Text = 52
            Label58.Text = 53
            Label60.Text = 54
            Label59.Text = 55
            Label61.Text = 56
            Label68.Text = 57
            Label65.Text = 58
            Label62.Text = 59
            Label66.Text = 60
            Label64.Text = 61
            Label63.Text = 62
            Label67.Text = 63
            Label55.Text = 64
            Label53.Text = 65
            Label54.Text = 66
            Label56.Text = 67
            Label57.Text = 68
            Label77.Text = 69
            Label72.Text = 70
        ElseIf ComboBox1.Text = "CYB" Or ComboBox1.Text = "BIB" Then
            Panel1.Visible = True
            Panel5.Visible = True
            Panel13.Visible = False
            Panel12.Visible = False
            Button8011.Visible = False
            Button8012.Visible = False
            Button8014.Visible = False
            Button8013.Visible = False
            Label166.Visible = False
            Label165.Visible = False
            Label164.Visible = False
            Label163.Visible = False
        ElseIf ComboBox1.Text = "S04" Or ComboBox1.Text = "S05" Or ComboBox1.Text = "S06" Or ComboBox1.Text = "S07" Or ComboBox1.Text = "S08" Or ComboBox1.Text = "S09" Or ComboBox1.Text = "S10" Or ComboBox1.Text = "S11" Or ComboBox1.Text = "S12" Or ComboBox1.Text = "S13" Or ComboBox1.Text = "S14" Then
            Panel2.Visible = True
            Panel3.Visible = False
            Panel1.Visible = False
            Panel13.Visible = False
            Panel12.Visible = False
            Button8011.Visible = False
            Button8012.Visible = False
            Button8014.Visible = False
            Button8013.Visible = False
            Label166.Visible = False
            Label165.Visible = False
            Label164.Visible = False
            Label163.Visible = False
        ElseIf ComboBox1.Text = "S18" Or ComboBox1.Text = "BP" Or ComboBox1.Text = "S19" Or ComboBox1.Text = "S20" Or ComboBox1.Text = "S21" Or ComboBox1.Text = "S22" Then
            Panel3.Visible = True
            Panel2.Visible = True
            Panel13.Visible = False
            Panel12.Visible = False
            Button8011.Visible = False
            Button8012.Visible = False
            Button8014.Visible = False
            Button8013.Visible = False
            Label166.Visible = False
            Label165.Visible = False
            Label164.Visible = False
            Label163.Visible = False
        ElseIf ComboBox1.Text = "A1" Or ComboBox1.Text = "A2" Then
            Panel4.Visible = True
            Panel9.Visible = True
            Panel10.Visible = True
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel11.Visible = False
            Panel13.Visible = False
            Panel12.Visible = False
            Button8011.Visible = False
            Button8012.Visible = False
            Button8014.Visible = False
            Button8013.Visible = False
            Label166.Visible = False
            Label165.Visible = False
            Label164.Visible = False
            Label163.Visible = False
            Label114.Text = 1
            Label105.Text = 2
            Label132.Text = 3
            Label112.Text = 4
            Label123.Text = 5
            Label140.Text = 6
            Label111.Text = 7
            Label116.Text = 8
            Label137.Text = 9
            Label110.Text = 10
            Label117.Text = 11
            Label138.Text = 12
            Label109.Text = 13
            Label115.Text = 14
            Label133.Text = 15
            Label99.Text = 16
            Label118.Text = 17
            Label139.Text = 18
            Label100.Text = 19
            Label122.Text = 20
            Label150.Text = 21
            Label101.Text = 22
            Label119.Text = 23
            Label149.Text = 24
            Label88.Text = 25
            Label96.Text = 26
            Label87.Text = 27
            Label95.Text = 28
            Label153.Text = 29
            Label52.Text = 30
            Label51.Text = 31
            Label154.Text = 32
            Label46.Text = 33
            Label80.Text = 34
            Label145.Text = 35
            Label81.Text = 36
            Label79.Text = 37
            Label142.Text = 38
            Label50.Text = 39
            Label85.Text = 40
            Label144.Text = 41
            Label49.Text = 42
            Label82.Text = 43
            Label147.Text = 44
            Label48.Text = 45
            Label84.Text = 46
            Label148.Text = 47
            Label47.Text = 48
            Label83.Text = 49
            Label141.Text = 50
        ElseIf ComboBox1.Text = "AP1" Or ComboBox1.Text = "AP2" Then
            Panel4.Visible = True
            Panel9.Visible = True
            Panel10.Visible = True
            Panel8.Visible = True
            Panel12.Visible = True
            Panel6.Visible = True
            Panel13.Visible = True
            Panel1.Visible = False
            Panel2.Visible = False
            Panel3.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
            Panel11.Visible = False
            Button8011.Visible = True
            Button8012.Visible = True
            Button8014.Visible = True
            Button8013.Visible = True
            Label166.Visible = True
            Label165.Visible = True
            Label164.Visible = True
            Label163.Visible = True
            Label114.Text = 1
            Label105.Text = 2
            Label132.Text = 3
            Label112.Text = 4
            Label123.Text = 5
            Label140.Text = 6
            Label111.Text = 7
            Label116.Text = 8
            Label137.Text = 9
            Label110.Text = 10
            Label117.Text = 11
            Label138.Text = 12
            Label109.Text = 13
            Label115.Text = 14
            Label133.Text = 15
            Label99.Text = 16
            Label118.Text = 17
            Label139.Text = 18
            Label100.Text = 19
            Label122.Text = 20
            Label150.Text = 21
            Label101.Text = 22
            Label119.Text = 23
            Label149.Text = 24
            Label102.Text = 25
            Label120.Text = 26
            Label161.Text = 27
            Label103.Text = 28
            Label113.Text = 29
            Label162.Text = 30
            Label104.Text = 31
            Label121.Text = 32
            Label159.Text = 33
            Label108.Text = 34
            Label106.Text = 35
            Label160.Text = 36
            Label98.Text = 37
            Label107.Text = 38
            Label158.Text = 39
            Label86.Text = 40
            Label92.Text = 41
            Label157.Text = 42
            Label93.Text = 43
            Label91.Text = 44
            Label152.Text = 45
            Label90.Text = 46
            Label69.Text = 47
            Label151.Text = 48
            Label71.Text = 49
            Label70.Text = 50
            Label156.Text = 51
            Label88.Text = 52
            Label96.Text = 53
            Label155.Text = 54
            Label87.Text = 55
            Label95.Text = 56
            Label153.Text = 57
            Label52.Text = 58
            Label51.Text = 59
            Label154.Text = 60
            Label46.Text = 61
            Label80.Text = 62
            Label145.Text = 63
            Label81.Text = 64
            Label79.Text = 65
            Label142.Text = 66
            Label50.Text = 67
            Label85.Text = 68
            Label144.Text = 69
            Label49.Text = 70
            Label82.Text = 71
            Label147.Text = 72
            Label48.Text = 73
            Label84.Text = 74
            Label148.Text = 75
            Label47.Text = 76
            Label83.Text = 77
            Label141.Text = 78
        End If
        ComboBox1.Enabled = True
    End Sub

    Private Sub Retour_Click(sender As Object, e As EventArgs) Handles Retour.Click
        If (Not (CP1.Enabled = False And CP2.Enabled = False And CS1.Enabled = False And SQ2.Enabled = False And SL2.Enabled = False And ST2.Enabled = False)) Then
            Dim f As System.IO.File
            f.Delete("../../Sauvegarde/" + nameBDD)
        End If
        Dim frm1 As New AffectINI
        frm1.Show()
        frm1.pnlAffectation1.Visible = True
        frm1.pnlAffectation2.Visible = False
        frm1.pnlLogin.Visible = False
        frm1.view3.Visible = False
        frm1.pnlMenu.Visible = False
        Me.Close()
        Exit Sub
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Me.rowIndex = e.RowIndex
    End Sub

    Private Sub ChangerModules_Click(sender As Object, e As EventArgs) Handles ChangerModules.Click
        If MsgBox("voulez vous changer le groupe de module  ?", MsgBoxStyle.YesNo, "confirmation") = MsgBoxResult.Yes Then
            sauvegarde()
            Panel1.Visible = False
            Panel2.Visible = False
            Panel3.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox1.Items.Clear()
            ComboBox1.Text = ""
            ComboBox1.Enabled = False
            DataGridView1.Enabled = False
            ComboBox2.Items.Remove(ComboBox2.Text)
            ComboBox2.Text = ""
            If ComboBox2.Items.Count = 0 Then
                MsgBox("Vous avez terminé ce groupe de modules.")
                ComboBox1.Enabled = False
                If ComboBox1.Items.Count = 0 Then
                    MsgBox("Vous avez terminé avec cette promo.")
                    LastClicked.Enabled = False
                    LastClicked.Font = New Font("Copperplate Gothic Bold", 12.5)
                    ComboBox1.Enabled = False
                    ComboBox2.Enabled = False
                Else
                    ComboBox2.Enabled = True
                End If
            End If
            viderpannaux()
        End If
    End Sub

    Private Sub ChangerLocal_Click(sender As Object, e As EventArgs) Handles ChangerLocal.Click
        Dim ola As Integer = 0
        If MsgBox("voulez vous changer ce local ?", MsgBoxStyle.YesNo, "confirmation") = MsgBoxResult.Yes Then
            sauvegarde()
            ComboBox1.Items.Remove(ComboBox1.Text)
            ComboBox1.Text = ""
            ComboBox1.Enabled = True
            DataGridView1.Enabled = False
            If ComboBox1.Items.Count = 0 Then
                ComboBox2.Items.Remove(ComboBox2.Text)
                ComboBox2.Text = ""
                If ComboBox2.Items.Count = 0 Then
                    MsgBox("Vous avez terminé avec cette promo.")
                    LastClicked.Enabled = False
                    LastClicked.Font = New Font("Copperplate Gothic Bold", 12.5)
                    ComboBox2.Enabled = True
                End If
            End If
            viderpannaux()
        End If
    End Sub

    Private Sub ShowStudentInfo(sender As Object, e As EventArgs) Handles Button145.MouseHover, Button146.MouseHover, Button147.MouseHover, Button148.MouseHover, Button139.MouseHover, Button140.MouseHover, Button141.MouseHover, Button142.MouseHover, Button143.MouseHover, Button144.MouseHover, Button71.MouseHover, Button72.MouseHover, Button73.MouseHover, Button74.MouseHover, Button75.MouseHover, Button70.MouseHover, Button76.MouseHover, Button77.MouseHover, Button78.MouseHover, Button79.MouseHover, Button8006.MouseHover, Button8010.MouseHover, Button8009.MouseHover, Button8008.MouseHover, Button8007.MouseHover, Button8005.MouseHover, Button8001.MouseHover, Button8002.MouseHover, Button8003.MouseHover, Button8004.MouseHover, Button63.MouseHover, Button64.MouseHover, Button65.MouseHover, Button66.MouseHover, Button81.MouseHover, Button82.MouseHover, Button260.MouseHover, Button27.MouseHover, Button62.MouseHover, Button80.MouseHover, Button83.MouseHover, Button30.MouseHover, Button84.MouseHover, Button67.MouseHover, Button68.MouseHover, Button61.MouseHover, Button85.MouseHover, Button86.MouseHover, Button6006.MouseHover, Button6007.MouseHover, Button6002.MouseHover, Button6001.MouseHover, Button6000.MouseHover, Button6003.MouseHover, Button6004.MouseHover, Button6005.MouseHover, Button1997.MouseHover, Button1998.MouseHover, Button1995.MouseHover, Button1996.MouseHover, Button1993.MouseHover, Button1994.MouseHover, Button1999.MouseHover, Button1992.MouseHover, Button1000.MouseHover, Button29.MouseHover, Button28.MouseHover, Button25.MouseHover, Button240.MouseHover, Button300.MouseHover, Button200.MouseHover, Button69.MouseHover, Button109.MouseHover, Button110.MouseHover, Button106.MouseHover, Button105.MouseHover, Button104.MouseHover, Button103.MouseHover, Button101.MouseHover, Button102.MouseHover, Button113.MouseHover, Button114.MouseHover, Button115.MouseHover, Button116.MouseHover, Button117.MouseHover, Button118.MouseHover, Button119.MouseHover, Button120.MouseHover, Button121.MouseHover, Button122.MouseHover, button15.MouseHover, button16.MouseHover, button17.MouseHover, button18.MouseHover, button19.MouseHover, button20.MouseHover, button21.MouseHover, button22.MouseHover, button23.MouseHover, button14.MouseHover, button5.MouseHover, button6.MouseHover, button7.MouseHover, button8.MouseHover, button9.MouseHover, button10.MouseHover, button11.MouseHover, button12.MouseHover, button13.MouseHover, Button4.MouseHover, Button8014.MouseHover, Button8011.MouseHover, Button8013.MouseHover, Button8012.MouseHover, Button87.MouseHover, Button88.MouseHover, Button89.MouseHover, Button90.MouseHover, Button91.MouseHover, Button92.MouseHover, Button93.MouseHover, Button94.MouseHover, Button95.MouseHover, Button96.MouseHover, Button97.MouseHover, Button98.MouseHover, Button99.MouseHover, Button100.MouseHover, Button107.MouseHover, Button108.MouseHover, Button111.MouseHover, Button112.MouseHover, Button123.MouseHover, Button124.MouseHover, Button125.MouseHover, Button126.MouseHover, Button127.MouseHover, Button128.MouseHover, Button129.MouseHover, Button130.MouseHover, Button131.MouseHover, Button132.MouseHover, Button133.MouseHover, Button134.MouseHover, Button135.MouseHover, Button136.MouseHover, Button137.MouseHover, Button138.MouseHover, button52.MouseHover, button53.MouseHover, button54.MouseHover, button55.MouseHover, button56.MouseHover, button57.MouseHover, button58.MouseHover, button59.MouseHover, button60.MouseHover, button51.MouseHover, button42.MouseHover, button43.MouseHover, button44.MouseHover, button45.MouseHover, button46.MouseHover, button47.MouseHover, button48.MouseHover, button49.MouseHover, button50.MouseHover, button41.MouseHover, button32.MouseHover, button33.MouseHover, button34.MouseHover, button35.MouseHover, button36.MouseHover, button37.MouseHover, button38.MouseHover, button39.MouseHover, button40.MouseHover, button31.MouseHover
        Dim toolTip1 As New ToolTip()
        Dim bt As Button = CType(sender, Button)
        Dim tabEtudiant As String()
        If bt.Text <> "" Then
            tabEtudiant = Split(bt.Text, "-")
            toolTip1.SetToolTip(bt, "Matricule: " + tabEtudiant(0) + vbNewLine + "Nom:         " + tabEtudiant(1) + vbNewLine + "Prénom:    " + tabEtudiant(2) + vbNewLine + "Promo:      " + LastClicked.Text + vbNewLine + "Groupe:     " + tabEtudiant(3) + vbNewLine + "Position:    " + DonnerPos(bt))
        End If
    End Sub

    Private Function PositionAdjacente(ByVal b As Button) As Button
        Dim btt As New Button
        Dim TablesGauches As Button() = {Button144, Button140, Button142, Button145, Button147, button31, button33, button35, button37, button39, button41, button43, button45, button47, button49, button51, button53, button55, button57, button59}
        Dim TablesDroites As Button() = {Button139, Button141, Button143, Button146, Button148, button32, button34, button36, button38, button40, button42, button44, button46, button48, button50, button52, button54, button56, button58, button60}
        For indice = 0 To 19
            If b.Name = TablesGauches(indice).Name Then
                btt = TablesDroites(indice)
                Exit For
            ElseIf b.Name = TablesDroites(indice).Name Then
                btt = TablesGauches(indice)
                Exit For
            End If
        Next
        Return btt
    End Function

    Private Sub ViderTable1(sender As Object, e As EventArgs) Handles Button71.Click, Button72.Click, Button73.Click, Button74.Click, Button75.Click, Button70.Click, Button76.Click, Button77.Click, Button78.Click, Button79.Click, Button8006.Click, Button8010.Click, Button8009.Click, Button8008.Click, Button8007.Click, Button8005.Click, Button8001.Click, Button8002.Click, Button8003.Click, Button8004.Click, Button63.Click, Button64.Click, Button65.Click, Button66.Click, Button81.Click, Button82.Click, Button260.Click, Button27.Click, Button62.Click, Button80.Click, Button83.Click, Button30.Click, Button84.Click, Button67.Click, Button68.Click, Button61.Click, Button85.Click, Button86.Click, Button6006.Click, Button6007.Click, Button6002.Click, Button6001.Click, Button6000.Click, Button6003.Click, Button6004.Click, Button6005.Click, Button1997.Click, Button1998.Click, Button1995.Click, Button1996.Click, Button1993.Click, Button1994.Click, Button1999.Click, Button1992.Click, Button1000.Click, Button29.Click, Button28.Click, Button25.Click, Button240.Click, Button300.Click, Button200.Click, Button69.Click, Button109.Click, Button110.Click, Button106.Click, Button105.Click, Button104.Click, Button103.Click, Button101.Click, Button102.Click, Button113.Click, Button114.Click, Button115.Click, Button116.Click, Button117.Click, Button118.Click, Button119.Click, Button120.Click, Button121.Click, Button122.Click, button15.Click, button16.Click, button17.Click, button18.Click, button19.Click, button20.Click, button21.Click, button22.Click, button23.Click, button14.Click, button5.Click, button6.Click, button7.Click, button8.Click, button9.Click, button10.Click, button11.Click, button12.Click, button13.Click, Button4.Click, Button8014.Click, Button8011.Click, Button8013.Click, Button8012.Click, Button87.Click, Button88.Click, Button89.Click, Button90.Click, Button91.Click, Button92.Click, Button93.Click, Button94.Click, Button95.Click, Button96.Click, Button97.Click, Button98.Click, Button99.Click, Button100.Click, Button107.Click, Button108.Click, Button111.Click, Button112.Click, Button123.Click, Button124.Click, Button125.Click, Button126.Click, Button127.Click, Button128.Click, Button129.Click, Button130.Click, Button131.Click, Button132.Click, Button133.Click, Button134.Click, Button135.Click, Button136.Click, Button137.Click, Button138.Click
        Dim bt As Button = CType(sender, Button)
        If bt.Text <> "" Then
            If bt.Text IsNot Nothing = True Then
                chaine = bt.Text
                bt.Text = Nothing
                Varc = Split(chaine, "-")
                matricule = Varc(0)
                nome = Varc(1)
                pre = Varc(2)
                promo = Varc(3)
                bt.BackColor = Color.Maroon
                Ajouter_etudiant()
                bt.AllowDrop = True
            End If
        End If
    End Sub

    Private Sub TablesDragDrop1(sender As Object, e As DragEventArgs) Handles Button71.DragDrop, Button72.DragDrop, Button73.DragDrop, Button74.DragDrop, Button75.DragDrop, Button70.DragDrop, Button76.DragDrop, Button77.DragDrop, Button78.DragDrop, Button79.DragDrop, Button8006.DragDrop, Button8010.DragDrop, Button8009.DragDrop, Button8008.DragDrop, Button8007.DragDrop, Button8005.DragDrop, Button8001.DragDrop, Button8002.DragDrop, Button8003.DragDrop, Button8004.DragDrop, Button63.DragDrop, Button64.DragDrop, Button65.DragDrop, Button66.DragDrop, Button81.DragDrop, Button82.DragDrop, Button260.DragDrop, Button27.DragDrop, Button62.DragDrop, Button80.DragDrop, Button83.DragDrop, Button30.DragDrop, Button84.DragDrop, Button67.DragDrop, Button68.DragDrop, Button61.DragDrop, Button85.DragDrop, Button86.DragDrop, Button6006.DragDrop, Button6007.DragDrop, Button6002.DragDrop, Button6001.DragDrop, Button6000.DragDrop, Button6003.DragDrop, Button6004.DragDrop, Button6005.DragDrop, Button1997.DragDrop, Button1998.DragDrop, Button1995.DragDrop, Button1996.DragDrop, Button1993.DragDrop, Button1994.DragDrop, Button1999.DragDrop, Button1992.DragDrop, Button1000.DragDrop, Button29.DragDrop, Button28.DragDrop, Button25.DragDrop, Button240.DragDrop, Button300.DragDrop, Button200.DragDrop, Button69.DragDrop, Button109.DragDrop, Button110.DragDrop, Button106.DragDrop, Button105.DragDrop, Button104.DragDrop, Button103.DragDrop, Button101.DragDrop, Button102.DragDrop, Button113.DragDrop, Button114.DragDrop, Button115.DragDrop, Button116.DragDrop, Button117.DragDrop, Button118.DragDrop, Button119.DragDrop, Button120.DragDrop, Button121.DragDrop, Button122.DragDrop, button15.DragDrop, button16.DragDrop, button17.DragDrop, button18.DragDrop, button19.DragDrop, button20.DragDrop, button21.DragDrop, button22.DragDrop, button23.DragDrop, button14.DragDrop, button5.DragDrop, button6.DragDrop, button7.DragDrop, button8.DragDrop, button9.DragDrop, button10.DragDrop, button11.DragDrop, button12.DragDrop, button13.DragDrop, Button4.DragDrop, Button8014.DragDrop, Button8011.DragDrop, Button8013.DragDrop, Button8012.DragDrop, Button87.DragDrop, Button88.DragDrop, Button89.DragDrop, Button90.DragDrop, Button91.DragDrop, Button92.DragDrop, Button93.DragDrop, Button94.DragDrop, Button95.DragDrop, Button96.DragDrop, Button97.DragDrop, Button98.DragDrop, Button99.DragDrop, Button100.DragDrop, Button107.DragDrop, Button108.DragDrop, Button111.DragDrop, Button112.DragDrop, Button123.DragDrop, Button124.DragDrop, Button125.DragDrop, Button126.DragDrop, Button127.DragDrop, Button128.DragDrop, Button129.DragDrop, Button130.DragDrop, Button131.DragDrop, Button132.DragDrop, Button133.DragDrop, Button134.DragDrop, Button135.DragDrop, Button136.DragDrop, Button137.DragDrop, Button138.DragDrop
        Dim bt As Button = CType(sender, Button)
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        If bt.Text = Nothing Then
            bt.BackColor = Color.White
            bt.Text = e.Data.GetData(DataFormats.Text)
            DataGridView1.Rows.RemoveAt(i)
            dtTemp = DataGridView1.DataSource
            bt.AllowDrop = False
        End If
    End Sub

    Private Sub ViderTable2(sender As Object, e As EventArgs) Handles button52.Click, button53.Click, button54.Click, button55.Click, button56.Click, button57.Click, button58.Click, button59.Click, button60.Click, button51.Click, button42.Click, button43.Click, button44.Click, button45.Click, button46.Click, button47.Click, button48.Click, button49.Click, button50.Click, button41.Click, button32.Click, button33.Click, button34.Click, button35.Click, button36.Click, button37.Click, button38.Click, button39.Click, button40.Click, button31.Click, Button145.Click, Button146.Click, Button147.Click, Button148.Click, Button139.Click, Button140.Click, Button141.Click, Button142.Click, Button143.Click, Button144.Click
        Dim bt As Button = CType(sender, Button)
        If PositionAdjacente(bt).Text = Nothing Then
            bt.AllowDrop = True
            PositionAdjacente(bt).AllowDrop = True
        End If
        If bt.Text <> "" Then
            If bt.Text IsNot Nothing = True Then
                chaine = bt.Text
                bt.Text = Nothing
                Varc = Split(chaine, "-")
                matricule = Varc(0)
                nome = Varc(1)
                pre = Varc(2)
                promo = Varc(3)
                bt.BackColor = Color.Maroon
                Ajouter_etudiant()
            End If
        End If
    End Sub

    Private Sub TablesDragDrop2(sender As Object, e As DragEventArgs) Handles button52.DragDrop, button53.DragDrop, button54.DragDrop, button55.DragDrop, button56.DragDrop, button57.DragDrop, button58.DragDrop, button59.DragDrop, button60.DragDrop, button51.DragDrop, button42.DragDrop, button43.DragDrop, button44.DragDrop, button45.DragDrop, button46.DragDrop, button47.DragDrop, button48.DragDrop, button49.DragDrop, button50.DragDrop, button41.DragDrop, button32.DragDrop, button33.DragDrop, button34.DragDrop, button35.DragDrop, button36.DragDrop, button37.DragDrop, button38.DragDrop, button39.DragDrop, button40.DragDrop, button31.DragDrop, Button145.DragDrop, Button146.DragDrop, Button147.DragDrop, Button148.DragDrop, Button139.DragDrop, Button140.DragDrop, Button141.DragDrop, Button142.DragDrop, Button143.DragDrop, Button144.DragDrop
        Dim bt As Button = CType(sender, Button)
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        If bt.Text = Nothing And PositionAdjacente(bt).Text = Nothing Then
            bt.Text = e.Data.GetData(DataFormats.Text)
            bt.BackColor = Color.White
            DataGridView1.Rows.RemoveAt(i)
            dtTemp = DataGridView1.DataSource
            PositionAdjacente(bt).AllowDrop = False
            bt.AllowDrop = False
        End If
    End Sub

    Private Sub TablesDragEnter(sender As Object, e As DragEventArgs) Handles Button145.DragEnter, Button146.DragEnter, Button147.DragEnter, Button148.DragEnter, Button139.DragEnter, Button140.DragEnter, Button141.DragEnter, Button142.DragEnter, Button143.DragEnter, Button144.DragEnter, Button71.DragEnter, Button72.DragEnter, Button73.DragEnter, Button74.DragEnter, Button75.DragEnter, Button70.DragEnter, Button76.DragEnter, Button77.DragEnter, Button78.DragEnter, Button79.DragEnter, Button8006.DragEnter, Button8010.DragEnter, Button8009.DragEnter, Button8008.DragEnter, Button8007.DragEnter, Button8005.DragEnter, Button8001.DragEnter, Button8002.DragEnter, Button8003.DragEnter, Button8004.DragEnter, Button63.DragEnter, Button64.DragEnter, Button65.DragEnter, Button66.DragEnter, Button81.DragEnter, Button82.DragEnter, Button260.DragEnter, Button27.DragEnter, Button62.DragEnter, Button80.DragEnter, Button83.DragEnter, Button30.DragEnter, Button84.DragEnter, Button67.DragEnter, Button68.DragEnter, Button61.DragEnter, Button85.DragEnter, Button86.DragEnter, Button6006.DragEnter, Button6007.DragEnter, Button6002.DragEnter, Button6001.DragEnter, Button6000.DragEnter, Button6003.DragEnter, Button6004.DragEnter, Button6005.DragEnter, Button1997.DragEnter, Button1998.DragEnter, Button1995.DragEnter, Button1996.DragEnter, Button1993.DragEnter, Button1994.DragEnter, Button1999.DragEnter, Button1992.DragEnter, Button1000.DragEnter, Button29.DragEnter, Button28.DragEnter, Button25.DragEnter, Button240.DragEnter, Button300.DragEnter, Button200.DragEnter, Button69.DragEnter, Button109.DragEnter, Button110.DragEnter, Button106.DragEnter, Button105.DragEnter, Button104.DragEnter, Button103.DragEnter, Button101.DragEnter, Button102.DragEnter, Button113.DragEnter, Button114.DragEnter, Button115.DragEnter, Button116.DragEnter, Button117.DragEnter, Button118.DragEnter, Button119.DragEnter, Button120.DragEnter, Button121.DragEnter, Button122.DragEnter, button15.DragEnter, button16.DragEnter, button17.DragEnter, button18.DragEnter, button19.DragEnter, button20.DragEnter, button21.DragEnter, button22.DragEnter, button23.DragEnter, button14.DragEnter, button5.DragEnter, button6.DragEnter, button7.DragEnter, button8.DragEnter, button9.DragEnter, button10.DragEnter, button11.DragEnter, button12.DragEnter, button13.DragEnter, Button4.DragEnter, Button8014.DragEnter, Button8011.DragEnter, Button8013.DragEnter, Button8012.DragEnter, Button87.DragEnter, Button88.DragEnter, Button89.DragEnter, Button90.DragEnter, Button91.DragEnter, Button92.DragEnter, Button93.DragEnter, Button94.DragEnter, Button95.DragEnter, Button96.DragEnter, Button97.DragEnter, Button98.DragEnter, Button99.DragEnter, Button100.DragEnter, Button107.DragEnter, Button108.DragEnter, Button111.DragEnter, Button112.DragEnter, Button123.DragEnter, Button124.DragEnter, Button125.DragEnter, Button126.DragEnter, Button127.DragEnter, Button128.DragEnter, Button129.DragEnter, Button130.DragEnter, Button131.DragEnter, Button132.DragEnter, Button133.DragEnter, Button134.DragEnter, Button135.DragEnter, Button136.DragEnter, Button137.DragEnter, Button138.DragEnter, button52.DragEnter, button53.DragEnter, button54.DragEnter, button55.DragEnter, button56.DragEnter, button57.DragEnter, button58.DragEnter, button59.DragEnter, button60.DragEnter, button51.DragEnter, button42.DragEnter, button43.DragEnter, button44.DragEnter, button45.DragEnter, button46.DragEnter, button47.DragEnter, button48.DragEnter, button49.DragEnter, button50.DragEnter, button41.DragEnter, button32.DragEnter, button33.DragEnter, button34.DragEnter, button35.DragEnter, button36.DragEnter, button37.DragEnter, button38.DragEnter, button39.DragEnter, button40.DragEnter, button31.DragEnter
        If (e.Data.GetDataPresent(DataFormats.Text)) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub Form3_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If (Not (CP1.Enabled = False And CP2.Enabled = False And CS1.Enabled = False And SQ2.Enabled = False And SL2.Enabled = False And ST2.Enabled = False)) Then
            Dim f As System.IO.File
            f.Delete("../../Sauvegarde/" + nameBDD)
        End If
    End Sub
End Class