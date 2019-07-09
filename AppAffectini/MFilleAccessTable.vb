Imports System.Data.OleDb

Module MFilleAccessTable

    Public Sub FillAccessTable()

        Dim str As String
        Dim i As Integer
        Dim cpt As Integer = 1

        str = "Insert into CrystalReport([N], [Matricule], [Nom], [Prenom], [Section], [Gr], [Local], [Position], [Module], [DateExam], [HDebut], [HFin]) values (?,?,?,?,?,?,?,?,?,?,?,?)"
        AffectINI.connAccess.Open()

        For i = 0 To AffectINI.dtt.Rows.Count - 1

            Dim cmd As OleDbCommand = New OleDbCommand(str, AffectINI.connAccess)

            cmd.Parameters.Add(New OleDbParameter("N", CType(cpt, Integer)))
            If (AffectINI.dtt(i)(0).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Matricule", CType(AffectINI.dtt(i)(0), String)))
            End If
            If (AffectINI.dtt(i)(1).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Nom", CType(AffectINI.dtt(i)(1), String)))
            End If
            If (AffectINI.dtt(i)(2).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Prenom", CType(AffectINI.dtt(i)(2), String)))
            End If
            If (AffectINI.dtt(i)(3).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Section", CType(AffectINI.dtt(i)(3), String)))
            End If
            If (AffectINI.dtt(i)(4).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Gr", CType(AffectINI.dtt(i)(4), String)))
            End If
            If (AffectINI.dtt(i)(5).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Local", CType(AffectINI.dtt(i)(5), String)))
            End If
            If (IsNumeric(AffectINI.dtt(i)(6))) Then
                cmd.Parameters.Add(New OleDbParameter("Position", CType(AffectINI.dtt(i)(6), Integer)))
            End If
            If (AffectINI.dtt(i)(7).length > 0) Then
                cmd.Parameters.Add(New OleDbParameter("Module", CType(AffectINI.dtt(i)(7), String)))
            End If

            cmd.Parameters.Add(New OleDbParameter("DatExam", CType(AffectINI.dtt(i)(8), Date)))
            cmd.Parameters.Add(New OleDbParameter("HDebut", CType(AffectINI.dtt(i)(9), Date)))
            cmd.Parameters.Add(New OleDbParameter("HFin", CType(AffectINI.dtt(i)(10), Date)))

            cpt = cpt + 1

            Try
                cmd.ExecuteNonQuery()
                cmd.Dispose()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Next

        AffectINI.connAccess.Close()
    End Sub
End Module
