Imports System.Data.OleDb

Module MRequette

    Dim daa As OleDbDataAdapter


    Public Sub Requette(ByRef Requette As String)

        AffectINI.connAccess.Open()
        daa = New OleDbDataAdapter(Requette, AffectINI.connAccess)
        daa.Fill(AffectINI.dtt)
        AffectINI.connAccess.Close()
    End Sub
End Module
