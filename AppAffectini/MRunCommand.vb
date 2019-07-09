Module MRunCommand

    Sub RunCommand(ByVal SQLCommand As String, Optional ByVal message As String = "")

        Try
            If AffectINI.cont.State = ConnectionState.Closed Then AffectINI.cont.Open()
            AffectINI.cmde.CommandText = SQLCommand
            AffectINI.cmde.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            If AffectINI.cont.State = ConnectionState.Open Then AffectINI.cont.Close()
        End Try
    End Sub
End Module
