Module MMinFenetre

    Sub MinFenetre()

        If AffectINI.WindowState = FormWindowState.Maximized Then
            AffectINI.WindowState = FormWindowState.Normal
        Else
            AffectINI.WindowState = FormWindowState.Minimized
        End If
    End Sub
End Module
