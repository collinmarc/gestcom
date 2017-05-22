Imports vini_DB
Public Class frmEtatServeur
    Dim oFtpVnc As clsFTPVinicom


    Private Sub frmEtatServeur_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        oFtpVnc = New clsFTPVinicom()

        laUrl.Text = oFtpVnc.ftpConnection.Hostname
        laUtilisateur.Text = oFtpVnc.ftpConnection.Username
        laPassword.Text = oFtpVnc.ftpConnection.Password

        AfficheEtat()
    End Sub
    Private Sub AfficheEtat()
        If oFtpVnc.ftpConnection.isConnected Then
            laEtatServeur.Text = "Connecté"
            cbLockFrom.Enabled = False
            cbLockTo.Enabled = False
            cbUnlock.Enabled = False
            If oFtpVnc.IsLockFrom() Then
                laProtection.Text = "Lock From"
                cbUnlock.Enabled = True
            End If
            If oFtpVnc.IsLockTo() Then
                laProtection.Text = "Lock To"
                cbUnlock.Enabled = True
            End If
            If (Not oFtpVnc.IsLockTo() And Not oFtpVnc.IsLockFrom()) Then
                laProtection.Text = "Libre"
                cbLockTo.Enabled = True
                cbLockFrom.Enabled = True
            End If
        Else
            laEtatServeur.Text = "Déconnecté"
            laProtection.Text = ""
            cbLockFrom.Enabled = False
            cbLockTo.Enabled = False
            cbUnlock.Enabled = False
        End If

    End Sub
    Private Sub cbLockFrom_Click(sender As System.Object, e As System.EventArgs) Handles cbLockFrom.Click
        oFtpVnc.lockFrom()
        AfficheEtat()
    End Sub

    Private Sub cbLockTo_Click(sender As System.Object, e As System.EventArgs) Handles cbLockTo.Click
        oFtpVnc.lockTo()
        AfficheEtat()
    End Sub

    Private Sub cbUnlock_Click(sender As System.Object, e As System.EventArgs) Handles cbUnlock.Click
        If oFtpVnc.IsLockFrom() Then
            oFtpVnc.unlockfrom()
        Else
            oFtpVnc.unlockTo()
        End If
        AfficheEtat()

    End Sub
End Class