Public Class LoseForm

    Private Sub btnYes_Click(sender As Object, e As EventArgs) Handles btnYes.Click
        'Set Level Count back to 1
        intLvlCount = 0

        'Call LoadBattleForm sub from LoadForms module
        LoadBattleForm()

        'Close this form
        Close()

    End Sub

    Private Sub btnNo_Click(sender As Object, e As EventArgs) Handles btnNo.Click

        'Close this form
        Close()

    End Sub

End Class