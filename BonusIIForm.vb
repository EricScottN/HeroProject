Public Class BonusIIForm

    Private Sub btnProceed_Click(sender As Object, e As EventArgs) Handles btnProceed.Click

        'Make Poison bool true
        blnPoison = True

        'Call LoadBattleForm from LoadForms module
        LoadBattleForm()

        'Close this form
        Close()

    End Sub

End Class