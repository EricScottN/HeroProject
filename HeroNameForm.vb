Public Class HeroNameForm

    Private Sub HeroNameForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Apply focus to Name textbox
        txtName.Focus()

    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click

        'Assign user input to HeroName string
        strHeroName = txtName.Text

        'Call LoadBattleForm sub from LoadForms module
        LoadBattleForm()

        'Close this form
        Close()

    End Sub

End Class