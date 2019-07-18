Module LoadFormsModule

    Public Sub LoadBattleForm()

        'Declare BattleForm as variable
        Dim frmBattleForm As New BattleForm

        'Show BattleForm
        frmBattleForm.Show()

    End Sub

    Public Sub LoadBonusIForm()

        'Declare BonusIForm as variable
        Dim frmBonusIForm As New BonusIForm

        'Show BonusIForm
        frmBonusIForm.Show()

    End Sub

    Public Sub LoadBonusIIForm()

        'Declare BonusIIForm as vairable
        Dim frmBonusIIForm As New BonusIIForm

        'Show BonusIIForm
        frmBonusIIForm.Show()

    End Sub

    Public Sub LoadHelpForm()

        'Declare HelpForm as variable
        Dim frmHelpForm As New HelpForm

        'Show HelpForm
        frmHelpForm.ShowDialog()

    End Sub

End Module
