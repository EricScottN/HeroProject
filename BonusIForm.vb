Public Class BonusIForm

    'Variable to count buttons selected
    Dim intButtonSelect As Integer = 0

    Private Sub btnWeapon_Click(sender As Object, e As EventArgs) Handles btnWeapon.Click

        'Change backround color when selected
        btnWeapon.BackColor = Color.LightBlue

        'When weapon button is selected, disable armor button
        btnArmor.Enabled = False

        'Call SimpleAdd to add 1 to ButtonSelect
        SimpleAdd(intButtonSelect, 1)

    End Sub

    Private Sub btnArmor_Click(sender As Object, e As EventArgs) Handles btnArmor.Click

        'Change selected button color
        btnArmor.BackColor = Color.LightBlue

        'When armor button is selected, disable weapon button
        btnWeapon.Enabled = False

        'Call SimpleAdd to add 1 to ButtonSelect
        SimpleAdd(intButtonSelect, 1)

    End Sub

    Private Sub btnLightning_Click(sender As Object, e As EventArgs) Handles btnLightning.Click

        'Change backround color when selected
        btnLightning.BackColor = Color.LightBlue

        'When thunder button is selected, disable heal button
        btnHeal.Enabled = False

        'Call SimpleAdd to add 1 to ButtonSelect
        SimpleAdd(intButtonSelect, 1)

    End Sub

    Private Sub btnHeal_Click(sender As Object, e As EventArgs) Handles btnHeal.Click

        'Change selected button color
        btnHeal.BackColor = Color.LightBlue

        'When heal button is selected, disable thunder button
        btnLightning.Enabled = False

        'Call SimpleAdd to add 1 to ButtonSelect
        SimpleAdd(intButtonSelect, 1)

    End Sub

    Private Sub btnProceed_Click(sender As Object, e As EventArgs) Handles btnProceed.Click

        If intButtonSelect = 2 Then
            'If two buttons have been selected then..

            If btnWeapon.Enabled = True Then

                'If Weapon button is selected, set AtkBonus bool to true
                blnAtkBonus = True

            ElseIf btnArmor.Enabled = True Then

                'If Armor button is selected, set HPBonus bool to true
                blnHPBonus = True

            End If

            If btnLightning.Enabled = True Then

                'If Thunder button is selected, set Thunder bool to true
                blnThunder = True

            ElseIf btnHeal.Enabled = True Then

                'If Heal button is selected, set Heal bool to true
                blnHeal = True

            End If

            'Open BattleForm
            LoadBattleForm()

            'Close this form
            Me.Close()

        Else
            MessageBox.Show("Please select either 'Weapon' or 'Armor', " &
                            "as well as 'Heal' or 'Lightning'")
        End If


    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        'Enable all buttons
        btnWeapon.Enabled = True
        btnArmor.Enabled = True
        btnLightning.Enabled = True
        btnHeal.Enabled = True

        'Restore buttons to default color
        btnWeapon.BackColor = DefaultBackColor
        btnArmor.BackColor = DefaultBackColor
        btnLightning.BackColor = DefaultBackColor
        btnHeal.BackColor = DefaultBackColor

        'Reset ButtonSelect
        intButtonSelect = 0

    End Sub

End Class