Imports System.IO
Public Class BattleForm

    Private Sub UpdateDisplayStats(ByVal intHP As Integer, ByVal intMP As Integer)

        'Call CheckHeroHP from SubFuncModule
        CheckHeroHP()

        If intHeroHP < intHeroLowHP Then

            'If HeroHP is less than HeroLowHP turn font red
            lblHealth.ForeColor = Color.Red

        End If

        'Display Health and Magic
        lblHealth.Text = intHP.ToString() + " / " + intHeroTotalHP.ToString()
        lblMagic.Text = intMP.ToString() + " / " + intHeroTotalMP.ToString()

    End Sub

    Public Sub UpdateEnemyDamageLbls()

        'Display damage done in Enemy damage label
        lblEnemyDamage.Text = intHeroAtk.ToString()

        'Display DamageDesc in EnemyDamageDesc labale
        lblEnemyDamDesc.Text = strDamageDesc

    End Sub

    Private Sub UpdateHeroDamageLbls()

        'Display damage done in Hero damage label
        lblHeroDamage.Text = intEnemyAtk.ToString()

        'Display damage desc in Hero damage desc labal
        lblHeroDamDesc.Text = strDamageDesc

    End Sub

    Private Sub ClearHeroDamageLbls()
        'Sub to clear HeroDamage and HeroDamDesc labels

        lblHeroDamage.Text = ""
        lblHeroDamDesc.Text = ""

    End Sub

    Private Sub ClearEnemyDamangeLbls()
        'Sub to clear EnemyDamage and EnemyDamDesc Labels

        lblEnemyDamage.Text = ""
        lblEnemyDamDesc.Text = ""

    End Sub

    Private Sub MagicTierILbls()
        If blnHeal = True Then

            'Update HeroDamDesc and HeroDamage labels
            lblHeroDamDesc.Text = "HEAL!"
            lblHeroDamage.Text = intMagHeal.ToString()

            'Start EnemyTurnTimer
            tmrEnemyTurnTimer.Start()

        End If

        If blnEnemyCanAtk = True Then

            'Call UpdateEnemyDamageLbls
            UpdateEnemyDamageLbls()

        End If
    End Sub

    Private Sub ButtonsOnOff(ByRef bool As Boolean)

        'Sub to reenable all buttons
        btnAttack.Enabled = bool
        btnDefend.Enabled = bool
        btnMagic.Enabled = bool

        'If the Attack button is enabled, receive focus
        If btnAttack.Enabled = True Then

            btnAttack.Focus()

        End If


    End Sub

    Private Sub LoadMagic(ByRef blnMagic As Boolean, ByRef strMagic As String,
                          ByRef intCost As Integer)

        If blnMagic = True Then

            'Add Magic to Magic combobox with cost
            cboMagic.Items.Add(strMagic + " (" + intCost.ToString + "MP)")

        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Display Hero Name in HeroHealth and HeroMagic Labels
        lblHerHealth.Text = strHeroName & " Health"
        lblHeroMagic.Text = strHeroName & " Magic"

        'Display Level in Level label
        lblLevel.Text = "Level " & intLvlCount + 1

        'Make EnemyIsPoisoned False
        blnEnemyIsPoisoned = False

        Select Case intLvlCount
            Case 1

                'Call LoadMagic sub from MagicModule to test which magic
                'bool is true and assign a magic cost
                LoadMagic(blnThunder, "Thunder", MagicCost(0))

                'Call LoadMagic sub from MagicModule to test which magic
                'bool is true and assign a magic cost
                LoadMagic(blnHeal, "Heal", MagicCost(0))

                'Make Gnome picture visible
                picGnome.Visible = True

            Case 2

                'Call LoadMagic sub from MagicModule to test which magic
                'bool is true and assign a magic cost
                LoadMagic(blnThunder, "Thunder", MagicCost(1))

                'Call LoadMagic sub from MagicModule to test which magic
                'bool is true and assign a magic cost
                LoadMagic(blnHeal, "Heal", MagicCost(1))

                'Call LoadMagic sub from MagicModule to test which magic
                'bool is true and assign a magic cost
                LoadMagic(blnPoison, "Poison", MagicCost(0))

                'Make Dragon picture visible
                picDragon.Visible = True

            Case 3

                'Call LoadMagic sub from MagicModule to test which magic
                'bool is true and assign a magic cost
                LoadMagic(blnThunder, "Thunder", MagicCost(2))

                'Call LoadMagic sub from MagicModule to test which magic
                'bool is true and assign a magic cost
                LoadMagic(blnHeal, "Heal", MagicCost(2))

                'Call LoadMagic sub from MagicModule to test which magic
                'bool is true and assign a magic cost
                LoadMagic(blnPoison, "Poison", MagicCost(1))

                'Make Orc picture visible
                picOrc.Visible = True

            Case 4

                'Call LoadMagic sub from MagicModule to test which magic
                'bool is true and assign a magic cost
                LoadMagic(blnThunder, "Thunder", MagicCost(2))

                'Call LoadMagic sub from MagicModule to test which magic
                'bool is true and assign a magic cost
                LoadMagic(blnHeal, "Heal", MagicCost(2))

                'Call LoadMagic sub from MagicModule to test which magic
                'bool is true and assign a magic cost
                LoadMagic(blnPoison, "Poison", MagicCost(2))

                'Make finalBoss picture visible
                picFinalBoss.Visible = True

        End Select

        If intLvlCount = 0 Then

            'Set/Reset stats
            ResetStats()

            'Call the LoadStats sub from SubFuncModule
            LoadStats()

            'Call DisplayStats sub using Hero health and Hero magic
            UpdateDisplayStats(intHeroHP, intHeroMP)

            'Make Bat picture visible
            picBat.Visible = True

        Else

            'If LvlCount is greater than zero..

            'Make Magic button visible
            btnMagic.Visible = True

            'Call LevelUp sub
            LevelUp()

            'Call UpdateDisplayStats sub
            UpdateDisplayStats(intHeroHP, intHeroMP)

        End If
    End Sub

    Private Sub btnAttack_Click(sender As Object, e As EventArgs) Handles btnAttack.Click

        'Call ButtonsOnOff sub to disable buttons
        ButtonsOnOff(False)

        'Hide Magic panel if open
        pnlMagicSelect.Visible = False

        Select Case intLvlCount
            Case 0

                'Call Attack Sub with Hero damage ranges plugged in
                Attack(intHeroAtk, 10, 10, intEnemyHP, 6, 6)

                'Start Turn Timer tick event
                tmrHeroTurnTimer.Start()

            Case 1

                If blnAtkBonus = True Then

                    'If AtkBonus bool is true, increase attack numbers
                    'by 3
                    Attack(intHeroAtk, 5, 34, intEnemyHP, 8, 17)

                    'Start HeroTurnTimer tick event
                    tmrHeroTurnTimer.Start()

                Else

                    'Call Attack Sub with Hero damage ranges plugged in
                    Attack(intHeroAtk, 5, 31, intEnemyHP, 8, 14)

                    'Start Turn Timer tick event
                    tmrHeroTurnTimer.Start()

                End If

            Case 2

                If blnAtkBonus = True Then

                    'If AtkBonus bool is true, increase attack numbers
                    'by 6
                    Attack(intHeroAtk, 11, 42, intEnemyHP, 8, 30)

                    'Start HeroTurnTimer tick event
                    tmrHeroTurnTimer.Start()

                Else

                    'Call Attack Sub with Hero damage ranges plugged in
                    Attack(intHeroAtk, 11, 36, intEnemyHP, 13, 24)

                    'Start Turn Timer tick event
                    tmrHeroTurnTimer.Start()

                End If


            Case 3

                If blnAtkBonus = True Then

                    'If AtkBonus bool is true, increase attack numbers
                    'by 9
                    Attack(intHeroAtk, 11, 73, intEnemyHP, 11, 65)

                    'Start HeroTurnTimer tick event
                    tmrHeroTurnTimer.Start()

                Else

                    'Call Attack Sub with Hero damage ranges plugged in
                    Attack(intHeroAtk, 11, 64, intEnemyHP, 11, 56)

                    'Start Turn Timer tick event
                    tmrHeroTurnTimer.Start()

                End If

            Case 4

                If blnAtkBonus = True Then

                    'If AtkBonus bool is true, increase attack numbers
                    'by 9
                    Attack(intHeroAtk, 21, 125, intEnemyHP, 16, 110)

                    'Start HeroTimerTick event
                    tmrHeroTurnTimer.Start()

                Else

                    'Call Attack Sub with Hero damage ranges plugged in
                    Attack(intHeroAtk, 21, 116, intEnemyHP, 16, 101)

                    'Start Turn Timer tick event
                    tmrHeroTurnTimer.Start()

                End If


        End Select

        'Call UpdateEnemyDamageLbls sub
        UpdateEnemyDamageLbls()

        'Reset timer
        intTimer = 1

    End Sub

    Public Sub tmrHeroTurnTimer_Tick(sender As Object, e As EventArgs) Handles tmrHeroTurnTimer.Tick

        If intTimer > 0 Then

            'If the timer is greater than zero, subtract one from timer
            intTimer -= 1

        Else

            'If timer equals zero, stop the timer
            tmrHeroTurnTimer.Stop()

            'Reset timer
            intTimer = 1

            'Call ClearEnemyDamageLbls sub / Not needed for Poison Casting
            ClearEnemyDamangeLbls()

            'If the enemy health is above zero
            If intEnemyHP > 0 Then

                Select Case intLvlCount
                    Case 0

                        'Call Attack Sub with Enemy damage ranges assigned
                        Attack(intEnemyAtk, 10, 11, intHeroHP, 6, 6)

                        'Start EnemyTurn Timer
                        tmrEnemyTurnTimer.Start()

                    Case 1

                        'Call Attack Sub with Enemy damage ranges assigned
                        Attack(intEnemyAtk, 6, 26, intHeroHP, 11, 15)

                        'Call CheckHeroLvl sub
                        CheckHeroHP()

                        'Start EnemyTurn Timer
                        tmrEnemyTurnTimer.Start()

                    Case 2

                        'Call CheckPoison from Magic module
                        CheckPoison()

                        'Call Attack Sub with Enemy damage ranges assigned
                        Attack(intEnemyAtk, 6, 41, intHeroHP, 21, 21)

                        'Start EnemyTurn Timer
                        tmrEnemyTurnTimer.Start()

                    Case 3

                        'Call CheckPoison from Magic module
                        CheckPoison()

                        'Call Attack Sub with Enemy damage ranges assigned
                        Attack(intEnemyAtk, 21, 61, intHeroHP, 16, 46)

                        'Start EnemyTurn Timer
                        tmrEnemyTurnTimer.Start()

                    Case 4

                        'Call CheckPoison from Magic module
                        CheckPoison()

                        'Call Attack Sub with Enemy damage ranges assigned
                        Attack(intEnemyAtk, 21, 151, intHeroHP, 51, 101)

                        'Start EnemyTurn Timer
                        tmrEnemyTurnTimer.Start()

                End Select

                'CheckHeroLvlSub
                CheckHeroHP()

                'Call UpdateHeroDamageLbls()
                UpdateHeroDamageLbls()

                'Call IfEnemyIsPoisoned sub
                IfEnemyIsPoisoned()

                'Update Hero Stats
                UpdateDisplayStats(intHeroHP, intHeroMP)

            Else

                Select Case intLvlCount
                    Case 0

                        'Call AddLevel sub from SubFuncModule
                        AddLevel()

                        'Display MessageBox that Hero won
                        MessageBox.Show("HERO WON!")
                        LoadBonusIForm()
                        Me.Close()

                    Case 1

                        'Call AddLevel sub from SubFuncModule
                        AddLevel()

                        'Display MessageBox that Hero won
                        MessageBox.Show("HERO WON!")

                        'Call LoadBonusIIForm from LoadForms module
                        LoadBonusIIForm()

                        'Close this form
                        Me.Close()


                    Case 2

                        'Call AddLevel sub from SubFuncModule
                        AddLevel()

                        'Display MessageBox that Hero won
                        MessageBox.Show("HERO WON!")

                        'Call LoadBattleForm from LoadForms module
                        LoadBattleForm()

                        'Close this form
                        Me.Close()

                    Case 3
                        'Call AddLevel sub from SubFuncModule
                        AddLevel()

                        'Display MessageBox that Hero won
                        MessageBox.Show("HERO WON!")

                        'Call LoadBattleForm from LoadForms moduel
                        LoadBattleForm()

                        'Close this form
                        Me.Close()

                    Case 4

                        'Call AddLevel sub from SubFuncModule
                        AddLevel()

                        'Display MessageBox that Hero won the game
                        MessageBox.Show("YOU ARE A TRUE HERO " & strHeroName & "!")

                        'Close this form
                        Me.Close()

                End Select
            End If
        End If
    End Sub

    Private Sub tmrEnemyTurnTimer_Tick(sender As Object, e As EventArgs) Handles tmrEnemyTurnTimer.Tick

        If intTimer > 0 Then

            'Subtract 1 from Timer
            intTimer -= 1
        Else

            'Stop Enemy timer
            tmrEnemyTurnTimer.Stop()

            If intHeroHP = 0 Then

                'Create instance of LoseForm
                Dim frmLoseForm As New LoseForm

                'Show LoseForm
                frmLoseForm.Show()

                'Close this form
                Close()

            End If

            'Not neccessary in defending, or Attack, Or Lightning / Good for Poison
            ClearEnemyDamangeLbls()

            'Call ClearHeroDamageLbls sub
            ClearHeroDamageLbls()

            'Call ButtonsOnOff sub to reenable all buttons
            ButtonsOnOff(True)

        End If

    End Sub

    Private Sub btnMagic_Click(sender As Object, e As EventArgs) Handles btnMagic.Click

        If pnlMagicSelect.Visible = True Then

            'If the MagicSelect panel is visible, make it not
            pnlMagicSelect.Visible = False

        Else

            'If the MagicSelect panel is not visitble, make it so
            pnlMagicSelect.Visible = True

        End If
    End Sub

    Private Sub btnCastMagic_Click(sender As Object, e As EventArgs) Handles btnCastMagic.Click

        'Upon casting magic, hide panel.
        pnlMagicSelect.Visible = False

        Select Case intLvlCount

            Case 1

                'If level one and first index is selected test MP
                If cboMagic.SelectedIndex = 0 Then

                    'Call TierIMagic sub from Magic Module
                    TierIMagic()

                    'Call MagicTierIlbls sub
                    MagicTierILbls()

                End If

            Case 2
                'If level one and first index is selected test MP
                If cboMagic.SelectedIndex = 0 Then

                    'Call TierIMagic sub from Magic module
                    TierIMagic()

                    'Call MagicTierIbls sub
                    MagicTierILbls()

                ElseIf cboMagic.SelectedIndex = 1 Then

                    'Call TierIIMagic sub from Magic Module
                    TierIIMagic()

                End If

            Case 3
                'If level one and first index is selected call
                'TierIMagic sub from Magic module
                If cboMagic.SelectedIndex = 0 Then

                    'Call TierIMagic sub from Magic Module
                    TierIMagic()

                    'Call MagicTierILbls sub
                    MagicTierILbls()

                ElseIf cboMagic.SelectedIndex = 1 Then

                    'Call TierIIMagic sub from Magic Module
                    TierIIMagic()

                End If

            Case 4
                'If level one and first index is selected call
                'TierIMagic sub from Magic Module
                If cboMagic.SelectedIndex = 0 Then

                    'Call TierIMagic sub from Magic Module
                    TierIMagic()

                    'Call MagicTierILbls sub
                    MagicTierILbls()

                ElseIf cboMagic.SelectedIndex = 1 Then

                    'Call TierIIMagic sub from Magic Module
                    TierIIMagic()

                End If

        End Select


        If blnEnemyCanAtk = True Then

            'If EnemyCanAtk bool is true then start HeroTurnTimer
            'and turn buttons off
            tmrHeroTurnTimer.Start()
            ButtonsOnOff(False)

        End If

        'Reset EnemyCanAtk bool
        blnEnemyCanAtk = False

        'Display Stats
        UpdateDisplayStats(intHeroHP, intHeroMP)

        'Reset timer
        intTimer = 1

    End Sub

    Private Sub btnDefend_Click(sender As Object, e As EventArgs) Handles btnDefend.Click

        Select Case intLvlCount
            Case 0

                'Call Attack Sub with Enemy damage ranges lowered and assigned
                Attack(intEnemyAtk, 10, 10 - 3, intHeroHP, 5, 5 - 2)

                'Start EnemyTurn Timer
                tmrEnemyTurnTimer.Start()

            Case 1

                'Call Attack Sub with Enemy damage ranges lowered and assigned
                Attack(intEnemyAtk, 6, 21 - 5, intHeroHP, 11, 11 - 3)

                'Start EnemyTurn Timer
                tmrEnemyTurnTimer.Start()

            Case 2

                'Call Attack Sub with Enemy damage ranges lowered and assigned
                Attack(intEnemyAtk, 6, 41 - 7, intHeroHP, 21, 21 - 4)

                'Start EnemyTurn Timer
                tmrEnemyTurnTimer.Start()

            Case 3

                'Call Attack Sub with Enemy damage ranges lowered and assigned
                Attack(intEnemyAtk, 21, 61 - 9, intHeroHP, 16, 46 - 5)

                'Start EnemyTurn Timer
                tmrEnemyTurnTimer.Start()

            Case 4

                'Call Attack Sub with Enemy damage ranges lowered and assigned
                Attack(intEnemyAtk, 21, 151 - 11, intHeroHP, 51, 101 - 6)

                'Start EnemyTurn Timer
                tmrEnemyTurnTimer.Start()

        End Select

        'Call ButtonsOnOff sub
        ButtonsOnOff(False)

        'Call CheckHeroHP
        CheckHeroHP()

        'Call UpdateHeroDamageLbls
        UpdateHeroDamageLbls()

        'Update Hero Stats
        UpdateDisplayStats(intHeroHP, intHeroMP)

        'Call CheckPoisonSub
        CheckPoison()

        'Call IfEnemyIsPoisoned
        IfEnemyIsPoisoned()

        'Reset Timer
        intTimer = 1

    End Sub

    Private Sub IfEnemyIsPoisoned()

        If blnEnemyIsPoisoned = True Then

            'Assign POISON to DamageDesc label
            strDamageDesc = "POISON"

            'Call UpdateEnemyDamageLbls
            UpdateEnemyDamageLbls()

            If intPoisonCount = 0 Then

                'If Poison Count is zero, make EnemyIsPoisoned false
                blnEnemyIsPoisoned = False

            End If
        End If

    End Sub

    Private Sub mnuFileNewGame_Click(sender As Object, e As EventArgs) Handles mnuFileNewGame.Click

        'Call the ResetStats sub from SubFunc module
        ResetStats()

        'Call the LoadBattleform sub from LoadForms module
        LoadBattleForm()

        Close()

    End Sub

    Private Sub mnuFileSave_Click(sender As Object, e As EventArgs) Handles mnuFileSave.Click

        'Message warning
        MessageBox.Show("This will only save from begining of level")

        'StreamWriter variable to save file
        Dim saveFile As StreamWriter

        'Create Save file
        saveFile = File.CreateText("herosave.txt")

        Try

            'Save variables to file
            saveFile.WriteLine(strHeroName)
            saveFile.WriteLine(intLvlCount)
            saveFile.WriteLine(intHeroTotalHP)
            saveFile.WriteLine(intHeroTotalMP)
            saveFile.WriteLine(intEnemyTotalHP)
            saveFile.WriteLine(blnAtkBonus)
            saveFile.WriteLine(blnHPBonus)
            saveFile.WriteLine(blnThunder)
            saveFile.WriteLine(blnHeal)
            saveFile.WriteLine(blnPoison)

            'Close the save file
            saveFile.Close()

            'Message confirmation
            MessageBox.Show("Game has been saved")

        Catch ex As Exception

            'Error Message
            MessageBox.Show("You broke the program")

        End Try

    End Sub

    Private Sub mnuFileLoad_Click(sender As Object, e As EventArgs) Handles mnuFileLoad.Click

        If MessageBox.Show("Do you want to load game?", "Title",
                           MessageBoxButtons.YesNo) = DialogResult.Yes Then

            'Declare loadfile variable as StreamReader
            Dim loadFile As StreamReader

            'Open hersave.txt
            loadFile = File.OpenText("herosave.txt")

            'Read lines and assign to variables
            strHeroName = loadFile.ReadLine()
            intLvlCount = loadFile.ReadLine()
            intHeroTotalHP = loadFile.ReadLine()
            intHeroTotalMP = loadFile.ReadLine()
            intEnemyTotalHP = loadFile.ReadLine()
            blnAtkBonus = loadFile.ReadLine()
            blnHPBonus = loadFile.ReadLine()
            blnThunder = loadFile.ReadLine()
            blnHeal = loadFile.ReadLine()
            blnPoison = loadFile.ReadLine()

            'Close the Load file
            loadFile.Close()

            'Load new BattleForm
            LoadBattleForm()

            'Close this form
            Me.Close()

        End If



    End Sub

    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click

        'Close the form
        Me.Close()

    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click

        'Call LoadHelpForm sub from LoadForms sub
        LoadHelpForm()

    End Sub
End Class
