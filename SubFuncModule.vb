Module SubFuncModule

    Public Sub ResetStats()

        'Reset all base stats and bools
        intLvlCount = 0
        intHeroTotalHP = 50
        intHeroTotalMP = 0
        intEnemyTotalHP = 25
        blnAtkBonus = False
        blnHPBonus = False
        blnThunder = False
        blnHeal = False
        blnPoison = False
        blnEnemyCanAtk = False
        blnEnemyIsPoisoned = False
        blnInsufficientMP = False

    End Sub

    Public Sub LoadStats()

        'Set Hero Health to Total Hero Health
        intHeroHP = intHeroTotalHP

        'Set Hero Magic to Total Hero Magic
        intHeroMP = intHeroTotalMP

        'Set Enemy Health to Total Enemy Health
        intEnemyHP = intEnemyTotalHP

    End Sub

    Public Sub AddLevel()

        'Call SimpleAdd sub to add one to LvlCount
        SimpleAdd(intLvlCount, 1)

    End Sub

    Public Sub SimpleMultiply(ByRef intNumberOne As Integer, ByRef intNumberTwo As Decimal)

        'Number one multiplied by number two equals number one.
        intNumberOne *= intNumberTwo

    End Sub

    Public Sub SimpleSubtract(ByRef intNumberOne As Integer, ByRef intNumberTwo As Integer)

        'Number one minus number two equals number one
        intNumberOne -= intNumberTwo

    End Sub

    Public Sub SimpleAdd(ByRef intNumberOne As Integer, ByRef intNumberTwo As Integer)

        'Number one plus number two equals number one
        intNumberOne += intNumberTwo

    End Sub

    Public Sub HPBonus()
        'Test for HPBonus

        If blnHPBonus = True Then

            'Select Case to test LvlCount
            Select Case intLvlCount
                Case 1

                    'Give 25 bonus HP
                    intHPBonus = 25

                Case 2

                    'Call SimpleMultiply sub
                    SimpleMultiply(intHPBonus, Multipliers(1))

                Case 3

                    'Call SimplyMultiply sub
                    SimpleMultiply(intHPBonus, Multipliers(0))

            End Select

            'Call SimpleAdd sub
            SimpleAdd(intHeroTotalHP, intHPBonus)

        End If
    End Sub

    Public Sub LevelUp()

        'Public Sub to increase all total HP and MP
        'at beginning of level and load them 
        'into current hp and mp variables by calling
        'the LoadStats sub

        Select Case intLvlCount
            Case 1

                'Call SimpleMultiply sub
                SimpleMultiply(intHeroTotalHP, Multipliers(1))

                'Test HPBonus Sub
                HPBonus()

                'Call SimpleMultiply sub
                SimpleMultiply(intEnemyTotalHP, Multipliers(3))

                'Declare HeroTotalMP to integer 10
                intHeroTotalMP = 10

            Case 2

                'Call SimpleSubtract to get base HeroTotalHP
                SimpleSubtract(intHeroTotalHP, intHPBonus)

                'Call SimpleMultiply sub to get HeroTotalHP
                SimpleMultiply(intHeroTotalHP, Multipliers(2))

                'Test HPBonus sub
                HPBonus()

                'Call SimpleMultiply sub to get EnemyTotalHP
                SimpleMultiply(intEnemyTotalHP, Multipliers(3))

                'Call SimplyMultiply sub to get HeroTotalMP
                SimpleMultiply(intHeroTotalMP, Multipliers(2))

            Case 3

                'Call SimpleMultiply sub to get HeroTotalHP
                SimpleMultiply(intHeroTotalHP, Multipliers(2))

                'Test HPBonus sub
                HPBonus()

                'Call SimpleMultiply sub to get EnemyTotalHP
                SimpleMultiply(intEnemyTotalHP, Multipliers(3))

                'Call SimpleMultiply sub to get HeroTotalMP
                SimpleMultiply(intHeroTotalMP, Multipliers(1))

            Case 4

                'Call SimpleMultiply sub to get HeroTotalHP
                SimpleMultiply(intHeroTotalHP, Multipliers(1))

                'Call SimpleMultiply sub to get EnemyTotalHP
                SimpleMultiply(intEnemyTotalHP, Multipliers(1))

                'Call SimpleMultiply sub to get HeroTotalMP
                SimpleMultiply(intHeroTotalMP, Multipliers(1))

        End Select

        'Call LoadStats sub
        LoadStats()

    End Sub

    Public Sub CheckHeroHP()

        'Calculates 15% of HeroTotalHP and stores in HeroLowHP
        intHeroLowHP = (intHeroTotalHP * 0.2)

        'If Hero HP is less than zero, set to zero 
        If intHeroHP < 0 Then
            intHeroHP = 0
        End If

    End Sub

    Public Sub Attack(ByRef intAtk As Integer, ByRef intCritRange As Integer,
                   ByRef intCritLow As Integer, ByRef intHP As Integer,
                   ByRef intAtkRange As Integer, ByRef intAtkLow As Integer)

        'Random number for atk sequence
        Dim intAtkSequence As Integer

        'Generate random number between 1-20 and assign to intAtkSequence
        intAtkSequence = rand.Next(20 + 1)

        'If AtkSequence equals 1
        If intAtkSequence = 1 Then

            'Assign "miss" to DamageDesc variable
            strDamageDesc = "MISS!"

            'Assign zero to Atk variable
            intAtk = 0

        ElseIf intAtkSequence = 19 Or intAtkSequence = 20 Then

            'Assign random number Atk to range of critical numbers
            intAtk = rand.Next(intCritRange) + intCritLow

            'Call SimpleSubtract sub to subtract Atk from HP
            SimpleSubtract(intHP, intAtk)

            'Assign "critical" to DamageDesc variable
            strDamageDesc = "CRITICAL!"

        Else

            'Assign random number Atk to range of normal numbers
            intAtk = rand.Next(intAtkRange) + intAtkLow

            'Call SimpleSubtract sub to subtract Atk from HP
            SimpleSubtract(intHP, intAtk)

            'Assign "hit" to DamageDesc variable
            strDamageDesc = "HIT!"

        End If

    End Sub

End Module
