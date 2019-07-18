
Module MagicModule

    Public Sub MagicAttack(ByRef intMagBonusRange As Integer, ByRef intMagBonusLow As Integer,
                       ByRef intAtkRange As Integer, ByRef intAttackLow As Integer)

        'Generate random number for MagAtkBonus to range of numbers
        intMagDamage = rand.Next(intMagBonusRange) + (intMagBonusLow)

        'Generate random number for HeroAtk to range of numbers
        intHeroAtk = rand.Next(intAtkRange) + (intAttackLow)

        'Subtract the result of Hero plus MagBonus from Enemy HP
        intEnemyHP -= (intHeroAtk + intMagDamage)
    End Sub

    Public Sub MagicHeal(ByRef intMagHealRange As Integer, IntMagHealLow As Integer)

        'Generate random number starting at MagHealLow with a range of MagHealRange
        'and assign to MagHeal variable
        intMagHeal = rand.Next(intMagHealRange) + IntMagHealLow

        If (intMagHeal + intHeroHP) > intHeroTotalHP Then

            'Set MagHeal to the difference of HeroTotalHP and HeroHP
            intMagHeal = intHeroTotalHP - intHeroHP

            'If MagHeal is greater than HeroTotalHP, set to HeroTotalHP
            intHeroHP = intHeroTotalHP

        Else

            'Add MagHeal to HeroHP
            intHeroHP += intMagHeal

        End If
    End Sub

    Public Sub TierIMagic()

        If blnThunder = True Then

            'Call MagicAttack sub from SubFunc module
            Thunder()

            If blnEnemyCanAtk = True Then

                'If EnemyCanAtk is true then change DamageDesc string
                'to Lightning
                strDamageDesc = "LIGHTNING"

            End If


        End If

        'If Heal bool is true then test hero MP
        If blnHeal = True Then

            'Call MagicHeal sub from SubFuncModule
            Heal()

        End If
    End Sub

    Public Sub Thunder()
        'Select case for LvlCount
        Select Case intLvlCount
            Case 1

                If intHeroMP >= 5 Then

                    'Subtract 5MP from HeroMP
                    intHeroMP -= 5

                    'Call MagicAttack sub 
                    MagicAttack(3, 6, 8, 14)

                    'Set EnemyCanAtk bool to true
                    blnEnemyCanAtk = True

                Else

                    'Call InsufficientMP sub
                    InsufficientMP()

                End If

            Case 2

                If intHeroMP >= 10 Then

                    'Subtract 10MP from HeroMP
                    intHeroMP -= 10

                    'Call MagicAttack sub with multipliers used for parameters
                    MagicAttack(3 * Multipliers(1), 6 * Multipliers(1),
                                8 * Multipliers(1), 14 * Multipliers(1))

                    'Set EnemyCanAtk boolean to true
                    blnEnemyCanAtk = True

                Else

                    'Call InsufficientMP sub
                    InsufficientMP()

                End If

            Case 3

                If intHeroMP >= 15 Then

                    'Subtract 15MP from HeroMP
                    intHeroMP -= 15

                    'Call MagicAttack sub with multipliers used for parameters
                    MagicAttack(3 * Multipliers(3), 6 * Multipliers(3),
                                8 * Multipliers(3), 14 * Multipliers(3))

                    'Set EnemyCanAtk boolean to true
                    blnEnemyCanAtk = True

                Else

                    'Call InsufficientMP sub
                    InsufficientMP()

                End If

            Case 4

                If intHeroMP >= 15 Then

                    'Subtract 15MP from HeroMP
                    intHeroMP -= 15

                    'Call MagicAttack sub with multipliers used for parameters
                    MagicAttack(3 * Multipliers(3), 6 * Multipliers(3),
                                8 * Multipliers(3), 14 * Multipliers(3))

                    'Set EnemyCanAtkboolean to true
                    blnEnemyCanAtk = True

                Else

                    'Call InsufficientMP sub
                    InsufficientMP()

                End If

        End Select
    End Sub

    Public Sub Heal()
        Select Case intLvlCount
            'Select Case for LvlCount
            Case 1

                If intHeroMP >= 5 Then

                    'Subtract 5MP from HeroMP
                    intHeroMP -= 5

                    'Call MagicHeal sub
                    MagicHeal(6, 26)

                    'Assign "Heal" to DamageDesc Label
                    strDamageDesc = "Heal"

                Else

                    'Call InsufficientMP sub
                    InsufficientMP()

                End If

            Case 2

                If intHeroMP >= 10 Then

                    'Subtract 10MP from HeroMP
                    intHeroMP -= 10

                    'Call MagicHeal sub with multipliers factored into 
                    'parameters
                    MagicHeal(6 * Multipliers(1), 26 * Multipliers(1))

                    'Assign "Heal" to DamageDesc Label
                    strDamageDesc = "Heal"

                Else

                    'Call InsufficientMP sub
                    InsufficientMP()

                End If

            Case 3

                If intHeroMP >= 15 Then

                    'Subtract 10MP from HeroMP
                    intHeroMP -= 15

                    'Cal' MagicHeal sub with multipliers factored into
                    'parameteres
                    MagicHeal(6 * Multipliers(3), 26 * Multipliers(3))

                    'Assign "Heal" to DamageDesc Label
                    strDamageDesc = "Heal"

                Else

                    'Call Insufficient MP sub
                    InsufficientMP()

                End If

            Case 4

                If intHeroMP >= 15 Then

                    'Subtract 10MP from HeroMP
                    intHeroMP -= 15

                    'Cal' MagicHeal sub with multipliers factored into
                    'parameteres
                    MagicHeal(6 * Multipliers(3), 26 * Multipliers(3))

                    'Assign "Heal" to DamageDesc Label
                    strDamageDesc = "Heal"

                Else

                    'Call Insufficient MP sub
                    InsufficientMP()

                End If
        End Select
    End Sub

    Public Sub TierIIMagic()

        'If Poison bool is true..
        If blnPoison = True Then

            'Check if Hero has sufficient MP
            If intHeroMP >= 5 Then

                'Subtract 5 from HeroMP
                intHeroMP -= 5

                'Set PoisonCount to three
                intPoisonCount = 3

                'Set EnemyIsPoisoned to true
                blnEnemyIsPoisoned = True

                'Set EnemyCanAtk to true
                blnEnemyCanAtk = True

            Else

                'Call InsufficientMP sub
                InsufficientMP()

            End If
        End If
    End Sub

    Public Sub CheckPoison()
        'Sub that checks enemy for poison

        If blnEnemyIsPoisoned = True Then

            'Call Poison Sub
            Poison()

        End If
    End Sub

    Public Sub Poison()
        'Sub that poisons the enemy for three turns

        'Select case to verify level
        Select Case intLvlCount
            Case 2

                If intPoisonCount > 0 Then

                    'Generate random number between 5-10
                    'and assign to HeroAtk
                    intHeroAtk = rand.Next(6) + 6

                    'Subtract HeroAtk from EnemyHP
                    intEnemyHP -= intHeroAtk

                    'Subtract one from PoisonCount
                    intPoisonCount -= 1

                End If

            Case 3

                If intPoisonCount > 0 Then

                    'Generate random number between 5-10
                    'and assign to HeroAtk
                    intHeroAtk = rand.Next(11) + 11

                    'Subtract HeroAtk from EnemyHP
                    intEnemyHP -= intHeroAtk

                    'Subtract one from PoisonCount
                    intPoisonCount -= 1

                End If

            Case 4
                If intPoisonCount > 0 Then

                    'Generate random number between 5-10
                    'and assign to HeroAtk
                    intHeroAtk = rand.Next(21) + 21

                    'Subtract HeroAtk from EnemyHP
                    intEnemyHP -= intHeroAtk

                    'Subtract one from PoisonCount
                    intPoisonCount -= 1

                End If

        End Select
    End Sub

    Public Sub InsufficientMP()

        'Show MessageBox error
        MessageBox.Show("Not Enough Magic Points!")

        'Set EnemyCanAtk booleon to False
        blnEnemyCanAtk = False

    End Sub

End Module
