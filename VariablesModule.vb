
Module VariablesModule

    'For random number
    Public rand As New Random

    'Global Strings
    Public strHeroName As String                'To store hero name
    Public strDamageDesc As String              'To store damage description

    'Primary integer
    Public intLvlCount As Integer = 0           'Game level

    'Counter integers
    Public intPoisonCount As Integer = 3        'To count turns poisoned
    Public intTimer As Integer = 1              'Variable for timer

    'Base Stat Integers
    Public intHeroTotalHP As Integer = 50       'Hero total health
    Public intHeroTotalMP As Integer = 0        'Hero total magic
    Public intEnemyTotalHP As Integer = 25      'Enemy total health


    'Game time integers subject to change
    Public intHeroHP As Integer                 'Hero current health
    Public intHeroMP As Integer                 'Hero current magic
    Public intEnemyHP As Integer                'Enemy current health
    Public intHeroAtk As Integer                'Hero damage given
    Public intEnemyAtk As Integer               'Enemy damage given
    Public intHeroLowHP As Integer              'Low Hero HP 
    Public intHPBonus As Integer                'HP Bonus

    'Game time bools subject to change
    Public blnHeroLowHP As Boolean = False      'Bool for Hero Low HP
    Public blnEnemyCanAtk As Boolean = False    'Bool for deciding if enemy can attack
    Public blnInsufficientMP As Boolean = False 'Bool for insufficiant MP

    'Bonus TierI Stat Bools
    Public blnAtkBonus As Boolean = False       'Bool for attack bonus
    Public blnHPBonus As Boolean = False        'Bool for HPBonus


    'Bonus Magic Bools and Integers
    Public blnThunder As Boolean = False        'Bool for Thunder
    Public blnHeal As Boolean = False           'Bool for Heal
    Public blnScan As Boolean = False           'Bool for Scan
    Public blnPoison As Boolean = False         'Bool for Poison
    Public blnEnemyIsPoisoned = False           'Bool for deciding if poisoned
    Public intMagDamage As Integer              'Magic bonus damage
    Public intMagHeal As Integer                'Magic heal integer

    'Array to give multipliers
    Public Multipliers = New Decimal() {1.5, 2, 2.5, 3}

    'Array for MagicCosts
    Public MagicCost = New Integer() {5, 10, 15}


End Module
