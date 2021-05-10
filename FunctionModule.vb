Module FunctionModule
    Dim rand As New Random 'reserve random location
    Public g_intBestGame As Integer = 9999
    Public g_intWorstGame As Integer = 0

    Public Function RollDie() As Integer
        'rolldie accepts no arguments
        'generates a random integer from 1 to 6
        'returns a single random integer from 1 to 6

        'Variable to hold and return random integer
        Dim intDie As Integer

        'generate a random die roll
        intDie = rand.Next(6) + 1

        'return random integer
        Return intDie

    End Function

    Public Function FirstRoll() As Integer()
        'firstroll accepts no arguments
        'calls rolldie to append the list dice with the returned value
        'Return an array of 12 values all from 1 to 6

        'initialize variables
        Dim index As Integer
        Const intMAX_SUBSCRIPTS As Integer = 11
        Dim intDice(intMAX_SUBSCRIPTS) As Integer

        'loop throug the array, adding a random die roll at each index
        For index = 0 To 11
            intDice(index) = RollDie()
        Next

        Return intDice

    End Function

    Public Function CountFrequency(ByVal intDice As Integer(), intNumber As Integer) As Integer
        'accepts two arguments, an array of 12 dice and a number 1 to 6
        'finds the frequency the number 1-6 occurs in the array
        'returns the frequency the number occurs in the array

        'initialize variables
        Dim intFrequency As Integer

        'count frequency of the number one
        For Each die In intDice
            If die = intNumber Then
                intFrequency += 1
            End If
        Next

        Return intFrequency
    End Function

    Public Function FindMode(ByVal intDice As Integer()) As Integer
        'accepts the array of 12 dice rolls
        'use countfrequency to determine which die number occurs the most (best-so-far variable)
        'assign that number as mode
        'return so called mode

        'initialize variables
        Dim intMode As Integer 'most frequently occuring die number
        Dim intRecord As Integer = 0 'the highest frequency

        'find most occuring die number in array intDice
        For intCount = 1 To 6
            For Each die In intDice
                If CountFrequency(intDice, intCount) >= intRecord Then
                    'set mode if it is broken
                    intMode = intCount
                    'set record if it is broken
                    intRecord = CountFrequency(intDice, intCount)
                End If
            Next
        Next

        Return intMode
    End Function

    Public Function ListUnmatchedDice(ByVal intDice As Integer()) As Integer()
        'Accepts one argument, an array of 12 dice rolls
        'find the mode
        'loop through the first roll array, creating a separate array of indexs <> mode
        'Return the new array only containing the indexes of the dice to reroll

        'initialize variables
        Dim intMode As Integer
        Dim intIndex As Integer = 0 'set our index for the new array
        Const intMAX_SUBSCRIPTS As Integer = 11
        Dim intFreq As Integer
        Dim intSubscript As Integer

        intMode = FindMode(intDice) 'find and assign the mode

        'Set up new array
        intFreq = CountFrequency(intDice, intMode)
        intSubscript = intMAX_SUBSCRIPTS - intFreq
        Dim intNew(intSubscript) As Integer


        'loop through the array of the dice, adding each index to new array that is <> mode
        For intCount = 0 To (intDice.Length - 1)
            If intDice(intCount) <> intMode Then
                intNew(intIndex) = intCount
                intIndex += 1
            End If
        Next

        'Retunr the resulting list of indexes to be rerolled
        Return intNew

    End Function

    Public Function RerollOne(ByVal intDice As Integer(), intIndex As Integer) As Integer()
        'RerollOne recieves 2 arguments, an array of dice, and indexes
        'calls rolldie to reoll that index
        'returns an array with the new roll

        intDice(intIndex) = RollDie()

        Return intDice
    End Function

    Public Sub RerollMany(ByVal intDice As Integer())
        'Accepts an array of 12 dice
        'Call any previous functions to play the game

        'initialize variables
        Dim intNumberOfRolls As Integer = 1
        Dim intMode As Integer = FindMode(intDice)
        Dim intNewDice(11) As Integer
        Dim intUnmatched() As Integer
        Dim intNumberMatched As Integer = 0

        intUnmatched = ListUnmatchedDice(intDice)

        'Copy the entire array
        For intCount = 0 To (intDice.Length - 1)
            intNewDice(intCount) = intDice(intCount)
        Next

        Do While intNumberMatched < 12
            'Reroll each unmatched die
            For Each index In intUnmatched
                intNewDice = RerollOne(intNewDice, index)
            Next

            'loop to see if all die = mode
            For Each die In intNewDice
                intMode = FindMode(intNewDice)
                If die = intMode Then
                    intNumberMatched += 1
                Else
                    intNumberMatched = 0
                End If
            Next

            'Re-index the unmatched array
            ReDim intUnmatched(ListUnmatchedDice(intNewDice).Length - 1)

            'get a new list of unmatched dice, increment the roll counter, output the array of dice
            intUnmatched = ListUnmatchedDice(intNewDice)
            intNumberOfRolls += 1
            OutputDice(intNewDice)
        Loop

        'Find the best and worst games
        If intNumberOfRolls <= g_intBestGame Then
            g_intBestGame = intNumberOfRolls
        End If

        If intNumberOfRolls >= g_intWorstGame Then
            g_intWorstGame = intNumberOfRolls
        End If

        'output best/worst rolls
        MainForm.lblBestGame.Text = g_intBestGame.ToString
        MainForm.lblWorstGame.Text = g_intWorstGame.ToString
        MainForm.lblNumberOfRolls.Text = intNumberOfRolls.ToString
    End Sub

    Public Sub ResetAll()
        'Accepts no arguments
        'resets all controls to starting values
        MainForm.lblBestGame.Text = String.Empty
        MainForm.lblWorstGame.Text = String.Empty
        g_intBestGame = 9999
        g_intWorstGame = 0
        MainForm.lblConsole.Text = String.Empty
        MainForm.lstGameRolls.Items.Clear
        MainForm.lblNumberOfRolls.Text = String.Empty
        MainForm.txtNumberOfGames.Text = "1"
        MainForm.txtNumberOfGames.Focus()

    End Sub


    Public Sub StartGame()
        'Accepts no arguments
        'begin calling functions to play the games
        'play as many games as the user inputs

        Dim intGames As Integer = 1 'default of 1 game
        Dim intTotalGamesPlayed As Integer = 0 'counter for total games
        Dim intDice As Integer() = FirstRoll() 'first array of dice
        Dim intNumberOfRolls As Integer = 0 'Number of rolls per game
        Dim intCount As Integer = 0

        'validate the user hasnt entered a letter
        If Integer.TryParse(MainForm.txtNumberOfGames.Text, intGames) Then
            MainForm.lstGameRolls.Items.Add("Die Number: " & vbCrLf)
            MainForm.lstGameRolls.Items.Add("0      1      2      3      4      5      6      7      8      9      10      11")

            OutputDice(intDice)

            'loop to play the specified number games
            While intCount < intGames

                RerollMany(intDice)

                intCount += 1

                MainForm.lstGameRolls.Items.Add(vbCrLf)
                If intCount < intGames Then
                    MainForm.lstGameRolls.Items.Add("Die Number: " & vbCrLf)
                    MainForm.lstGameRolls.Items.Add("0      1      2      3      4      5      6      7      8      9      10      11")
                End If

            End While
        Else
            MessageBox.Show("Must enter a valid number of games.")
        End If






        'MainForm.lblConsole.Text = String.Empty
        'MainForm.lblConsole.Text = "The game has started. "

        'For Each element In intDice
        'MainForm.lblConsole.Text &= element
        'MainForm.lblConsole.Text &= " "
        'Next

        'MainForm.lblConsole.Text &= "  New index 0: "
        'For Each element In RerollOne(intDice, 0)
        'MainForm.lblConsole.Text &= element & " "
        'Next

    End Sub

    Public Sub OutputDice(ByVal intDice As Integer())
        'accepts an array of dice
        'loops to output the array

        Dim strDice As String = String.Empty


        For Each element As Integer In intDice
            strDice &= element.ToString
            strDice &= "------"
        Next

        MainForm.lstGameRolls.Items.Add(strDice)
    End Sub

End Module
