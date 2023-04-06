Option Explicit

' Set up variables
Dim words, word, guess, attempts, remaining_attempts, coins, hint_cost, reward, sound_folder
words = Array("password", "banana", "apple", "car", "house")
coins = 0
hint_cost = 10
reward = 5
sound_folder = "C:\sounds\" ' Replace with the path to your sound files

' Create a function to play a sound effect
Sub PlaySound(sound_file)
  Dim sound_path
  sound_path = sound_folder & sound_file
  Dim player
  Set player = CreateObject("WMPlayer.OCX")
  player.URL = sound_path
  player.settings.volume = 100
  player.controls.play
  Do While player.playState <> 1
    WScript.Sleep 100
  Loop
  Set player = Nothing
End Sub

Do
  ' Choose a random word
  word = words(Int(Rnd * UBound(words) + 1))
  
  ' Reset the attempts and remaining attempts
  attempts = 5
  remaining_attempts = attempts
  
  ' Show the game section and hide the start section
  WScript.Echo "Guess a Word Game"
  WScript.Echo "--------------------"
  WScript.Echo "You have " & attempts & " attempts to guess the word."
  
  ' Prompt for hint
  Dim hint
  hint = MsgBox("Would you like a hint? Cost: " & hint_cost & " coins.", vbQuestion + vbYesNo, "Hint")
  If hint = vbYes And coins >= hint_cost Then
    coins = coins - hint_cost
    Dim hint_word
    hint_word = Mid(word, Int(Rnd * Len(word)) + 1, 1)
    WScript.Echo "The word contains the letter """ & hint_word & """."
  ElseIf hint = vbYes And coins < hint_cost Then
    WScript.Echo "You don't have enough coins for a hint."
  End If
  
  Do While remaining_attempts > 0
    ' Get the user's guess
    guess = LCase(InputBox("Guess the word (" & remaining_attempts & " attempts remaining):"))
    
    ' Decrement the remaining attempts
    remaining_attempts = remaining_attempts - 1
    
    ' Check if the guess is correct
    If guess = word Then
      WScript.Echo "Congratulations! You guessed the word and earned " & reward & " coins."
      coins = coins + reward
      PlaySound "correct.wav"
      Exit Do
    Else
      ' Check if the user has any attempts left
      If remaining_attempts > 0 Then
        WScript.Echo "Incorrect. You have " & remaining_attempts & " attempts remaining."
      Else
        WScript.Echo "Game over. The word was """ & word & """."
        PlaySound "incorrect.wav"
        Exit Do
      End If
    End If
  Loop
  
  ' Show the user's current coins
  WScript.Echo "You now have " & coins & " coins."
  
  ' Prompt the user to play again
  Dim play_again
  play_again = MsgBox("Would you like to play again?", vbQuestion + vbYesNo, "Play Again")
Loop While play_again = vbYes
