Option Explicit

Dim question, willYouBeMyValentine, noCount
noCount = 0

Msgbox "Your pc has been hacked, LOL", vbExclamation
Msgbox "However, I just have one question for you...", vbExclamation
willYouBeMyValentine = Msgbox("Will you be my valentine?", vbExclamation+vbYesNo)
If willYouBeMyValentine = vbYes Then
    Msgbox "I love you, babe", vbExclamation
Else
    do until noCount = 9 ' Changed to 9 questions
        ' Each question is different; you can customize these as needed
        question = Msgbox("Will you say yes now?", vbExclamation+vbYesNo)
        If question = vbYes Then
            Msgbox "I love you, babe", vbExclamation
            exit do
        End If

        question = Msgbox("How about now?", vbExclamation+vbYesNo)
        If question = vbYes Then
            Msgbox "I love you, babe", vbExclamation
            exit do
        End If

        question = Msgbox("Come on, please?", vbExclamation+vbYesNo)
        If question = vbYes Then
            Msgbox "I love you, babe", vbExclamation
            exit do
        End If

        question = Msgbox("I can be really sweet, you know?", vbExclamation+vbYesNo)
        If question = vbYes Then
            Msgbox "I love you, babe", vbExclamation
            exit do
        End If

        question = Msgbox("What will it take for you to say yes?", vbExclamation+vbYesNo)
        If question = vbYes Then
            Msgbox "I love you, babe", vbExclamation
            exit do
        End If

        question = Msgbox("Isn't it obvious that I adore you?", vbExclamation+vbYesNo)
        If question = vbYes Then
            Msgbox "I love you, babe", vbExclamation
            exit do
        End If

        question = Msgbox("Are you sure you don't love me back?", vbExclamation+vbYesNo)
        If question = vbYes Then
            Msgbox "I love you, babe", vbExclamation
            exit do
        End If

        question = Msgbox("I promise I'm a great valentine!", vbExclamation+vbYesNo)
        If question = vbYes Then
            Msgbox "I love you, babe", vbExclamation
            exit do
        End If

        question = Msgbox("Please, just say yes! 😘", vbExclamation+vbYesNo)
        If question = vbYes Then
            Msgbox "I love you, babe", vbExclamation
            exit do
        End If

        If question = vbNo Then
            noCount = noCount + 1
        End If
    loop
    
    If noCount = 9 Then Msgbox "You're making me work for it, but I love you anyway!", vbExclamation 
End If
