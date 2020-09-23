Attribute VB_Name = "modExplodForm"
' by        :   Rudy Alex Kohn
' e-mail    :   rudyalexkohn@hotmail.com
' Please use if you like, but credit me for it

Option Explicit

Sub ExplodeForm(Frm As Form, Optional Maximize As Boolean = True)
' This sub is just to add some 'spice' to you'r program
' Now.. go play with the values in the for..next loops to get other effects
    
    Dim i           As Integer
    Dim iWidth      As Integer
    Dim iHeight     As Integer

    With Frm

    ' Store original sizes
    iWidth = .Width
    iHeight = .Height

    If Maximize Then
        .Width = 0      ' Sets width and height to 0
        .Height = 0
        .Show           ' Displays form
        For i = 0 To 5000 Step 500
            .Width = i
            .Height = i
            .Left = (Screen.Width - .Width) / 2
            .Top = (Screen.Height - .Height) / 2
        Next
        ' Reset form to original size
        .Width = iWidth
        .Height = iHeight
    Else
        For i = 5000 To 0 Step -250
            .Width = i
            .Height = i
            .Left = (Screen.Width - .Width) / 2
            .Top = (Screen.Height - .Height) / 2
        Next
        ' Unloads when done
        Unload Frm
    End If

    End With
End Sub
