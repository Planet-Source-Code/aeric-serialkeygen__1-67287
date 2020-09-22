Attribute VB_Name = "modKeyGen"
' This Serial Number Verification is used to 'protect' a program name GoldFish
' Programmed by Aeric Poon
Option Explicit

' Function to generate a random 6-digit Challenge Code
Public Function GenKey() As String
    Dim i As Integer
    Randomize
    For i = 1 To 6
        GenKey = GenKey & (Rnd * 100 Mod 26) Mod 10
    Next
End Function

' For example: Input is 123456
' The generated results is  GF16 J8E8 Z0U5 P0K0 F0A5
' First 2 characters is GF follow by 2 numbers
' Third character is taken from first input digit (i.e 1)
' Forth character is taken from sixth input digit (i.e 6)
' CInt(Mid(ch, 4, 3)) get last 3 digit (i.e 456)
' You can modify to take remaining digit (i.e. 23456)
' by replacing CInt(Mid(ch, 4, 3)) with CInt(Mid(ch, 2, 4))
' Then number is Modulus 26 to get a number between 0 to 25 (A-Z is 26 letters)
' Plus 65 to get A to Z, Example: A = 0 + 65, Z = 25 + 65
' Decipher the rest of code yourself
' Ask me if you still not understand
' I created this algorithm long time ago and
' I wanted to generate a serial number which consist of alpha and numeric alternately
' The complete string is 4 + (4 x 4) = 20 characters

Public Function GenSerial(ch As String) As String
    Dim m As Integer
    Dim n As Integer
    Dim s As String
    Dim c As String
    Dim k As Integer
    
    'A-Z = 65-90 , 0-9 = 48-57
    GenSerial = "GF" & Mid(ch, 1, 1) & Mid(ch, 6, 1)
    c = Chr((CInt(Mid(ch, 4, 3)) Mod 26) + 65)
    k = Asc(c)
    For m = 1 To 4
        For n = 1 To 2
            c = Chr((ch + k) Mod 26 + 65)
            k = k + Mid(ch, m + 2, 1)
            s = s + c
            GenSerial = GenSerial & s
            s = ""
            k = Asc(c)
            s = s & (k * Mid(ch, m + 1, 1)) Mod 10
            GenSerial = GenSerial & s
            s = ""
        Next
    Next
End Function
