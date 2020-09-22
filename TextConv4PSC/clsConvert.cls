VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer, j As Integer '//my main for loop vars
Public Function StrToBin(ByVal Data As String) As String
'This functin will run through a string and return that string
'in base 2
'(binary)

'If you can't figure our why this function works
'you can read a nice tutorial about binary and hex #'s
'@: http://www.gamehacking.com/tutorials/hex.php

Dim Buffer As String, TmpBuffer As String, ChkVal As Integer
    For i = 1 To Len(Data)
        ChkVal = Asc(Mid(Data, i, 1))
        While ChkVal <> 0
            TmpBuffer = TmpBuffer & (ChkVal Mod 2)
            ChkVal = ChkVal \ 2
        Wend
        While Len(TmpBuffer) < 8
            TmpBuffer = TmpBuffer & "0"
        Wend
        Buffer = Buffer & StrRev(TmpBuffer)
        TmpBuffer = ""
    Next i
    
    StrToBin = Buffer
End Function

Public Function StrRev(ByVal Data As String) As String
Dim Buffer As String
    While Data <> ""
        Buffer = Buffer & Right$(Data, 1)
        Data = Left$(Data, Len(Data) - 1)
    Wend
    
    StrRev = Buffer
End Function

Public Function StrToHex(ByVal Data As String) As String
'This function will convert a string to hex...
Dim Buffer As String, Tmp As Variant

    For i = 1 To Len(Data)
        Tmp = Hex(Asc(Mid(Data, i, 1)))
        If Tmp < 10 Then
            Buffer = Buffer & "0" & Tmp
        Else
            Buffer = Buffer & Tmp
        End If
    Next i
    
    StrToHex = Buffer
End Function

Public Function StrToAsc(ByVal Data As String) As String

Dim Buffer As String, Tmp As Integer
    For i = 1 To Len(Data)
        Tmp = Asc(Mid(Data, i, 1))
        If Tmp < 100 Then
            Tmp = 0 & Tmp
        End If
        Buffer = Buffer & Tmp
    Next i
    
    StrToAsc = Buffer
End Function

Public Function StrToOct(ByVal Data As String) As String

Dim Buffer As String, Tmp As Integer
    For i = 1 To Len(Data)
        Tmp = Oct(Asc((Mid(Data, i, 1))))
        If Tmp < 100 Then
            Tmp = 0 & Tmp
        End If
        Buffer = Buffer & Tmp
    Next i
    
    StrToOct = Buffer
End Function

Public Function HexToStr(ByVal Data As String) As String
Dim Buffer As String

    If Len(Data) Mod 2 <> 0 Then 'this would catch false input
        HexToStr = vbNullString  'such as 776F3 <-- (well, maybe not false input, but it would cause an error)
    Else
        For i = 1 To Len(Data) Step 2
            Buffer = Buffer & Chr("&H" & Mid(Data, i, 2))
        Next i
        HexToStr = Buffer
    End If
End Function

Public Sub Error_()
    MsgBox "Error checking: check for typos and make sure the format type is correct.", vbCritical + vbOKOnly
End Sub

Public Function BinToHex(ByVal Data As String) As String
Dim Buffer As String, TmpBuffer As String, ValOne As String, ValTwo As String
Dim FirstDigit As Integer, SecondDigit As Integer
Dim Final As String

'If you don't understand why this function converts
'a binary value to a asc value then read the info.
'@: http://www.gamehacking.com/tutorials/hex.php

If Len(Data) Mod 8 <> 0 Then
    BinToHex = vbNullString
Else
    For i = 1 To Len(Data) Step 8
        TmpBuffer = Mid(Data, i, 8)
        ValOne = Left(TmpBuffer, 4)
        ValTwo = Right(TmpBuffer, 4)
        For j = 1 To 4
            If Mid(ValOne, j, 1) = 1 Then
                FirstDigit = FirstDigit + (2 ^ (4 - j))
            End If
        Next j
        Select Case FirstDigit
        Case 10: Final = Final & "A"
        Case 11: Final = Final & "B"
        Case 12: Final = Final & "C"
        Case 13: Final = Final & "D"
        Case 14: Final = Final & "E"
        Case 15: Final = Final & "F"
        Case Else: Final = Final & FirstDigit
        End Select
        For j = 1 To 4
            If Mid(ValTwo, j, 1) = 1 Then
                SecondDigit = SecondDigit + (2 ^ (4 - j))
            End If
        Next j
        Select Case SecondDigit
        Case 10: Final = Final & "A"
        Case 11: Final = Final & "B"
        Case 12: Final = Final & "C"
        Case 13: Final = Final & "D"
        Case 14: Final = Final & "E"
        Case 15: Final = Final & "F"
        Case Else: Final = Final & SecondDigit
        End Select
       Buffer = Buffer & Final
       Final = ""
       FirstDigit = 0
       SecondDigit = 0
    Next i
    BinToHex = Buffer
End If
End Function

Public Function AscToStr(ByVal Data As String) As String
Dim Buffer As String
If Len(Data) Mod 3 <> 0 Then
    AscToStr = vbNullString
Else
    For i = 1 To Len(Data) Step 3
        Buffer = Buffer & Chr(Mid(Data, i, 3))
    Next i
    AscToStr = Buffer
End If
End Function

Public Function BinToStr(ByVal Data As String) As String
Dim Buffer As String

If Len(Data) Mod 8 <> 0 Then
    BinToStr = vbNullString
Else
    BinToStr = HexToStr(BinToHex(Data))
End If

End Function
