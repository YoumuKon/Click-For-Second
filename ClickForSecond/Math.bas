Attribute VB_Name = "Math"
Option Explicit
Public IT%, Ts#, I2%, TI%, I%
Public Function BuyCheck(Value&, money#) As Boolean
    BuyCheck = False
    If money >= Value Then
        money = money - Value
        BuyCheck = True
    End If
End Function

Public Function bitHex(ByVal str$) As String
Dim bit As String
    Do While Not Len(str) = 0
        bit = Left(str, 4)
        Select Case bit
            Case "0000": bitHex = bitHex & "0"
            Case "0001": bitHex = bitHex & "1"
            Case "0010": bitHex = bitHex & "2"
            Case "0011": bitHex = bitHex & "3"
            Case "0100": bitHex = bitHex & "4"
            Case "0101": bitHex = bitHex & "5"
            Case "0110": bitHex = bitHex & "6"
            Case "0111": bitHex = bitHex & "7"
            Case "1000": bitHex = bitHex & "8"
            Case "1001": bitHex = bitHex & "9"
            Case "1010": bitHex = bitHex & "A"
            Case "1011": bitHex = bitHex & "B"
            Case "1100": bitHex = bitHex & "C"
            Case "1101": bitHex = bitHex & "D"
            Case "1110": bitHex = bitHex & "E"
            Case "1111": bitHex = bitHex & "F"
        End Select
        str = Mid(str, 4)
    Loop
    bitHex = "&H" & bitHex
End Function

Public Function hexBit(ByVal str$) As String
Dim hex As String
    str = Mid(str, 2)
    Do While Not Len(str) = 0
        hex = Left(str, 4)
        Select Case hex
            Case "0": hexBit = hexBit & "0000"
            Case "1": hexBit = hexBit & "0001"
            Case "2": hexBit = hexBit & "0010"
            Case "3": hexBit = hexBit & "0011"
            Case "4": hexBit = hexBit & "0100"
            Case "5": hexBit = hexBit & "0101"
            Case "6": hexBit = hexBit & "0110"
            Case "7": hexBit = hexBit & "0111"
            Case "8": hexBit = hexBit & "1000"
            Case "9": hexBit = hexBit & "1001"
            Case "A": hexBit = hexBit & "1010"
            Case "B": hexBit = hexBit & "1011"
            Case "C": hexBit = hexBit & "1100"
            Case "D": hexBit = hexBit & "1101"
            Case "E": hexBit = hexBit & "1110"
            Case "F": hexBit = hexBit & "1111"
        End Select
        str = Mid(str, 4)
    Loop
    hexBit = hexBit
End Function

Public Function ResSave() As String
Dim Boo(NumTopR) As Integer, stuff As Integer
    ResSave = ""
    stuff = ""
    For I = 0 To NumTopR
        Boo(I) = -NumTotalR(I)
        ResSave = ResSave & str(Boo(I))
    Next I
    For I = 0 To NumTopR
        Boo(I) = -ResTI(0, I)
        stuff = ResSave & str(Boo(I))
    Next I
    ResSave = bitHex(ResSave) & "+" & bitHex(stuff)
End Function

Public Sub bitBoo(str$, arrb() As Boolean)
    For I = 0 To NumTopR
        If Mid(str, I, 1) = "1" Then
            arrb(I) = True
            Else: arrb(I) = False
        End If
    Next I
End Sub
