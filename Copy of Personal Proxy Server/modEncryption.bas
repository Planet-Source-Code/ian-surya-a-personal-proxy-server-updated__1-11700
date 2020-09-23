Attribute VB_Name = "modEncryption"
Option Explicit

'i got this base64 code from a friend so if anybody claim this code and want me to replace this code just send me a mail
'i don't know if this code is copyrighted
Public Function Base64Encode(ByVal BinaryData As String) As String
    Dim retString As String
    Dim Byte1 As Byte
    Dim Byte2 As Byte
    Dim Byte3 As Byte
    Dim i As Integer
    Dim X As Integer

    Do While Len(BinaryData) > 0
        Byte1 = Asc(BinaryData)
        BinaryData = Mid(BinaryData, 2)
        X = 2
        If Len(BinaryData) >= 1 Then
            Byte2 = Asc(BinaryData)
            BinaryData = Mid(BinaryData, 2)
            X = 1
        End If
        If Len(BinaryData) >= 1 Then
            Byte3 = Asc(BinaryData)
            BinaryData = Mid(BinaryData, 2)
            X = 0
        End If
        
        retString = retString & Base64Char(Int(Byte1 / 4))
        retString = retString & Base64Char(((Byte1 And 3) * 16) + Int(Byte2 / 16))
        retString = retString & Base64Char(((Byte2 And 15) * 4) + Int(Byte3 / 64))
        retString = retString & Base64Char(Byte3 And 63)
    Loop
    If X = 1 Then
        retString = Left(retString, Len(retString) - 1) & "="
    ElseIf X = 2 Then
        retString = Left(retString, Len(retString) - 2) & "=="
    End If
    Base64Encode = retString
End Function

Public Function Base64Decode(AsciiData As String) As String
    Dim counter As Integer
    Dim Temp As String
    Dim DecodeTable As Variant
    Dim Out(2) As Byte
    Dim inp(3) As Byte
    DecodeTable = Array("255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "62", "255", "255", "255", "63", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", "255", "255", "255", "64", "255", "255", "255", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", _
    "18", "19", "20", "21", "22", "23", "24", "25", "255", "255", "255", "255", "255", "255", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255" _
    , "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255")

    For counter = 1 To Len(AsciiData) Step 4
        inp(0) = DecodeTable(Asc(Mid$(AsciiData, counter, 1)))
        inp(1) = DecodeTable(Asc(Mid$(AsciiData, counter + 1, 1)))
        inp(2) = DecodeTable(Asc(Mid$(AsciiData, counter + 2, 1)))
        inp(3) = DecodeTable(Asc(Mid$(AsciiData, counter + 3, 1)))
        Out(0) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
        Out(1) = ((inp(1) And &HF) * 16) Or ((inp(2) \ 4) And &HF)
        Out(2) = ((inp(2) And &H3) * 64) Or inp(3)

        If inp(2) = 64 Then
            Out(0) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
            Temp = Temp & Chr(Out(0) And &HFF)
        ElseIf inp(3) = 64 Then
            Out(0) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
            Out(1) = ((inp(1) And &HF) * 16) Or ((inp(2) \ 4) And &HF)
            Temp = Temp & Chr(Out(0) And &HFF) & Chr(Out(1) And &HFF)
        Else
            Temp = Temp & Chr(Out(0) And &HFF) & Chr(Out(1) And &HFF) & Chr(Out(2) And &HFF)
        End If
    Next
    Base64Decode = Temp
End Function

Private Function Base64Char(ByVal bit6Number As Byte) As String
    Select Case bit6Number
        Case 0 To 25
            Base64Char = Chr(65 + bit6Number)
        Case 26 To 51
            Base64Char = Chr(97 + (bit6Number - 26))
        Case 52 To 61
            Base64Char = Chr(48 + (bit6Number - 52))
        Case 62
            Base64Char = "+"
        Case 63
            Base64Char = "/"
        Case Else
            MsgBox "Error bit6Number > 63", vbOKOnly + vbCritical, "Base64Char"
    End Select
End Function

Public Function DoXOR(vData As String, Key As String) As String
    Dim lngCtr As Long, strtemp As String, strdata As String
    Dim ByteArray(3) As Byte, fl As Long, lngKey As Long
    
    lngKey = GenerateKey(Key)
    
    ByteArray(0) = lngKey And 255
    ByteArray(1) = (lngKey \ 2 ^ 8) And 255
    ByteArray(2) = (lngKey \ 2 ^ 16) And 255
    ByteArray(3) = (lngKey \ 2 ^ 24) And 255
    strdata = vData
    fl = Len(strdata)
    For lngCtr = 1 To fl
        strtemp = Chr(Asc(Mid(strdata, lngCtr, 1)) Xor ByteArray((lngCtr - 1) Mod UBound(ByteArray)))
        Mid(strdata, lngCtr, 1) = strtemp
    Next lngCtr
    DoXOR = strdata
End Function

Private Function GenerateKey(Key As String) As Long
Dim lngCtr As Long, lngTempKey As Long
    lngTempKey = 0
    For lngCtr = 1 To Len(Key)
        lngTempKey = lngTempKey + Asc(Mid(Key, lngCtr, 1))
    Next lngCtr
    GenerateKey = lngTempKey
End Function

