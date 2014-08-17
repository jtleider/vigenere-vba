Option Explicit

Function VIGENERE_ENCRYPT(pt As String, key As String)
'Encrypt plaintext pt using given key using Vigenere cipher.
'Assumes the key consists of a-z characters (either upper or lower case), no special characters
'including spaces. Plaintext may include spaces but otherwise only letters.

Dim ct As String, ptChar As String, keyChar As String, i As Integer

pt = Replace(pt, " ", "") ' Strip spaces from plaintext first

For i = 0 To Len(pt) - 1
    ptChar = UCase(Mid(pt, i + 1, 1))
    keyChar = UCase(Mid(key, (i Mod Len(key)) + 1, 1))
    ct = ct & Chr(Asc("A") + ((Asc(ptChar) + Asc(keyChar) - 2 * Asc("A")) Mod 26))
Next i

VIGENERE_ENCRYPT = ct

End Function

Sub VigenereEncryptVisually()
  'Do not run if you have a sheet named "Vigenere Encryption Process" with important data!
  'Assumes there is a sheet titled "Encrypt", where cells B1-B2 contain plaintext and key, respectively.
    Dim pt As String, key As String, ptChar As String, keyChar As String, i As Integer
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Vigenere Encryption Process").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Sheets.Add(After:=Sheets("Encrypt")).Name = "Vigenere Encryption Process"
    ActiveWindow.Zoom = 75
    
    Sheets("Encrypt").Activate
    Range("B1").Activate
    pt = ActiveCell.Value
    key = ActiveCell.Offset(1, 0).Value
    
    Sheets("Vigenere Encryption Process").Activate
    Range("A1").Value = "Plaintext"
    Range("A2").Value = "Key"
    Range("A3").Value = "Ciphertext"
    Columns("A:A").EntireColumn.AutoFit

    pt = UCase(Replace(pt, " ", ""))
    key = UCase(Replace(key, " ", ""))
    
    Range("A1").Activate
    For i = 1 To Len(pt)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = Mid(pt, i, 1)
    Next i
    Range("A2").Activate
    For i = 0 To Len(pt) - 1
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = Mid(key, (i Mod Len(key)) + 1, 1)
    Next i
    
    Range("A6").Activate
    GenerateVigenere
    
    For i = 1 To Len(pt)
        keyChar = Mid(key, ((i - 1) Mod Len(key)) + 1, 1)
        ptChar = Mid(pt, i, 1)
        Range("B7:AA7").Offset(Asc(keyChar) - Asc("A"), 0).Interior.ColorIndex = 6
        Range("B7:B32").Offset(0, Asc(ptChar) - Asc("A")).Interior.ColorIndex = 6
        Range("B7").Offset(Asc(keyChar) - Asc("A"), Asc(ptChar) - Asc("A")).Interior.ColorIndex = 3
        Range("A3").Offset(0, i).Value = Chr(Asc("A") + ((Asc(ptChar) + Asc(keyChar) - 2 * Asc("A")) Mod 26))
        Application.Wait (Now() + TimeValue("00:00:05"))
        Range("B7:AA7").Offset(Asc(keyChar) - Asc("A"), 0).ClearFormats
        Range("B7:B32").Offset(0, Asc(ptChar) - Asc("A")).ClearFormats
    Next i
    
End Sub

Private Sub GenerateVigenere()
'Generates a Vigenere square, used in the Vigenere cipher.
'Places square around current active cell.
'This is used by VigenereEncryptVisually() above.

Dim square(25, 25) As String

Dim row As Long, col As Long, i As Long, j As Long

For i = 25 To 0 Step -1
    For j = 0 To 25
        square(i, j) = Chr(Asc("A") + ((i + j) Mod 26))
    Next j
Next i

For i = 0 To 25
    For j = 25 To 0 Step -1
        ActiveCell.Offset(i + 1, j + 1).Value = square(i, j)
    Next j
Next i

End Sub
