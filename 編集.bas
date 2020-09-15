Attribute VB_Name = "•ÒW"
Option Explicit

'•ÒWˆ—
Public Sub ValidatData(ByVal Target As Range)
    Dim rep As Object
    Dim i As Long
    
    Application.EnableEvents = False
    Set rep = CreateObject("VBScript.RegExp")
    
    '“ü—Í‹K‘¥‚É”½‚·‚éê‡Ô‚­–ÔŠ|‚¯‚ğs‚È‚Á‚Ä’ˆÓ‚ğ‘£‚·
    For i = 1 To Target.Count: Do
        Target(i).Interior.ColorIndex = xlNone
        If Target(i).text = "" Then Exit Do
        
        'ƒwƒbƒ_‚©‚çˆ—‚ğ‘I‘ğ
        Select Case Cells(1, Target(i).Column)
        Case "Ôíº°ÄŞ"
            '”¼Šp‰pš1Œ…+”¼Šp”š2Œ…
            rep.Pattern = "[A-Za-z][0-9]{2}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "•”•iº°ÄŞ"
            '”¼Šp”š4Œ…
            rep.Pattern = "^[0-9]{4}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "•”•iº°ÄŞ}”Ô"
            '”¼Šp”š4Œ…
            rep.Pattern = "^[0-9]{4}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "C—•û–@º°ÄŞ"
            '”¼Šp‰pš1-3Œ…
            rep.Pattern = "^[A-Za-z]{1,3}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "”N®º°ÄŞ"
            '”¼Šp”š1Œ…
            rep.Pattern = "^[0-9]$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "ÎŞÃŞ¨Œ`óº°ÄŞ"
            '”¼Šp”š2Œ…
            rep.Pattern = "^[0-9]{2}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "¸ŞÚ°ÄŞº°ÄŞ"
            '”¼Šp‰pš1-5Œ…
            rep.Pattern = "^[A-Za-z]{1,5}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "VA"
            '”¼Šp‰pš1-2Œ…
            rep.Pattern = "^[A-Za-z]{1,2}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "ÃŞ°À’Šo”Ô†"
            '”¼Šp‰p”š1-18Œ…
            rep.Pattern = "^[\-0-9A-Za-z]{1,18}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "ŠJn"
            '”¼Šp”š1-12Œ…
            rep.Pattern = "^[0-9]{1,12}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "I—¹"
            '”¼Šp”š1-12Œ…
            rep.Pattern = "^[0-9]{1,12}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "¶‰E"
            '"L" or "R"
            rep.Pattern = "^[LlRr]$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "‘OŒã"
            '"F" or "R"
            rep.Pattern = "^[FfRr]$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "•”•i–¼Ì"
            '”¼Šp‰p”ƒJƒi1-18Œ…
            rep.Pattern = "^[0-9A-Za-z\uFF61-\uFF9F]{1,18}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "•”•i”Ô†"
            '”¼Šp‰p”š1-17Œ…
            rep.Pattern = "^[\-0-9A-Za-z]{1,17}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "\¬¸ŞÙ°Ìß"
            '”¼Šp‰p”š2Œ…
            rep.Pattern = "^[0-9A-Za-z]{2}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "ŠÖ˜Aì‹Æº°ÄŞ"
            '10Œ…
            If LenMbcs(Target(i).text) > 10 Then Call SetInputErr(Target(i)): Exit Do
            
        Case "”õl"
            '20Œ…
            If LenMbcs(Target(i).text) > 20 Then Call SetInputErr(Target(i)): Exit Do
            
        Case "H’À‹æ•ª"
            '1
            If Target(i).text <> "1" Then Call SetInputErr(Target(i)): Exit Do
            
        Case "¸ŞÙ°Ìßº°ÄŞ"
            '* or #
            rep.Pattern = "^[\*#]$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case ""
            'ƒwƒbƒ_‚ª‹ó”’ˆÈ~‚Íƒf[ƒ^‚ª‚È‚¢
            Exit Do
        End Select
        
        '‘å•¶š•ÏŠ·
        Target(i).Value = UCase(Target(i).text)
    Loop Until 1: Next
    
    Set rep = Nothing
    Application.EnableEvents = True
End Sub

Private Sub SetInputErr(ByVal Target As Range, Optional desc As String = "")
    Target.Interior.ColorIndex = 38
    If desc <> "" Then MsgBox desc
End Sub

