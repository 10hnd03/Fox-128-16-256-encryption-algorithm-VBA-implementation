Attribute VB_Name = "Funkcije"
Option Explicit
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Pohrana As LongPtr, ByVal Izvor As LongPtr, ByVal BrojElemenata As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' Key Vector
Public Function KeyVector() As Long()
Dim p(31) As Long
p(0) = 0: p(1) = 17: p(2) = 34: p(3) = 51: p(4) = 68: p(5) = 85: p(6) = 102: p(7) = 119
p(8) = 136: p(9) = 153: p(10) = 170: p(11) = 187: p(12) = 204: p(13) = 221: p(14) = 238: p(15) = 255
p(16) = 255: p(17) = 238: p(18) = 221: p(19) = 204: p(20) = 187: p(21) = 170: p(22) = 153: p(23) = 136
p(24) = 119: p(25) = 102: p(26) = 85: p(27) = 68: p(28) = 51: p(29) = 34: p(30) = 17: p(31) = 0
KeyVector = p
End Function
' Text Vector
Public Function TextVector() As Long()
Dim p(15) As Long
p(0) = 1: p(1) = 35: p(2) = 69: p(3) = 103: p(4) = 137: p(5) = 171: p(6) = 205: p(7) = 239
p(8) = 254: p(9) = 220: p(10) = 186: p(11) = 152: p(12) = 118: p(13) = 84: p(14) = 50: p(15) = 16
TextVector = p
End Function

Public Function Test32K()
Dim Text As String, Kljuc As String, obj As Object, Izlaz() As Long, PKljuc() As Long, TekstPolje() As Long, i As Long
Dim startKaoCurrency As Currency
Dim krajKaoCurrency As Currency
Dim frekvencija As Currency
Dim protekloSekundi As Double
Dim protekloNanosekundi As Double


Dim putanjaDatoteke As String
    Dim fFile As Integer
    Dim mapaZaSpremanje As String
    
 If CurrentProject.Path = "" Then
        mapaZaSpremanje = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    Else
        mapaZaSpremanje = CurrentProject.Path
    End If
    
    putanjaDatoteke = mapaZaSpremanje & "\Test.txt"
     

Text = GenerirajNasumicno(32000): Kljuc = "Test"
Set obj = New ObradaTeksta: TekstPolje = obj.ObradaStringa(Text)
Set obj = New Fox: PKljuc = obj.pPKljuc(Kljuc)

For i = 0 To 0
QueryPerformanceFrequency frekvencija
QueryPerformanceCounter startKaoCurrency

Izlaz = obj.Sifriraj(TekstPolje, PKljuc, 0)

QueryPerformanceCounter krajKaoCurrency
protekloSekundi = CDbl(krajKaoCurrency - startKaoCurrency) / CDbl(frekvencija)
Debug.Print "" & protekloSekundi
fFile = FreeFile
    Open putanjaDatoteke For Append As #fFile
    Print #fFile, protekloSekundi
    Close #fFile
Next i
End Function

Public Function GenerirajNasumicno(Broj As Long) As String
    Dim putanjaDatoteke As String
    Dim skupZnakova As String
    Dim konacniTekst As String
    Dim brojZnakova As Long
    Dim i As Long
    Dim duljinaSkupa As Long
    Dim datotekaBroj As Integer
    
    brojZnakova = Broj
    putanjaDatoteke = CurrentProject.Path & "\test.txt"
  
    ' Definiramo skup dopuštenih èitljivih znakova (ASCII 32 - 126)
    skupZnakova = "ABCÈÆDÐŽŠIJKLMNOPQRSTUVZWabcèædðefghijklmnopžš€123456789<>xy×÷}{\|€³£ß¤½´ÿ`²°¢^¡"
   
    duljinaSkupa = Len(skupZnakova)
    
    ' Pred-alokacija memorije za string (ubrzava izvoðenje na Long razini)
    konacniTekst = String$(brojZnakova, " ")
    
    ' Inicijaliziramo generator nasumiènih brojeva
    Randomize
    
    ' Punimo string nasumiènim odabirom iz definiranog skupa
    For i = 1 To brojZnakova
        Mid$(konacniTekst, i, 1) = Mid$(skupZnakova, Int(duljinaSkupa * Rnd) + 1, 1)
    Next i
    
    'Otvaramo datoteku i pišemo generirani tekst
    datotekaBroj = FreeFile
    Open putanjaDatoteke For Output As #datotekaBroj
    Print #datotekaBroj, konacniTekst; ' Toèka-zarez sprjeèava dodavanje skrivenog vbCrLf na kraj
    Close #datotekaBroj
    
    GenerirajNasumicno = konacniTekst
End Function

Public Function LongHex(ByVal bLong As Long) As String
Dim a As Long, b As Long, s As String, maska1 As Long, maska2 As Long, Tabela(15) As Long, i As Long

Tabela(0) = 48: Tabela(1) = 49: Tabela(2) = 50: Tabela(3) = 51: Tabela(4) = 52: Tabela(5) = 53: Tabela(6) = 54: Tabela(7) = 55
Tabela(8) = 56: Tabela(9) = 57: Tabela(10) = 65: Tabela(11) = 66: Tabela(12) = 67: Tabela(13) = 68: Tabela(14) = 69: Tabela(15) = 70


a = 0: b = 0


a = a + (((bLong And 128) <> 0) And 8)
a = a + (((bLong And 64) <> 0) And 4)
a = a + (((bLong And 32) <> 0) And 2)
a = a + (((bLong And 16) <> 0) And 1)

b = bLong And 15


maska1 = 0: maska2 = 0
For i = 0 To 15
maska1 = maska1 Or ((a = i) And Tabela(i))
maska2 = maska2 Or ((b = i) And Tabela(i))
Next i


LongHex = Chr(maska1) & Chr(maska2)

End Function

Public Function HexLong(Str As String) As Long()
Dim DuljinaStringa As Long, Tabela(15) As Long, i As Long, j As Long, Maska As Long
DuljinaStringa = Len(Str) - 1
ReDim Polje(DuljinaStringa) As Integer


CopyMemory VarPtr(Polje(0)), StrPtr(Str), DuljinaStringa + DuljinaStringa + 2

Tabela(0) = 48: Tabela(1) = 49: Tabela(2) = 50: Tabela(3) = 51: Tabela(4) = 52: Tabela(5) = 53: Tabela(6) = 54: Tabela(7) = 55
Tabela(8) = 56: Tabela(9) = 57: Tabela(10) = 65: Tabela(11) = 66: Tabela(12) = 67: Tabela(13) = 68: Tabela(14) = 69: Tabela(15) = 70


Dim r As Integer

For i = 0 To UBound(Polje)

r = 0
For j = 0 To 15
r = r Or (Not ((Polje(i) - Tabela(j)) <> 0) And j)
Next j
Polje(i) = r

Next i

Tabela(0) = 0: Tabela(1) = 16: Tabela(2) = 32: Tabela(3) = 48: Tabela(4) = 64: Tabela(5) = 80: Tabela(6) = 96: Tabela(7) = 112
Tabela(8) = 128: Tabela(9) = 144: Tabela(10) = 160: Tabela(11) = 176: Tabela(12) = 192: Tabela(13) = 208: Tabela(14) = 224: Tabela(15) = 240

For i = 0 To UBound(Polje) Step 2

r = 0

For j = 0 To 15
r = r Or (Not ((Polje(i) - j) <> 0) And Tabela(j))
Next j
Polje(i) = r


r = 0

For j = 0 To 15
r = r Or (Not ((Polje(i + 1) - j) <> 0) And j)
Next j
Polje(i + 1) = r


Next i

Maska = 0
Maska = Maska Or (((DuljinaStringa And 128) <> 0) And &H40)
Maska = Maska Or (((DuljinaStringa And 64) <> 0) And &H20)
Maska = Maska Or (((DuljinaStringa And 32) <> 0) And &H10)
Maska = Maska Or (((DuljinaStringa And 16) <> 0) And &H8)
Maska = Maska Or (((DuljinaStringa And 8) <> 0) And &H4)
Maska = Maska Or (((DuljinaStringa And 4) <> 0) And &H2)
Maska = Maska Or (((DuljinaStringa And 2) <> 0) And &H1)
    
    
ReDim Izlaz(Maska) As Long



j = 0

For i = 0 To UBound(Polje) Step 2
Izlaz(j) = Polje(i) + Polje(i + 1)

j = j + 1
Next i

HexLong = Izlaz

End Function
