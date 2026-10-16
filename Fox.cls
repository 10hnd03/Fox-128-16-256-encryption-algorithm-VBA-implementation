VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Pohrana As LongPtr, ByVal Izvor As LongPtr, ByVal BrojElemenata As Long)
Private Declare PtrSafe Sub SecureZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByVal Pohrana As LongPtr, ByVal BrojElemenata As LongPtr)

' GF2 Multiplication
Private Function MulGF2_8(ByVal a As Long, ByVal b As Long) As Long
Dim i As Long, r As Long, e As Long

r = 0: e = 1

For i = 0 To 7
r = r Xor (a And (b And e) <> 0)
a = a + a
e = e + e
Next i

e = ((r And &H4000) <> 0): r = r Xor (e And &H3E40)
e = ((r And &H2000) <> 0): r = r Xor (e And &H1F20)
e = ((r And &H1000) <> 0): r = r Xor (e And &HF90)
e = ((r And &H800) <> 0): r = r Xor (e And &H7C8)
e = ((r And &H400) <> 0): r = r Xor (e And &H3E4)
e = ((r And &H200) <> 0): r = r Xor (e And &H1F2)
e = ((r And &H100) <> 0): r = r Xor (e And &HF9)

r = r And 255

MulGF2_8 = r
End Function

' PKey Function
Private Function PKljuc(Kljuc As String) As Long() 'rijesiti brisanje ulaza i integer polja
Dim Spremnik(31) As Integer, i As Long, Maska As Long, DuljinaUlaza As Long, Adresa As LongPtr
Dim PrazniString As LongPtr, LongSpremnik(31) As Long, Rezultat As Long

Dim Konstanta(31) As Integer
Konstanta(0) = 183:  Konstanta(16) = 98
Konstanta(1) = 225:  Konstanta(17) = 231
Konstanta(2) = 81:   Konstanta(18) = 22
Konstanta(3) = 98:   Konstanta(19) = 15
Konstanta(4) = 138:  Konstanta(20) = 56
Konstanta(5) = 237:  Konstanta(21) = 180
Konstanta(6) = 42:   Konstanta(22) = 218
Konstanta(7) = 106:  Konstanta(23) = 86
Konstanta(8) = 191:  Konstanta(24) = 167
Konstanta(9) = 113:  Konstanta(25) = 132
Konstanta(10) = 88:  Konstanta(26) = 217
Konstanta(11) = 128: Konstanta(27) = 4
Konstanta(12) = 156: Konstanta(28) = 81
Konstanta(13) = 244: Konstanta(29) = 144
Konstanta(14) = 243: Konstanta(30) = 207
Konstanta(15) = 199: Konstanta(31) = 239

Maska = ((Len(Kljuc) - 32) And &H80000000) <> 0
DuljinaUlaza = (LenB(Kljuc) And Maska) Or (64 And Not Maska)
PrazniString = Len(Kljuc) <> 0

Adresa = (StrPtr(Kljuc) And PrazniString) Or (VarPtr(Konstanta(0)) And (Not PrazniString))
CopyMemory VarPtr(Spremnik(0)), Adresa, DuljinaUlaza
DuljinaUlaza = Len(Kljuc)
SecureZeroMemory StrPtr(Kljuc), LenB(Kljuc)

Maska = 0

For i = 0 To UBound(Spremnik)
Maska = Maska Xor CLng(i = DuljinaUlaza)
Spremnik(i) = Spremnik(i) Or (Konstanta((i - DuljinaUlaza) And Maska)) And Maska
LongSpremnik(i) = Spremnik(i)
Next i

SecureZeroMemory VarPtr(Spremnik(0)), 32

For i = 0 To 31
Rezultat = 0
Rezultat = Rezultat Or 128 And Not ((LongSpremnik(i) Xor 8364) <> 0)
Rezultat = Rezultat Or 130 And Not ((LongSpremnik(i) Xor 8218) <> 0)
Rezultat = Rezultat Or 132 And Not ((LongSpremnik(i) Xor 8222) <> 0)
Rezultat = Rezultat Or 133 And Not ((LongSpremnik(i) Xor 8230) <> 0)
Rezultat = Rezultat Or 134 And Not ((LongSpremnik(i) Xor 8224) <> 0)
Rezultat = Rezultat Or 135 And Not ((LongSpremnik(i) Xor 8225) <> 0)
Rezultat = Rezultat Or 137 And Not ((LongSpremnik(i) Xor 8240) <> 0)
Rezultat = Rezultat Or 138 And Not ((LongSpremnik(i) Xor 352) <> 0)
Rezultat = Rezultat Or 139 And Not ((LongSpremnik(i) Xor 8249) <> 0)
Rezultat = Rezultat Or 140 And Not ((LongSpremnik(i) Xor 346) <> 0)
Rezultat = Rezultat Or 141 And Not ((LongSpremnik(i) Xor 356) <> 0)
Rezultat = Rezultat Or 142 And Not ((LongSpremnik(i) Xor 381) <> 0)
Rezultat = Rezultat Or 143 And Not ((LongSpremnik(i) Xor 377) <> 0)
Rezultat = Rezultat Or 145 And Not ((LongSpremnik(i) Xor 8216) <> 0)
Rezultat = Rezultat Or 146 And Not ((LongSpremnik(i) Xor 8217) <> 0)
Rezultat = Rezultat Or 147 And Not ((LongSpremnik(i) Xor 8220) <> 0)
Rezultat = Rezultat Or 148 And Not ((LongSpremnik(i) Xor 8221) <> 0)
Rezultat = Rezultat Or 149 And Not ((LongSpremnik(i) Xor 8226) <> 0)
Rezultat = Rezultat Or 150 And Not ((LongSpremnik(i) Xor 8211) <> 0)
Rezultat = Rezultat Or 151 And Not ((LongSpremnik(i) Xor 8212) <> 0)
Rezultat = Rezultat Or 153 And Not ((LongSpremnik(i) Xor 8482) <> 0)
Rezultat = Rezultat Or 154 And Not ((LongSpremnik(i) Xor 353) <> 0)
Rezultat = Rezultat Or 155 And Not ((LongSpremnik(i) Xor 8250) <> 0)
Rezultat = Rezultat Or 156 And Not ((LongSpremnik(i) Xor 347) <> 0)
Rezultat = Rezultat Or 157 And Not ((LongSpremnik(i) Xor 357) <> 0)
Rezultat = Rezultat Or 158 And Not ((LongSpremnik(i) Xor 382) <> 0)
Rezultat = Rezultat Or 159 And Not ((LongSpremnik(i) Xor 378) <> 0)
Rezultat = Rezultat Or 161 And Not ((LongSpremnik(i) Xor 711) <> 0)
Rezultat = Rezultat Or 162 And Not ((LongSpremnik(i) Xor 728) <> 0)
Rezultat = Rezultat Or 163 And Not ((LongSpremnik(i) Xor 321) <> 0)
Rezultat = Rezultat Or 165 And Not ((LongSpremnik(i) Xor 260) <> 0)
Rezultat = Rezultat Or 170 And Not ((LongSpremnik(i) Xor 350) <> 0)
Rezultat = Rezultat Or 175 And Not ((LongSpremnik(i) Xor 379) <> 0)
Rezultat = Rezultat Or 178 And Not ((LongSpremnik(i) Xor 731) <> 0)
Rezultat = Rezultat Or 179 And Not ((LongSpremnik(i) Xor 322) <> 0)
Rezultat = Rezultat Or 185 And Not ((LongSpremnik(i) Xor 261) <> 0)
Rezultat = Rezultat Or 186 And Not ((LongSpremnik(i) Xor 351) <> 0)
Rezultat = Rezultat Or 188 And Not ((LongSpremnik(i) Xor 317) <> 0)
Rezultat = Rezultat Or 189 And Not ((LongSpremnik(i) Xor 733) <> 0)
Rezultat = Rezultat Or 190 And Not ((LongSpremnik(i) Xor 318) <> 0)
Rezultat = Rezultat Or 191 And Not ((LongSpremnik(i) Xor 380) <> 0)
Rezultat = Rezultat Or 192 And Not ((LongSpremnik(i) Xor 340) <> 0)
Rezultat = Rezultat Or 195 And Not ((LongSpremnik(i) Xor 258) <> 0)
Rezultat = Rezultat Or 197 And Not ((LongSpremnik(i) Xor 313) <> 0)
Rezultat = Rezultat Or 198 And Not ((LongSpremnik(i) Xor 262) <> 0)
Rezultat = Rezultat Or 200 And Not ((LongSpremnik(i) Xor 268) <> 0)
Rezultat = Rezultat Or 202 And Not ((LongSpremnik(i) Xor 280) <> 0)
Rezultat = Rezultat Or 204 And Not ((LongSpremnik(i) Xor 282) <> 0)
Rezultat = Rezultat Or 207 And Not ((LongSpremnik(i) Xor 270) <> 0)
Rezultat = Rezultat Or 208 And Not ((LongSpremnik(i) Xor 272) <> 0)
Rezultat = Rezultat Or 209 And Not ((LongSpremnik(i) Xor 323) <> 0)
Rezultat = Rezultat Or 210 And Not ((LongSpremnik(i) Xor 327) <> 0)
Rezultat = Rezultat Or 213 And Not ((LongSpremnik(i) Xor 336) <> 0)
Rezultat = Rezultat Or 216 And Not ((LongSpremnik(i) Xor 344) <> 0)
Rezultat = Rezultat Or 217 And Not ((LongSpremnik(i) Xor 366) <> 0)
Rezultat = Rezultat Or 219 And Not ((LongSpremnik(i) Xor 368) <> 0)
Rezultat = Rezultat Or 222 And Not ((LongSpremnik(i) Xor 354) <> 0)
Rezultat = Rezultat Or 224 And Not ((LongSpremnik(i) Xor 341) <> 0)
Rezultat = Rezultat Or 227 And Not ((LongSpremnik(i) Xor 259) <> 0)
Rezultat = Rezultat Or 229 And Not ((LongSpremnik(i) Xor 314) <> 0)
Rezultat = Rezultat Or 230 And Not ((LongSpremnik(i) Xor 263) <> 0)
Rezultat = Rezultat Or 232 And Not ((LongSpremnik(i) Xor 269) <> 0)
Rezultat = Rezultat Or 234 And Not ((LongSpremnik(i) Xor 281) <> 0)
Rezultat = Rezultat Or 236 And Not ((LongSpremnik(i) Xor 283) <> 0)
Rezultat = Rezultat Or 239 And Not ((LongSpremnik(i) Xor 271) <> 0)
Rezultat = Rezultat Or 240 And Not ((LongSpremnik(i) Xor 273) <> 0)
Rezultat = Rezultat Or 241 And Not ((LongSpremnik(i) Xor 324) <> 0)
Rezultat = Rezultat Or 242 And Not ((LongSpremnik(i) Xor 328) <> 0)
Rezultat = Rezultat Or 245 And Not ((LongSpremnik(i) Xor 337) <> 0)
Rezultat = Rezultat Or 248 And Not ((LongSpremnik(i) Xor 345) <> 0)
Rezultat = Rezultat Or 249 And Not ((LongSpremnik(i) Xor 367) <> 0)
Rezultat = Rezultat Or 251 And Not ((LongSpremnik(i) Xor 369) <> 0)
Rezultat = Rezultat Or 254 And Not ((LongSpremnik(i) Xor 355) <> 0)
Rezultat = Rezultat Or 255 And Not ((LongSpremnik(i) Xor 729) <> 0)
Rezultat = Rezultat Or 131 And Not ((LongSpremnik(i) Xor 402) <> 0)
Rezultat = Rezultat Or 136 And Not ((LongSpremnik(i) Xor 710) <> 0)
Rezultat = Rezultat Or 140 And Not ((LongSpremnik(i) Xor 338) <> 0)
Rezultat = Rezultat Or 152 And Not ((LongSpremnik(i) Xor 732) <> 0)
Rezultat = Rezultat Or 156 And Not ((LongSpremnik(i) Xor 339) <> 0)
Rezultat = Rezultat Or 159 And Not ((LongSpremnik(i) Xor 376) <> 0)

LongSpremnik(i) = Rezultat Or Not (Rezultat <> 0) And LongSpremnik(i)
Next i

PKljuc = LongSpremnik
End Function

Public Function pPKljuc(Ulaz As String) As Long()
pPKljuc = PKljuc(Ulaz)
End Function

' MKey Function
Private Function MKljuc(PKljuc() As Long, Maska As Long) As Long()
Dim i As Long, KonstantaA As Long, KonstantaB As Long, Izracun As Long, Izlaz(31) As Long

KonstantaA = 118: KonstantaB = 106
Izracun = (KonstantaA + KonstantaB) And 255
Izlaz(0) = PKljuc(0) Xor (Izracun And Not Maska)
Izracun = KonstantaA: Izracun = (Izracun + Izracun + KonstantaB) And 255
Izlaz(1) = PKljuc(1) Xor (Izracun And Not Maska)
KonstantaA = (KonstantaA + KonstantaB) And 255
KonstantaB = Izracun

For i = 2 To 31
Izracun = KonstantaA
Izracun = (Izracun + KonstantaB) And 255
Izlaz(i) = PKljuc(i) Xor (Izracun And Not Maska)
KonstantaA = KonstantaB
KonstantaB = Izracun
Next i

MKljuc = Izlaz
End Function

' DKey Function
Private Function DKljuc(KljucM() As Long, ByVal BrojRunde As Long) As Long()
Dim LFSR_Runde(32) As Long, i As Long, Maska As Long, Prijenos As Long, a As Long, b As Long, c As Long, MKljuc(31) As Long

a = LFSRVrijednosti(BrojRunde + BrojRunde + BrojRunde): b = LFSRVrijednosti(BrojRunde + BrojRunde + BrojRunde + 1): c = LFSRVrijednosti(BrojRunde + BrojRunde + BrojRunde + 2)

For i = 0 To 32 Step 3
Maska = (a And 128) <> 0: Prijenos = (c And 128) <> 0: c = ((c + c) And 255) Xor (Maska And 27)
Maska = (b And 128) <> 0: b = ((b + b) And 255) Xor -Prijenos
a = ((a + a) And 255) Xor -Maska
LFSR_Runde(i) = a: LFSR_Runde(i + 1) = b: LFSR_Runde(i + 2) = c
a = LFSR_Runde(i): b = LFSR_Runde(i + 1): c = LFSR_Runde(i + 2)
Next i

For i = 0 To 31
MKljuc(i) = LFSR_Runde(i) Xor KljucM(i)
Next i

DKljuc = MKljuc
End Function

' LFSR Index Array
Private Function LFSRVrijednosti(ByVal Vrijednost As Long) As Long
Dim LFSR(47) As Long
LFSR(0) = 181: LFSR(1) = 8: LFSR(2) = 122: LFSR(3) = 67: LFSR(4) = 169: LFSR(5) = 56
LFSR(6) = 73: LFSR(7) = 247: LFSR(8) = 31: LFSR(9) = 184: LFSR(10) = 200: LFSR(11) = 89
LFSR(12) = 66: LFSR(13) = 180: LFSR(14) = 26: LFSR(15) = 160: LFSR(16) = 231: LFSR(17) = 199
LFSR(18) = 62: LFSR(19) = 79: LFSR(20) = 65: LFSR(21) = 122: LFSR(22) = 26: LFSR(23) = 166
LFSR(24) = 213: LFSR(25) = 23: LFSR(26) = 240: LFSR(27) = 191: LFSR(28) = 212: LFSR(29) = 56
LFSR(30) = 161: LFSR(31) = 190: LFSR(32) = 18: LFSR(33) = 240: LFSR(34) = 231: LFSR(35) = 175
LFSR(36) = 61: LFSR(37) = 52: LFSR(38) = 193: LFSR(39) = 166: LFSR(40) = 27: LFSR(41) = 227
LFSR(42) = 223: LFSR(43) = 109: LFSR(44) = 208: LFSR(45) = 110: LFSR(46) = 211: LFSR(47) = 101
LFSRVrijednosti = LFSR(Vrijednost)
End Function

' Sigma8 processing
Private Function ObradaSBoxUlaza(Ulaz() As Long) As Long()
Dim i As Long

For i = 0 To UBound(Ulaz) Step 8
Ulaz(i) = RacunajSboxVrijednost(Ulaz(i))
Ulaz(i + 1) = RacunajSboxVrijednost(Ulaz(i + 1))
Ulaz(i + 2) = RacunajSboxVrijednost(Ulaz(i + 2))
Ulaz(i + 3) = RacunajSboxVrijednost(Ulaz(i + 3))
Ulaz(i + 4) = RacunajSboxVrijednost(Ulaz(i + 4))
Ulaz(i + 5) = RacunajSboxVrijednost(Ulaz(i + 5))
Ulaz(i + 6) = RacunajSboxVrijednost(Ulaz(i + 6))
Ulaz(i + 7) = RacunajSboxVrijednost(Ulaz(i + 7))
Next i

ObradaSBoxUlaza = Ulaz
End Function

' Sigma8 On the Fly Calculating Function
Private Function RacunajSboxVrijednost(ByVal bajt As Long) As Long
Dim prvi As Long, drugi As Long, Razlika As Long, Maska As Long, i As Long, Rezultat As Long, OR4(3) As Long

prvi = bajt And 240
prvi = ((prvi And 128) <> 0 And 8) Or _
       ((prvi And 64) <> 0 And 4) Or _
       ((prvi And 32) <> 0 And 2) Or _
       ((prvi And 16) <> 0 And 1)

drugi = bajt And 15
Razlika = prvi Xor drugi
Rezultat = 0

For i = 0 To 15
Maska = Not ((Razlika - i) <> 0)
Rezultat = Rezultat Or (Maska And Sbox0(i))
Next i

prvi = prvi Xor Rezultat
drugi = drugi Xor Rezultat

OR4(3) = -((prvi And 1) <> 0)
OR4(2) = -((prvi And 2) <> 0)
OR4(1) = -((prvi And 4) <> 0)
OR4(0) = -((prvi And 8) <> 0)

prvi = -(OR4(2)) And 8
prvi = prvi + (-OR4(3) And 4)
prvi = prvi + (-(OR4(0) Xor OR4(2)) And 2)
prvi = prvi + (-(OR4(1) Xor OR4(3)) And 1)

Razlika = prvi Xor drugi

Rezultat = 0
For i = 0 To 15
Maska = Not ((Razlika - i) <> 0)
Rezultat = Rezultat Or (Maska And Sbox1(i))
Next i

prvi = prvi Xor Rezultat
drugi = drugi Xor Rezultat

OR4(3) = -((prvi And 1) <> 0)
OR4(2) = -((prvi And 2) <> 0)
OR4(1) = -((prvi And 4) <> 0)
OR4(0) = -((prvi And 8) <> 0)

prvi = -(OR4(2)) And 8
prvi = prvi + (-OR4(3) And 4)
prvi = prvi + (-(OR4(0) Xor OR4(2)) And 2)
prvi = prvi + (-(OR4(1) Xor OR4(3)) And 1)

Razlika = prvi Xor drugi

Rezultat = 0
For i = 0 To 15
Maska = Not ((Razlika - i) <> 0)
Rezultat = Rezultat Or (Maska And Sbox2(i))
Next i

prvi = prvi Xor Rezultat
drugi = drugi Xor Rezultat

OR4(0) = ((prvi And 8) <> 0) And 128
OR4(1) = ((prvi And 4) <> 0) And 64
OR4(2) = ((prvi And 2) <> 0) And 32
OR4(3) = ((prvi And 1) <> 0) And 16

Rezultat = OR4(0) + OR4(1) + OR4(2) + OR4(3)
Rezultat = Rezultat + drugi

RacunajSboxVrijednost = Rezultat And 255
End Function

' Mini-Sbox
Private Function Sbox0(ByVal v As Long) As Long
Dim b(15) As Long
b(0) = 2: b(1) = 5: b(2) = 1: b(3) = 9: b(4) = 14: b(5) = 10: b(6) = 12: b(7) = 8: b(8) = 6: b(9) = 4: b(10) = 7: b(11) = 15: b(12) = 13: b(13) = 11: b(14) = 0: b(15) = 3
Sbox0 = b(v)
End Function

' Mini-SBox
Private Function Sbox1(ByVal v As Long) As Long
Dim b(15) As Long
b(0) = 11: b(1) = 4: b(2) = 1: b(3) = 15: b(4) = 0: b(5) = 3: b(6) = 14: b(7) = 13: b(8) = 10: b(9) = 8: b(10) = 7: b(11) = 5: b(12) = 12: b(13) = 2: b(14) = 9: b(15) = 6
Sbox1 = b(v)
End Function

' Mini-SBox
Private Function Sbox2(ByVal v As Long) As Long
Dim b(15) As Long
b(0) = 13: b(1) = 10: b(2) = 11: b(3) = 1: b(4) = 4: b(5) = 3: b(6) = 8: b(7) = 9: b(8) = 5: b(9) = 7: b(10) = 2: b(11) = 12: b(12) = 15: b(13) = 0: b(14) = 6: b(15) = 14
Sbox2 = b(v)
End Function

' MU8 Function
Private Function MU8(Ulaz() As Long) As Long()
Dim i As Long, a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, h As Long
Dim prvi As Long, drugi As Long, treci As Long, cetvrti As Long, Peti As Long, Sesti As Long, Sedmi As Long, Osmi As Long

a = 0: b = 0: c = 0: d = 0: e = 0: f = 0: g = 0: h = 0
prvi = 0: drugi = 0: treci = 0: cetvrti = 0: Peti = 0: Sesti = 0: Sedmi = 0: Osmi = 0

For i = 0 To UBound(Ulaz) Step 8
a = Ulaz(i): b = Ulaz(i + 1): c = Ulaz(i + 2): d = Ulaz(i + 3): e = Ulaz(i + 4): f = Ulaz(i + 5): g = Ulaz(i + 6): h = MulGF2_8(Ulaz(i + 7), 3)
prvi = a Xor b Xor c Xor d Xor e Xor f Xor g Xor h
a = Ulaz(i): b = MulGF2_8(Ulaz(i + 1), 3): c = MulGF2_8(Ulaz(i + 2), 130): d = MulGF2_8(Ulaz(i + 3), 2): e = MulGF2_8(Ulaz(i + 4), 4): f = MulGF2_8(Ulaz(i + 5), 252): g = MulGF2_8(Ulaz(i + 6), 126): h = Ulaz(i + 7)
drugi = a Xor b Xor c Xor d Xor e Xor f Xor g Xor h
a = MulGF2_8(Ulaz(i), 3): b = MulGF2_8(Ulaz(i + 1), 130): c = MulGF2_8(Ulaz(i + 2), 2): d = MulGF2_8(Ulaz(i + 3), 4): e = MulGF2_8(Ulaz(i + 4), 252): f = MulGF2_8(Ulaz(i + 5), 126): g = Ulaz(i + 6): h = Ulaz(i + 7)
treci = a Xor b Xor c Xor d Xor e Xor f Xor g Xor h
a = MulGF2_8(Ulaz(i), 130): b = MulGF2_8(Ulaz(i + 1), 2): c = MulGF2_8(Ulaz(i + 2), 4): d = MulGF2_8(Ulaz(i + 3), 252): e = MulGF2_8(Ulaz(i + 4), 126): f = Ulaz(i + 5): g = MulGF2_8(Ulaz(i + 6), 3): h = Ulaz(i + 7)
cetvrti = a Xor b Xor c Xor d Xor e Xor f Xor g Xor h
a = MulGF2_8(Ulaz(i), 2): b = MulGF2_8(Ulaz(i + 1), 4): c = MulGF2_8(Ulaz(i + 2), 252): d = MulGF2_8(Ulaz(i + 3), 126): e = Ulaz(i + 4): f = MulGF2_8(Ulaz(i + 5), 3): g = MulGF2_8(Ulaz(i + 6), 130): h = Ulaz(i + 7)
Peti = a Xor b Xor c Xor d Xor e Xor f Xor g Xor h
a = MulGF2_8(Ulaz(i), 4): b = MulGF2_8(Ulaz(i + 1), 252): c = MulGF2_8(Ulaz(i + 2), 126): d = Ulaz(i + 3): e = MulGF2_8(Ulaz(i + 4), 3): f = MulGF2_8(Ulaz(i + 5), 130): g = MulGF2_8(Ulaz(i + 6), 2): h = Ulaz(i + 7)
Sesti = a Xor b Xor c Xor d Xor e Xor f Xor g Xor h
a = MulGF2_8(Ulaz(i), 252): b = MulGF2_8(Ulaz(i + 1), 126): c = Ulaz(i + 2): d = MulGF2_8(Ulaz(i + 3), 3): e = MulGF2_8(Ulaz(i + 4), 130): f = MulGF2_8(Ulaz(i + 5), 2): g = MulGF2_8(Ulaz(i + 6), 4): h = Ulaz(i + 7)
Sedmi = a Xor b Xor c Xor d Xor e Xor f Xor g Xor h
a = MulGF2_8(Ulaz(i), 126): b = Ulaz(i + 1): c = MulGF2_8(Ulaz(i + 2), 3): d = MulGF2_8(Ulaz(i + 3), 130): e = MulGF2_8(Ulaz(i + 4), 2): f = MulGF2_8(Ulaz(i + 5), 4): g = MulGF2_8(Ulaz(i + 6), 252): h = Ulaz(i + 7)
Osmi = a Xor b Xor c Xor d Xor e Xor f Xor g Xor h

Ulaz(i) = prvi: Ulaz(i + 1) = drugi: Ulaz(i + 2) = treci: Ulaz(i + 3) = cetvrti: Ulaz(i + 4) = Peti: Ulaz(i + 5) = Sesti: Ulaz(i + 6) = Sedmi: Ulaz(i + 7) = Osmi
Next i

MU8 = Ulaz
End Function

' Mix128 Function
Private Function Mix128(Ulaz() As Long) As Long()
Dim PrviBlok(7) As Long, DrugiBlok(7) As Long, TreciBlok(7) As Long, CetvrtiBlok(7) As Long, i As Long

For i = 0 To 7
PrviBlok(i) = Ulaz(i + 8) Xor Ulaz(i + 16) Xor Ulaz(i + 24)
DrugiBlok(i) = Ulaz(i) Xor Ulaz(i + 16) Xor Ulaz(i + 24)
TreciBlok(i) = Ulaz(i) Xor Ulaz(i + 8) Xor Ulaz(i + 24)
CetvrtiBlok(i) = Ulaz(i) Xor Ulaz(i + 8) Xor Ulaz(i + 16)
Ulaz(i) = PrviBlok(i)
Ulaz(i + 8) = DrugiBlok(i)
Ulaz(i + 16) = TreciBlok(i)
Ulaz(i + 24) = CetvrtiBlok(i)
Next i

Mix128 = Ulaz
End Function

' NL128 Padding Function
Private Function NL128Konstanta(Mix_128() As Long) As Long()
Dim Konstanta(31) As Long, i As Long
Konstanta(0) = 183:  Konstanta(16) = 98
Konstanta(1) = 225:  Konstanta(17) = 231
Konstanta(2) = 81:   Konstanta(18) = 22
Konstanta(3) = 98:   Konstanta(19) = 15
Konstanta(4) = 138:  Konstanta(20) = 56
Konstanta(5) = 237:  Konstanta(21) = 180
Konstanta(6) = 42:   Konstanta(22) = 218
Konstanta(7) = 106:  Konstanta(23) = 86
Konstanta(8) = 191:  Konstanta(24) = 167
Konstanta(9) = 113:  Konstanta(25) = 132
Konstanta(10) = 88:  Konstanta(26) = 217
Konstanta(11) = 128: Konstanta(27) = 4
Konstanta(12) = 156: Konstanta(28) = 81
Konstanta(13) = 244: Konstanta(29) = 144
Konstanta(14) = 243: Konstanta(30) = 207
Konstanta(15) = 199: Konstanta(31) = 239

For i = 0 To 31
Konstanta(i) = Mix_128(i) Xor Konstanta(i)
Next i
 
NL128Konstanta = Konstanta
End Function

' NL128 Bit Flip Function ( If k = ek )
Private Function PreokreniBitove(NizKonstanta() As Long, Maska As Long) As Long()
Dim i As Long

For i = 0 To 31
NizKonstanta(i) = ((255 - NizKonstanta(i)) And Maska) Or (NizKonstanta(i) And (Not Maska))
Next i

PreokreniBitove = NizKonstanta
End Function

' ELMOR128 Input Function
Private Function ELMOR128Ulaz(KomplementNiz() As Long) As Long()
Dim i As Long

For i = 0 To 7
KomplementNiz(i) = KomplementNiz(i) Xor KomplementNiz(i + 16)
Next i

For i = 8 To 15
KomplementNiz(i) = KomplementNiz(i) Xor KomplementNiz(i + 16)
Next i

ELMOR128Ulaz = KomplementNiz
End Function

' ELMOR128 Function
Private Function ELMOR128(UlazniNiz() As Long, Kljuc() As Long) As Long()
Dim F64Ulaz(7) As Long, i As Long, Izlaz() As Long, ORTransformacija(3) As Long

For i = 0 To 7
F64Ulaz(i) = UlazniNiz((i And 4) + i) Xor UlazniNiz((i And 4) + i + 4) Xor Kljuc(i)
Next i

Izlaz = ObradaSBoxUlaza(F64Ulaz)
Izlaz = MU8(Izlaz)

For i = 0 To 7
Izlaz(i) = Izlaz(i) Xor Kljuc(i + 8)
Next i

Izlaz = ObradaSBoxUlaza(Izlaz)

For i = 0 To 7
Izlaz(i) = Izlaz(i) Xor Kljuc(i)
Next i

For i = 0 To 7
UlazniNiz(i) = UlazniNiz(i) Xor Izlaz(i And 3)
UlazniNiz(i + 8) = UlazniNiz(i + 8) Xor Izlaz((i And 3) + 4)
Next i

ORTransformacija(0) = UlazniNiz(2): ORTransformacija(1) = UlazniNiz(3): ORTransformacija(2) = UlazniNiz(0) Xor UlazniNiz(2): ORTransformacija(3) = UlazniNiz(1) Xor UlazniNiz(3)
UlazniNiz(0) = ORTransformacija(0): UlazniNiz(1) = ORTransformacija(1): UlazniNiz(2) = ORTransformacija(2): UlazniNiz(3) = ORTransformacija(3)

ORTransformacija(0) = UlazniNiz(2 + 8): ORTransformacija(1) = UlazniNiz(3 + 8): ORTransformacija(2) = UlazniNiz(0 + 8) Xor UlazniNiz(2 + 8): ORTransformacija(3) = UlazniNiz(1 + 8) Xor UlazniNiz(3 + 8)
UlazniNiz(0 + 8) = ORTransformacija(0): UlazniNiz(1 + 8) = ORTransformacija(1): UlazniNiz(2 + 8) = ORTransformacija(2): UlazniNiz(3 + 8) = ORTransformacija(3)

ELMOR128 = UlazniNiz
End Function

'ELMID128 Function
Private Function ELMID128(UlazniNiz() As Long, Kljuc() As Long) As Long()
Dim F64Ulaz(7) As Long, i As Long, Izlaz() As Long

For i = 0 To 7
F64Ulaz(i) = UlazniNiz((i And 4) + i) Xor UlazniNiz((i And 4) + i + 4) Xor Kljuc(i)
Next i

Izlaz = ObradaSBoxUlaza(F64Ulaz)
Izlaz = MU8(Izlaz)

For i = 0 To 7
Izlaz(i) = Izlaz(i) Xor Kljuc(i + 8)
Next i

Izlaz = ObradaSBoxUlaza(Izlaz)

For i = 0 To 7
Izlaz(i) = Izlaz(i) Xor Kljuc(i)
Next i

For i = 0 To 7
UlazniNiz(i) = UlazniNiz(i) Xor Izlaz(i And 3)
UlazniNiz(i + 8) = UlazniNiz(i + 8) Xor Izlaz((i And 3) + 4)
Next i

ELMID128 = UlazniNiz
End Function

' ELMIO128 Function
Private Function ELMIO128(Ulaz() As Long, Kljuc() As Long) As Long()
Dim F64Ulaz(7) As Long, IOPromjena(3) As Long, i As Long, Izlaz() As Long

IOPromjena(0) = Ulaz(0) Xor Ulaz(2): IOPromjena(1) = Ulaz(1) Xor Ulaz(3): IOPromjena(2) = Ulaz(0): IOPromjena(3) = Ulaz(1)
Ulaz(0) = IOPromjena(0): Ulaz(1) = IOPromjena(1): Ulaz(2) = IOPromjena(2): Ulaz(3) = IOPromjena(3)

IOPromjena(0) = Ulaz(0 + 8) Xor Ulaz(2 + 8): IOPromjena(1) = Ulaz(1 + 8) Xor Ulaz(3 + 8): IOPromjena(2) = Ulaz(0 + 8): IOPromjena(3) = Ulaz(1 + 8)
Ulaz(0 + 8) = IOPromjena(0): Ulaz(1 + 8) = IOPromjena(1): Ulaz(2 + 8) = IOPromjena(2): Ulaz(3 + 8) = IOPromjena(3)

For i = 0 To 7
F64Ulaz(i) = Ulaz((i And 4) + i) Xor Ulaz((i And 4) + i + 4) Xor Kljuc(i)
Next i

Izlaz = ObradaSBoxUlaza(F64Ulaz)
Izlaz = MU8(Izlaz)

For i = 0 To 7
Izlaz(i) = Izlaz(i) Xor Kljuc(i + 8)
Next i

Izlaz = ObradaSBoxUlaza(Izlaz)

For i = 0 To 7
Izlaz(i) = Izlaz(i) Xor Kljuc(i)
Next i

For i = 0 To 7
Ulaz(i) = Ulaz(i) Xor Izlaz(i And 3)
Ulaz(i + 8) = Ulaz(i + 8) Xor Izlaz((i And 3) + 4)
Next i

ELMIO128 = Ulaz
End Function

' NL128 Function
Private Function NL128(PKljuc() As Long, BrojRunde As Long, DuljinaKljuca As Long) As Long()
Dim Izlaz() As Long, ElmorKljuc(15) As Long, ElmidKljuc(15) As Long, i As Long

Izlaz = MKljuc(PKljuc, DuljinaKljuca)
Izlaz = DKljuc(Izlaz, BrojRunde)

For i = 0 To 15
ElmorKljuc(i) = Izlaz(i)
ElmidKljuc(i) = Izlaz(i + 16)
Next i

Izlaz = ObradaSBoxUlaza(Izlaz)
Izlaz = MU8(Izlaz)
Izlaz = Mix128(Izlaz)
Izlaz = NL128Konstanta(Izlaz)
Izlaz = PreokreniBitove(Izlaz, DuljinaKljuca)
Izlaz = ObradaSBoxUlaza(Izlaz)
Izlaz = ELMOR128Ulaz(Izlaz)
Izlaz = ELMOR128(Izlaz, ElmorKljuc)
Izlaz = ELMID128(Izlaz, ElmidKljuc)
NL128 = Izlaz
End Function

' Encryption Function
Public Function Sifriraj(Ulaz() As Long, Kljuc() As Long, DuljinaKljuca As Long) As Long()
Dim Sifrirano() As Long, Blok(15) As Long, i As Long, j As Long, KljucRunde() As Long, k As Long

For i = 0 To 14
   
   KljucRunde = NL128(Kljuc, i, DuljinaKljuca)
   
   For j = 0 To UBound(Ulaz) Step 16
   Blok(0) = Ulaz(j): Blok(1) = Ulaz(j + 1): Blok(2) = Ulaz(j + 2): Blok(3) = Ulaz(j + 3): Blok(4) = Ulaz(j + 4): Blok(5) = Ulaz(j + 5): Blok(6) = Ulaz(j + 6): Blok(7) = Ulaz(j + 7)
   Blok(8) = Ulaz(j + 8): Blok(9) = Ulaz(j + 9): Blok(10) = Ulaz(j + 10): Blok(11) = Ulaz(j + 11): Blok(12) = Ulaz(j + 12): Blok(13) = Ulaz(j + 13): Blok(14) = Ulaz(j + 14): Blok(15) = Ulaz(j + 15)
   
   Sifrirano = ELMOR128(Blok, KljucRunde)
   
   Ulaz(j) = Sifrirano(0): Ulaz(j + 1) = Sifrirano(1): Ulaz(j + 2) = Sifrirano(2): Ulaz(j + 3) = Sifrirano(3): Ulaz(j + 4) = Sifrirano(4): Ulaz(j + 5) = Sifrirano(5): Ulaz(j + 6) = Sifrirano(6): Ulaz(j + 7) = Sifrirano(7)
   Ulaz(j + 8) = Sifrirano(8): Ulaz(j + 9) = Sifrirano(9): Ulaz(j + 10) = Sifrirano(10): Ulaz(j + 11) = Sifrirano(11): Ulaz(j + 12) = Sifrirano(12): Ulaz(j + 13) = Sifrirano(13): Ulaz(j + 14) = Sifrirano(14): Ulaz(j + 15) = Sifrirano(15)
   Next j
   
Next i

KljucRunde = NL128(Kljuc, i, DuljinaKljuca)

For j = 0 To UBound(Ulaz) Step 16
Blok(0) = Ulaz(j): Blok(1) = Ulaz(j + 1): Blok(2) = Ulaz(j + 2): Blok(3) = Ulaz(j + 3): Blok(4) = Ulaz(j + 4): Blok(5) = Ulaz(j + 5): Blok(6) = Ulaz(j + 6): Blok(7) = Ulaz(j + 7)
Blok(8) = Ulaz(j + 8): Blok(9) = Ulaz(j + 9): Blok(10) = Ulaz(j + 10): Blok(11) = Ulaz(j + 11): Blok(12) = Ulaz(j + 12): Blok(13) = Ulaz(j + 13): Blok(14) = Ulaz(j + 14): Blok(15) = Ulaz(j + 15)
Sifrirano = ELMID128(Blok, KljucRunde)
   
Ulaz(j) = Sifrirano(0): Ulaz(j + 1) = Sifrirano(1): Ulaz(j + 2) = Sifrirano(2): Ulaz(j + 3) = Sifrirano(3): Ulaz(j + 4) = Sifrirano(4): Ulaz(j + 5) = Sifrirano(5): Ulaz(j + 6) = Sifrirano(6): Ulaz(j + 7) = Sifrirano(7)
Ulaz(j + 8) = Sifrirano(8): Ulaz(j + 9) = Sifrirano(9): Ulaz(j + 10) = Sifrirano(10): Ulaz(j + 11) = Sifrirano(11): Ulaz(j + 12) = Sifrirano(12): Ulaz(j + 13) = Sifrirano(13): Ulaz(j + 14) = Sifrirano(14): Ulaz(j + 15) = Sifrirano(15)
Next j

Sifriraj = Ulaz
End Function

' Decryption Function
Public Function Desifriraj(Ulaz() As Long, Kljuc() As Long, DuljinaKljuca As Long) As Long()
Dim Desifrirano() As Long, Blok(15) As Long, i As Long, j As Long, KljucRunde() As Long, k As Long

For i = 15 To 15 Step -1

KljucRunde = NL128(Kljuc, i, DuljinaKljuca)

For j = 0 To UBound(Ulaz) Step 16
Blok(0) = Ulaz(j): Blok(1) = Ulaz(j + 1): Blok(2) = Ulaz(j + 2): Blok(3) = Ulaz(j + 3): Blok(4) = Ulaz(j + 4): Blok(5) = Ulaz(j + 5): Blok(6) = Ulaz(j + 6): Blok(7) = Ulaz(j + 7)
Blok(8) = Ulaz(j + 8): Blok(9) = Ulaz(j + 9): Blok(10) = Ulaz(j + 10): Blok(11) = Ulaz(j + 11): Blok(12) = Ulaz(j + 12): Blok(13) = Ulaz(j + 13): Blok(14) = Ulaz(j + 14): Blok(15) = Ulaz(j + 15)

Desifrirano = ELMID128(Blok, KljucRunde)
   
Ulaz(j) = Desifrirano(0): Ulaz(j + 1) = Desifrirano(1): Ulaz(j + 2) = Desifrirano(2): Ulaz(j + 3) = Desifrirano(3): Ulaz(j + 4) = Desifrirano(4): Ulaz(j + 5) = Desifrirano(5): Ulaz(j + 6) = Desifrirano(6): Ulaz(j + 7) = Desifrirano(7)
Ulaz(j + 8) = Desifrirano(8): Ulaz(j + 9) = Desifrirano(9): Ulaz(j + 10) = Desifrirano(10): Ulaz(j + 11) = Desifrirano(11): Ulaz(j + 12) = Desifrirano(12): Ulaz(j + 13) = Desifrirano(13): Ulaz(j + 14) = Desifrirano(14): Ulaz(j + 15) = Desifrirano(15)
Next j

Next i

For i = 14 To 0 Step -1

KljucRunde = NL128(Kljuc, i, DuljinaKljuca)

   For j = 0 To UBound(Ulaz) Step 16

   Blok(0) = Ulaz(j): Blok(1) = Ulaz(j + 1): Blok(2) = Ulaz(j + 2): Blok(3) = Ulaz(j + 3): Blok(4) = Ulaz(j + 4): Blok(5) = Ulaz(j + 5): Blok(6) = Ulaz(j + 6): Blok(7) = Ulaz(j + 7)
   Blok(8) = Ulaz(j + 8): Blok(9) = Ulaz(j + 9): Blok(10) = Ulaz(j + 10): Blok(11) = Ulaz(j + 11): Blok(12) = Ulaz(j + 12): Blok(13) = Ulaz(j + 13): Blok(14) = Ulaz(j + 14): Blok(15) = Ulaz(j + 15)

   Desifrirano = ELMIO128(Blok, KljucRunde)

   Ulaz(j) = Desifrirano(0): Ulaz(j + 1) = Desifrirano(1): Ulaz(j + 2) = Desifrirano(2): Ulaz(j + 3) = Desifrirano(3): Ulaz(j + 4) = Desifrirano(4): Ulaz(j + 5) = Desifrirano(5): Ulaz(j + 6) = Desifrirano(6): Ulaz(j + 7) = Desifrirano(7)
   Ulaz(j + 8) = Desifrirano(8): Ulaz(j + 9) = Desifrirano(9): Ulaz(j + 10) = Desifrirano(10): Ulaz(j + 11) = Desifrirano(11): Ulaz(j + 12) = Desifrirano(12): Ulaz(j + 13) = Desifrirano(13): Ulaz(j + 14) = Desifrirano(14): Ulaz(j + 15) = Desifrirano(15)

   Next j

Next i

Desifriraj = Ulaz
End Function
