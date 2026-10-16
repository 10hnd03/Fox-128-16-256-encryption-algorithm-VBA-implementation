VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObradaTeksta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare PtrSafe Sub SecureZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByVal Pohrana As LongPtr, ByVal BrojElemenata As LongPtr)
Private Declare PtrSafe Function GetProcessHeap Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function HeapAlloc Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function HeapFree Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal lpMem As LongPtr) As Boolean
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function VarPtrArray Lib "VBE7" Alias "VarPtr" (ByRef Var() As Any) As LongPtr
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Type SAFEARRAYGRANICA
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    Padding    As Long
    pvData     As LongPtr
    Bounds     As SAFEARRAYGRANICA
End Type

' Function that takes input string independetly of length
Function ObradaStringa(ByRef Ulaz As String) As Long() 'As Integer()


' Izracun maske na osnovi duljine stringa
Dim DuljinaStringa As Long: DuljinaStringa = Len(Ulaz) ': Debug.Print "DS: " & duljinaStringa
Dim LongMaska As Long, LongPtrMaska As LongPtr: LongMaska = DuljinaStringa <> 0: LongPtrMaska = LongMaska ': Debug.Print "LM i LPTRM: " & LongMaska & " "; LongPtrMaska

' Priprema praznog polja
Dim PraznoPolje(0 To 15) As Integer, AdresaPraznogPolja As LongPtr
AdresaPraznogPolja = VarPtr(PraznoPolje(0)) ': Debug.Print "APP: " & AdresaPraznogPolja

' Priprema kernel praznog polja
Dim AdresaKernelPolja As LongPtr, KernelPolje As LongPtr, DuljinaKernelPolja As Long
DuljinaKernelPolja = DuljinaStringa + 15 And -16
AdresaKernelPolja = GetProcessHeap: KernelPolje = HeapAlloc(AdresaKernelPolja, 8, DuljinaKernelPolja + DuljinaKernelPolja)
'Debug.Print "Poèetak i duljina kernel polja: " & AdresaKernelPolja & " " & DuljinaKernelPolja

' Kopiranje ulaznog stringa u kernel polje
CopyMemory ByVal KernelPolje, ByVal StrPtr(Ulaz), DuljinaStringa + DuljinaStringa

' Izraèun ciljne adrese i duljine polja
Dim CiljnaAdresa As LongPtr, CiljnaDuljinaPolja As Long
CiljnaAdresa = (LongPtrMaska And KernelPolje) Or (AdresaPraznogPolja And Not LongPtrMaska)
CiljnaDuljinaPolja = (DuljinaKernelPolja And LongMaska) Or (16 And Not LongMaska)

' Priprema SAFE ARRAY strukture i praznog integer polja

Dim SA As SAFEARRAY
    SA.cDims = 1
    SA.cbElements = 2 ' Integer (2 bajta)
    SA.pvData = CiljnaAdresa
    SA.Bounds.cElements = CiljnaDuljinaPolja
    SA.Bounds.lLbound = 0
    
Dim IntegerPolje() As Integer
ReDim LongSpremnik(CiljnaDuljinaPolja - 1) As Long

' Povezivanje Integer polja sa SAFE ARRAY strukturom
Dim AdresaIntegerPolja As LongPtr, AdresaSA As LongPtr: AdresaIntegerPolja = VarPtrArray(IntegerPolje): AdresaSA = VarPtr(SA)
CopyMemory ByVal AdresaIntegerPolja, AdresaSA, 8

' Direktan pristup string vrijednostima unutar granica kernel polja
Dim i As Long, Rezultat As Long

For i = 0 To UBound(IntegerPolje)
LongSpremnik(i) = CLng(IntegerPolje(i))
Next i

' Maskiranje UTF16 u ASCII konverzije
For i = 0 To UBound(LongSpremnik)

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


ObradaStringa = LongSpremnik

' Èišæenje memorije
Dim Nula As LongPtr: Nula = 0
CopyMemory ByVal AdresaIntegerPolja, Nula, 8 ' Postavljamo adresu integer polja na 0
    
' Kernel memoriju oslobaðamo jer je alocirana na poèetku
HeapFree AdresaKernelPolja, 0, KernelPolje
End Function


