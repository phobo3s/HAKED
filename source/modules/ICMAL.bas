Option Compare Database
Option Explicit
Sub addMetraj(wbook As Excel.Workbook)

    Dim icmalName As String
    icmalName = wbook.ActiveSheet.name
    Dim rowstr As String
    Dim j As Integer
    Dim k As Integer
    Dim a As Integer
    
    For k = 0 To (wbook.ActiveSheet.Range("t3").Value - 1)
        If k > 0 Then
        a = 1
        Else
        a = 0
        End If
        For j = (a + (31 * k)) To (16 + (31 * k))
            If wbook.ActiveSheet.Range("b9").Offset(j, 0).Value <> "" Or wbook.ActiveSheet.Range("d9").Offset(j, 0).Value <> "" Then
                rowstr = rowstr & "," & wbook.ActiveSheet.Range("b9").Offset(j, 0).Row
            Else
            End If
        Next j
    Next k
    rowstr = Right(rowstr, Len(rowstr) - 1)
    
    Dim metrCount As Long
    metrCount = Len(rowstr) - Len(Replace(rowstr, ",", "")) + 1
    Dim rownums
    rownums = Split(rowstr, ",")
    
    Dim i As Integer
    Dim rowNum As Long

For i = 0 To (metrCount - 1)

wbook.Worksheets(icmalName).Activate
rowNum = rownums(i)
    If rowNum = 0 Then
    Exit Sub
    Else
    End If
    
Dim siraNo As String
Dim siraNum As Integer
Dim pozNo As String
Dim maliyetKodu As String
Dim acik As String
Dim birim As String
Dim birimFiyat As String
Dim sozMik As String
Dim sozBedel As String
Dim binaadi As String
Dim paraBirim As String
Dim benNum As String

With wbook.ActiveSheet
siraNum = .Range("a1").Offset((rowNum - 1), 0).Value
siraNo = "'" & .name & "'!" & .Range("a1").Offset((rowNum - 1), 0).Address
pozNo = "'" & .name & "'!" & .Range("b1").Offset((rowNum - 1), 0).Address
maliyetKodu = "'" & .name & "'!" & .Range("c1").Offset((rowNum - 1), 0).Address
acik = "'" & .name & "'!" & .Range("d1").Offset((rowNum - 1), 0).Address
birim = "'" & .name & "'!" & .Range("f1").Offset((rowNum - 1), 0).Address
birimFiyat = "'" & .name & "'!" & .Range("g1").Offset((rowNum - 1), 0).Address
sozMik = "'" & .name & "'!" & .Range("h1").Offset((rowNum - 1), 0).Address
sozBedel = "'" & .name & "'!" & .Range("i1").Offset((rowNum - 1), 0).Address
binaadi = Mid(.name, 14, Len(.name) - 17)
paraBirim = Right(.name, 3)
End With

benNum = wbook.ActiveSheet.name

'yer belirleme!!
wbook.Sheets("METRAJ_SBLN").Visible = True
wbook.Sheets("METRAJ_SBLN").Copy After:=wbook.Worksheets(wbook.Sheets(9 + i).Index)
wbook.Sheets("METRAJ_SBLN").Visible = xlVeryHidden
wbook.ActiveSheet.name = siraNum & "_" & binaadi & "_" & paraBirim
wbook.ActiveSheet.Tab.ColorIndex = 12

With wbook.ActiveSheet
.Range("c8") = "=" & pozNo
.Range("i8") = "=" & maliyetKodu
.Range("c9") = "=" & acik
.Range("e13") = "=" & birim
.Range("m8") = binaadi
End With

'isimli alaný belirleme
Dim rangeNameK As String 'kümülatif
Dim rangeNameO As String 'önceki
Dim rangeNameB As String 'bu hakediþ

rangeNameK = "M_" & wbook.ActiveSheet.name & "_K"
rangeNameO = "M_" & wbook.ActiveSheet.name & "_O"
rangeNameB = "M_" & wbook.ActiveSheet.name & "_B"

Dim sayfAdi As String
sayfAdi = wbook.ActiveSheet.name
'sayfAdi = Replace(sayfAdi, "-", "_")
rangeNameK = Replace(rangeNameK, "-", "_")
rangeNameO = Replace(rangeNameO, "-", "_")
rangeNameB = Replace(rangeNameB, "-", "_")

'delete Names
Dim name As name

For Each name In wbook.Names
    If name.Parent.name = wbook.ActiveSheet.name Then
        name.Delete
    Else
    End If
Next

Dim newName As String

'1
newName = "=" & "'" & sayfAdi & "'" & "!" & wbook.ActiveSheet.Range("k49").Address
newName = Replace(newName, "chr(34)", "")
wbook.Names.Add rangeNameK, newName
'2
newName = "=" & "'" & sayfAdi & "'" & "!" & wbook.ActiveSheet.Range("L49").Address
newName = Replace(newName, "chr(34)", "")
wbook.Names.Add rangeNameO, newName
'3
newName = "=" & "'" & sayfAdi & "'" & "!" & wbook.ActiveSheet.Range("M49").Address
newName = Replace(newName, "chr(34)", "")
wbook.Names.Add rangeNameB, newName

wbook.ActiveSheet.PageSetup.PrintArea = "A1:M55"

wbook.Worksheets(benNum).Range("j1").Offset(rowNum - 1, 0).Value = "=" & rangeNameO
wbook.Worksheets(benNum).Range("K1").Offset(rowNum - 1, 0).Value = "=" & rangeNameB
wbook.Worksheets(benNum).Range("L1").Offset(rowNum - 1, 0).Value = "=" & rangeNameK

Next i
End Sub

Sub addIcmalPage()

'Ýcmal sayfasý ekleyicik
Dim i As Long
Dim sayfSayi As Long
Dim offsCount As Long
offsCount = 32

For i = 5 To 1048576 Step (offsCount - 1)
        If ActiveSheet.Range("R" & i).Value = "" Then
            Exit For
        Else
            sayfSayi = sayfSayi + 1
        End If
Next i
Range("t3").Value = sayfSayi

ActiveSheet.Range("A1:R31").Copy ActiveSheet.Range("a1").Offset((sayfSayi * (offsCount - 1)), 0)
For i = 0 To offsCount - 1
    ActiveSheet.Range("A1").Offset((sayfSayi * (offsCount - 1)) + i, 0).RowHeight = ActiveSheet.Range("A1").Offset(i, 0).EntireRow.Height
Next i

sayfSayi = sayfSayi + 1
Application.ActiveSheet.PageSetup.PrintArea = "A1:R" & (sayfSayi * (offsCount - 1))

Range("t3").Value = sayfSayi

'Sayfa Temizliði
Range("B9:H25").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = ""
Range("J9:L25").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = ""

'yeni sayfa Ek düzenlemeler
Range("R5").Offset((sayfSayi - 1) * (offsCount - 1), 0).Formula = "=R" & 5 + ((sayfSayi - 2) * (offsCount - 1)) & "+1"
Range("A9").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = ""
Range("B9").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = "ÖNCEKÝ SAYFADAN GELEN"

Range("H9").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = "=H" & Range("H26").Offset((sayfSayi - 2) * (offsCount - 1), 0).Row
Range("I9").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = "=I" & Range("I26").Offset((sayfSayi - 2) * (offsCount - 1), 0).Row
Range("J9").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = "=J" & Range("J26").Offset((sayfSayi - 2) * (offsCount - 1), 0).Row
Range("K9").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = "=K" & Range("K26").Offset((sayfSayi - 2) * (offsCount - 1), 0).Row
Range("L9").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = "=L" & Range("L26").Offset((sayfSayi - 2) * (offsCount - 1), 0).Row
Range("P9").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = "=P" & Range("P26").Offset((sayfSayi - 2) * (offsCount - 1), 0).Row
Range("Q9").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = "=Q" & Range("Q26").Offset((sayfSayi - 2) * (offsCount - 1), 0).Row
Range("R9").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = "=R" & Range("R26").Offset((sayfSayi - 2) * (offsCount - 1), 0).Row
Range("A10").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = "=A" & Range("A25").Offset(((sayfSayi - 2) * (offsCount - 1)), 0).Row & "+1"
Range("e5").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = "=$E$5"
Range("A5").Offset((sayfSayi - 1) * (offsCount - 1), 0).Value = "=$A$3"
Range("B9:G9").Offset((sayfSayi - 1) * (offsCount - 1), 0).Select
'Selection.MergeCells
'HÜCRELERÝ BÝRLEÞTÝR VE FORMATLA.


End Sub
