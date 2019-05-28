Version =21
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =162
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =18024
    DatasheetFontHeight =11
    ItemSuffix =101
    Right =26160
    Bottom =12312
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x6d518d598e49e540
    End
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =162
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            TextFontCharSet =162
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            TextFontCharSet =162
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ListBox
            TextFontCharSet =162
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            TextFontCharSet =162
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =7026
            Name ="Ayrıntı"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =228
                    Top =3576
                    Width =17700
                    Height =3300
                    TabIndex =4
                    BorderColor =10921638
                    Name ="Alt12"
                    SourceObject ="Query.SRHAKED"
                    GroupTable =1
                    LeftPadding =90
                    RightPadding =90
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =228
                    LayoutCachedTop =3576
                    LayoutCachedWidth =17928
                    LayoutCachedHeight =6876
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =228
                    Top =2952
                    Width =3804
                    Height =348
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Açılan_Kutu62"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [TBISLER].[Kimlik], [TBISLER].[ISADI] FROM TBISLER ORDER BY [ISADI]; "
                    ColumnWidths ="0;1440"
                    OnChange ="[Event Procedure]"
                    GroupTable =4
                    GridlineColor =10921638

                    LayoutCachedLeft =228
                    LayoutCachedTop =2952
                    LayoutCachedWidth =4032
                    LayoutCachedHeight =3300
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    GroupTable =4
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =4200
                    Top =2952
                    Width =1308
                    Height =300
                    TabIndex =2
                    BoundColumn =1
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="Açılan_Kutu78"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TBHAKED.Kimlik, TBHAKED.HAKEDISNO, TBHAKED.ISADI FROM TBHAKED WHERE (((TB"
                        "HAKED.ISADI)=[FORMS].[FRMAIN].[Açılan_Kutu62])) ORDER BY TBHAKED.HAKEDISNO; "
                    ColumnWidths ="0;1440"
                    OnChange ="[Event Procedure]"
                    GroupTable =6
                    GridlineColor =10921638

                    LayoutCachedLeft =4200
                    LayoutCachedTop =2952
                    LayoutCachedWidth =5508
                    LayoutCachedHeight =3252
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1360
                    Top =1133
                    Width =576
                    Height =576
                    ForeColor =4210752
                    Name ="Komut90"
                    Caption ="Komut90"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Uygulamadan Çık"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000003255d6273255d68d ,
                        0x3255d6cf3255d6ff3255d6ff3255d6cf3255d68d3255d6270000000000000000 ,
                        0x00000000000000000000000000000000000000003255d6723255d6f63255d6ff ,
                        0x3255d6ff3255d6ff3255d6ff3255d6ff3255d6ff3255d6f63255d67200000000 ,
                        0x0000000000000000000000003255d6063255d6b73255d6ff3255d6ff3255d6ff ,
                        0x3255d6ff3255d6ff3255d6ff3255d6ff3255d6ff3255d6ff3255d6ff3255d6b7 ,
                        0x3255d60600000000000000003255d6933255d6ff3255d6ff3759d7f94d6bdbe5 ,
                        0x3255d6ff3255d6ff3255d6ff3255d6ff4d6bdbe53759d7f93255d6ff3255d6ff ,
                        0x3255d690000000003255d62d3255d6fc3255d6ff3c5ed8f4eef1fccefcfcfee5 ,
                        0x4a69dbe73255d6ff3255d6ff4a69dbe7fcfcfee5eef1fcce3c5ed8f43255d6ff ,
                        0x3255d6fc3255d62d3255d6933255d6ff3255d6ff3457d6fce4e9fac8ffffffff ,
                        0xfafbfee04766dae94766dae9fafbfee0ffffffffe4e9fac83759d7f93255d6ff ,
                        0x3255d6ff3255d6903255d6db3255d6ff3255d6ff3255d6ff3759d7f9e8ecfaca ,
                        0xfffffffff9fafedcf8f9fedaffffffffe8ecfaca3759d7f93255d6ff3255d6ff ,
                        0x3255d6ff3255d6d53255d6f93255d6ff3255d6ff3255d6ff3255d6ff395bd7f6 ,
                        0xeceffbcdffffffffffffffffeceffbcd395bd7f63255d6ff3255d6ff3255d6ff ,
                        0x3255d6ff3255d6f33255d6f93255d6ff3255d6ff3255d6ff3255d6ff395bd7f6 ,
                        0xf2f4fcd3fffffffffffffffff2f4fcd33c5ed8f43255d6ff3255d6ff3255d6ff ,
                        0x3255d6ff3255d6f03255d6d83255d6ff3255d6ff3255d6ff395bd7f6eff2fcd0 ,
                        0xfffffffff5f6fdd4f5f6fdd4ffffffffeff2fcd0395bd7f63255d6ff3255d6ff ,
                        0x3255d6ff3255d6d53255d6903255d6ff3255d6ff3759d7f9ebeefbcbffffffff ,
                        0xf8f9feda4162d9ee4162d9eef8f9fedaffffffffebeefbcb3759d7f93255d6ff ,
                        0x3255d6ff3255d68d3255d62d3255d6fc3255d6ff395bd7f6ebeefbcbf9fafede ,
                        0x4464daec3255d6ff3255d6ff4464daecf9fafedeebeefbcb395bd7f63255d6ff ,
                        0x3255d6fc3255d62a000000003255d6903255d6ff3255d6ff3759d7f94766dae9 ,
                        0x3255d6ff3255d6ff3255d6ff3255d6ff4766dae93759d7f93255d6ff3255d6ff ,
                        0x3255d68d00000000000000003255d6063255d6b73255d6ff3255d6ff3255d6ff ,
                        0x3255d6ff3255d6ff3255d6ff3255d6ff3255d6ff3255d6ff3255d6ff3255d6b7 ,
                        0x3255d606000000000000000000000000000000003255d6723255d6f63255d6ff ,
                        0x3255d6ff3255d6ff3255d6ff3255d6ff3255d6ff3255d6f63255d67200000000 ,
                        0x0000000000000000000000000000000000000000000000003255d6273255d68d ,
                        0x3255d6cc3255d6fc3255d6fc3255d6cc3255d68d3255d6270000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =1360
                    LayoutCachedTop =1133
                    LayoutCachedWidth =1936
                    LayoutCachedHeight =1709
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3852
                    Top =1476
                    Width =1056
                    Height =288
                    TabIndex =3
                    ForeColor =4210752
                    Name ="Komut95"
                    Caption ="Komut95"
                    OnClick ="[Event Procedure]"
                    GroupTable =9
                    GridlineColor =10921638

                    LayoutCachedLeft =3852
                    LayoutCachedTop =1476
                    LayoutCachedWidth =4908
                    LayoutCachedHeight =1764
                    LayoutGroup =4
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    GroupTable =9
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Açılan_Kutu62_Change()
Alt12.Requery
Açılan_Kutu78.Requery
End Sub
Private Sub Açılan_Kutu78_Change()
Alt12.Requery
End Sub
Private Sub Komut90_Click()
Alt12.Requery
End Sub
Private Sub Komut95_Click()

Dim exl As Excel.Application
Dim wbook As Excel.Workbook
Dim fileName As String
Dim Path As String


If IsNull(Forms!FRMAIN!Açılan_Kutu62) Then
MsgBox ("iş seçiniz")
Exit Sub
ElseIf IsNull(Forms!FRMAIN!Açılan_Kutu78) Then
MsgBox ("hakediş seçiniz")
Exit Sub
Else
End If

'TASLAĞI AÇ
fileName = "HKDS.xlsm"
Path = CurrentProject.Path

If GENEL.AlreadyOpen(fileName, Path) Then
    Set exl = GetObject(, "Excel.Application")
    exl.Visible = True
    Set wbook = exl.Workbooks(fileName)
Else
    Set exl = CreateObject("Excel.Application")
    exl.Visible = True
    Set wbook = exl.Workbooks.Open(Path & "/" & fileName)
End If

'DATABASE AÇ
Dim dbase As DAO.Database
Set dbase = DBEngine.OpenDatabase("C:\Users\berk.kilinc\Desktop\HKDS\HKDS02\hkds-02.accdb", True)
Dim recSet As DAO.Recordset
Dim qdParams As DAO.QueryDef
Dim i As Long
Dim q As Long
Dim j As Integer

'GoTo 100 'DEBUG İÇİN

'PARA BİRİM SAYISI ÖĞRENME VE KAYDETME
Dim paraBirimCount As Integer
Dim paraBirims() As String
Set recSet = dbase.TableDefs("TBPARABIRIMS").OpenRecordset
recSet.MoveLast
paraBirimCount = recSet.RecordCount
recSet.MoveFirst
ReDim paraBirims(0 To (paraBirimCount - 1))
For i = 0 To (paraBirimCount - 1)
paraBirims(i) = recSet.Fields("PARABIRIMI").Value
recSet.MoveNext
Next i
recSet.Close

Set qdParams = dbase.QueryDefs("SRICMAL")
qdParams.Parameters("Forms.FRMAIN.Açılan_Kutu62").Value = Forms!FRMAIN!Açılan_Kutu62

'İMALAT VE MALZEME KALEMLERİ DOLURUCUSU İCMAL SAYFASINA
For j = 0 To (paraBirimCount - 1)
    qdParams.Parameters("paraBirim").Value = paraBirims(j)
    Set recSet = qdParams.OpenRecordset
    recSet.MoveLast
    q = recSet.RecordCount - 1
    recSet.MoveFirst
    For i = 0 To q
        If Not recSet.EOF Then
            If paraBirims(j) = "TRY" Then
            wbook.Worksheets("GENEL İMALAT İCMAL-TRY").Activate
            ElseIf paraBirims(j) = "EUR" Then
            wbook.Worksheets("GENEL İMALAT İCMAL-EUR").Activate
            Else
            MsgBox ("Taslağı var olamayan para birimi!!!")
            Exit For
            End If
            wbook.ActiveSheet.Range("B9").Offset(i, 0).Value = recSet.Fields("POZNO").Value 'POZNO
            wbook.ActiveSheet.Range("D9").Offset(i, 0).Value = recSet.Fields("ACIKLAMA").Value 'ACIKLAMA TO D9
            wbook.ActiveSheet.Range("G9").Offset(i, 0).Value = recSet.Fields("BIRIMFIYAT").Value 'BIRIMFIYAT TO G9
            wbook.ActiveSheet.Range("F9").Offset(i, 0).Value = recSet.Fields("BIRIMADI").Value 'BIRIMADI TO F9
            wbook.ActiveSheet.Range("H9").Offset(i, 0).Value = recSet.Fields("SOZLESMEMIKTARI").Value 'SOZLESMEMIKTARI
        Else
        Exit For
        End If
        recSet.MoveNext
    Next i
    recSet.Close
Next j

'İLGİLİ HAKEDİŞTEKİ BİNA SAYISI VE ADLARI
Dim binaCount As Integer
Dim binaNames() As String
Set qdParams = dbase.QueryDefs("SRHKDBINAS")
qdParams.Parameters("Forms.FRMAIN.Açılan_Kutu62").Value = Forms!FRMAIN!Açılan_Kutu62
qdParams.Parameters("Forms.FRMAIN.Açılan_Kutu78").Value = ""
Set recSet = qdParams.OpenRecordset
recSet.MoveLast
binaCount = recSet.RecordCount
recSet.MoveFirst
ReDim binaNames(0 To (binaCount - 1))
For i = 0 To (binaCount - 1)
    binaNames(i) = recSet.Fields("BINAADI").Value
    recSet.MoveNext
Next i
recSet.Close

'HER BİNA VE VAR OLAN HER PARA BİRİMİ İÇİN SAYFA AÇ.
'HER PARA BİRİMİ İÇİN ÇALIŞACAK ANCAK EN BAŞTAKİNİN TL OLDUĞUNU VS SIRALI Bİ ŞEKİLDE OLDUĞU VARSAYILIYOR.
Dim TLnum As Integer
For i = 0 To (paraBirimCount - 1)
    For j = 0 To (binaCount - 1)
        TLnum = wbook.Worksheets("GENEL İMALAT İCMAL-TRY").Index 'sayfa açılınca indexler değiştiği için içeride.
        wbook.Worksheets(TLnum + i).Activate
        GENELICMAL.newicmal binaNames(j), wbook
    Next j
Next i

For i = 0 To (paraBirimCount - 1)
    For j = 0 To (binaCount - 1)
        wbook.Worksheets("İMALAT İCMAL-" & binaNames(j) & "-" & paraBirims(i)).Activate
        ICMAL.addMetraj wbook
    Next j
Next i


'METRAJ YERLEŞTİRİCİSİ
'RECSET E HAKED İ TOPLA ÖNCE
Set qdParams = dbase.QueryDefs("SRHAKED")
qdParams.Parameters("Forms.FRMAIN.Açılan_Kutu62").Value = Forms!FRMAIN!Açılan_Kutu62
qdParams.Parameters("Forms.FRMAIN.Açılan_Kutu78").Value = ""
Set recSet = qdParams.OpenRecordset
Dim recCount As Long
recSet.MoveLast
recCount = recSet.RecordCount
recSet.MoveFirst

'GEREKLİ DEĞİŞKENLER
Dim binaadi As String
Dim paraBirimi As String
Dim sayfAdi As String
Dim metrajNo As Integer
Dim metrajSatir As Long
Dim pozNo As String
Dim metrajAd As String
Dim metrajSira As Long
Dim miktar As Double

For i = 0 To (recCount - 1)
    'GELECEK HAKEDİŞLER İÇİN YÖNLENDİRME.
    If recSet.Fields("HAKEDISNO").Value > Forms!FRMAIN!Açılan_Kutu78 Then
        GoTo 666 'GELECEK HAKEDİŞLER DIŞARI
    Else
    End If
    
    paraBirimi = recSet.Fields("PARABIRIMI").Value
    wbook.Worksheets("GENEL İMALAT İCMAL-" & paraBirimi).Activate
    pozNo = recSet.Fields("POZNO").Value
    metrajSatir = wbook.ActiveSheet.Range("b1").EntireColumn.Find(pozNo).Row
    metrajNo = wbook.ActiveSheet.Range("A" & metrajSatir).Value
    binaadi = recSet.Fields("BINAADI").Value
    wbook.Worksheets(metrajNo & "_" & binaadi & "_" & paraBirimi).Activate
    
    'AYNISINDAN VAR MI ?? YOK İSE START TO NEXT YAP
    'AYNISINDAN VAR MI?
    metrajAd = recSet.Fields("KATADI") & " " & recSet.Fields("BLOKADI") & " " & recSet.Fields("MAHAL/NOTLAR")
    On Error Resume Next
    metrajSira = wbook.ActiveSheet.Range("b1:c1").EntireColumn.Find(metrajAd).Row
    If Err.Number <> 0 Then
        'START TO NEXT (AYNI METRAJDAN YOK)
        wbook.ActiveSheet.Range("b13").Activate
        Do Until exl.ActiveCell.Value = ""
        exl.ActiveCell.Offset(1, 0).Select
        Loop
        exl.ActiveCell.Value = metrajAd
        metrajSira = exl.ActiveCell.Row
        Err.Clear
        On Error GoTo 0
    Else
        'AYNI METRAJDAN VAR BU KONUDA NE YAPILMASI GEREKTİĞİ ÜZERİNDE ÇALIŞMA GEREKİYOR. ***************************
        metrajSira = wbook.ActiveSheet.Range("b1:c1").EntireColumn.Find(metrajAd).Row
        wbook.ActiveSheet.Range("B" & metrajSira).Select
        On Error GoTo 0
    End If
    
    If recSet.Fields("HAKEDISNO").Value < Forms!FRMAIN!Açılan_Kutu78 Then
        'DOLDUR ESKİ METRAJ SAYFASINI!!
        miktar = recSet.Fields("MIKTAR") * recSet.Fields("BENZER") * recSet.Fields("PURS")
        exl.ActiveCell.Offset(0, 9).Value = exl.ActiveCell.Offset(0, 9).Value + miktar
        exl.ActiveCell.Offset(0, 1).Value = exl.ActiveCell.Offset(0, 1).Value + recSet.Fields("MIKTAR")
        exl.ActiveCell.Offset(0, 2).Value = recSet.Fields("BIRIMADI")
        exl.ActiveCell.Offset(0, 3).Value = recSet.Fields("PURS")
        exl.ActiveCell.Offset(0, 4).Value = recSet.Fields("BENZER")
        exl.ActiveCell.Offset(0, 5).Value = recSet.Fields("BOY")
        exl.ActiveCell.Offset(0, 6).Value = recSet.Fields("EN")
        exl.ActiveCell.Offset(0, 7).Value = recSet.Fields("YUKSEKLIK")
    ElseIf recSet.Fields("HAKEDISNO").Value > Forms!FRMAIN!Açılan_Kutu78 Then
        GoTo 666 'GELECEK HAKEDİŞLER DIŞARI
    ElseIf recSet.Fields("HAKEDISNO").Value = Forms!FRMAIN!Açılan_Kutu78 Then
        'DOLDUR METRAJ SAYFASINI!!
        exl.ActiveCell.Offset(0, 1).Value = exl.ActiveCell.Offset(0, 1).Value + recSet.Fields("MIKTAR")
        exl.ActiveCell.Offset(0, 2).Value = recSet.Fields("BIRIMADI")
        exl.ActiveCell.Offset(0, 3).Value = recSet.Fields("PURS")
        exl.ActiveCell.Offset(0, 4).Value = recSet.Fields("BENZER")
        exl.ActiveCell.Offset(0, 5).Value = recSet.Fields("BOY")
        exl.ActiveCell.Offset(0, 6).Value = recSet.Fields("EN")
        exl.ActiveCell.Offset(0, 7).Value = recSet.Fields("YUKSEKLIK")
                
    'BU HAKEDİŞ İÇİN İSTENEN FORMATLAMA VARSA YAZILABİLİR.
    Else
        MsgBox ("arıza")
        Exit Sub
    End If
    
666
    recSet.MoveNext
Next i

recSet.Close

'METRAJLARIN YERLEŞTİRİCİSİ BİTİİ!!
100 'DEBUG İÇİN!


'DATA SAYFASI DOLURUCU BAŞLADI
'RECSET E İŞ DATASINI TOPLA ÖNCE
Set qdParams = dbase.QueryDefs("SRISFIRMABEDEL")
qdParams.Parameters("Forms.FRMAIN.Açılan_Kutu62").Value = Forms!FRMAIN!Açılan_Kutu62
Set recSet = qdParams.OpenRecordset
recSet.MoveLast
recCount = recSet.RecordCount
recSet.MoveFirst

'GEREKLİ DEĞİŞKENLER
'YOK

'DATA SAYFASINA DOLDURMAYA BAŞLA
wbook.Worksheets("data").Activate
With wbook.Worksheets("data")
    .Range("c9").Value = recSet.Fields("FIRMAADI").Value
    .Range("C10").Value = recSet.Fields("ISADI").Value
    .Range("C14").Value = recSet.Fields("SOZLTURU").Value
    .Range("C15").Value = recSet.Fields("SOZLNO").Value
    .Range("C16").Value = recSet.Fields("SOZLTARIH").Value
    .Range("C17").Value = recSet.Fields("YERTESTARIH").Value
    .Range("C18").Value = recSet.Fields("ISBITIMTARIH").Value
    .Range("C19").Value = recSet.Fields("SUREUZATIMI").Value
    .Range("C22").Value = Forms!FRMAIN!Açılan_Kutu78
For i = 0 To (recCount - 1)
    If recSet.Fields("PARABIRIMI").Value = "TRY" Then 'TRY İÇİN
        .Range("C28").Value = recSet.Fields("ANASOZLKESIFBEDELI").Value
        .Range("C29").Value = recSet.Fields("ILAVESOZLBEDELI").Value
        .Range("C32").Value = recSet.Fields("KDV").Value
        .Range("C33").Value = 0 'KESİN TEMİNAT MEKTUBU MİKTARI !!! YOK !!!
        .Range("C35").Value = recSet.Fields("NAKITTEMINAT").Value
        .Range("C36").Value = recSet.Fields("DIGERTEMINAT").Value
        .Range("C37").Value = recSet.Fields("DAMGAVERGISI").Value
        .Range("C38").Value = recSet.Fields("STOPAJ").Value
        .Range("C39").Value = recSet.Fields("KDVTEVKIFATI").Value
        .Range("C40").Value = recSet.Fields("ISAVANSI").Value
        .Range("C41").Value = 0 'KESİN TEMİNAT MIKTARI !!! YOK !!!
    ElseIf recSet.Fields("PARABIRIMI").Value = "EUR" Then 'EUR İÇİN
        .Range("D28").Value = recSet.Fields("ANASOZLKESIFBEDELI").Value
        .Range("D29").Value = recSet.Fields("ILAVESOZLBEDELI").Value
        .Range("D32").Value = recSet.Fields("KDV").Value
        .Range("D33").Value = 0 'KESİN TEMİNAT MEKTUBU MİKTARI !!! YOK !!!
        .Range("D35").Value = recSet.Fields("NAKITTEMINAT").Value
        .Range("D36").Value = recSet.Fields("DIGERTEMINAT").Value
        .Range("D37").Value = recSet.Fields("DAMGAVERGISI").Value
        .Range("D38").Value = recSet.Fields("STOPAJ").Value
        .Range("D39").Value = recSet.Fields("KDVTEVKIFATI").Value
        .Range("D40").Value = recSet.Fields("ISAVANSI").Value
        .Range("D41").Value = 0 'KESİN TEMİNAT MIKTARI !!! YOK !!!
    Else
        MsgBox ("HATA")
    Exit Sub
    End If
    recSet.MoveNext
Next i


'TOPLAM ESKİ HAKEDİŞ BEDELLERİ İÇİN SORGU TASARIMI!


End With

















'DATA SAYFASI DOLDURUCU BİTTİİ!!


'HAKEDİŞ RAPORLARDAN VERİ ALIP HAKEDİŞ TAKİBE YAZICI






















End Sub
