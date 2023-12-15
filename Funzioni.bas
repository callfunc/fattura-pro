Attribute VB_Name = "Funzioni"
Public Enum StatoRecord
NonModificato = 0
modifica = 1
inserimento = 2
End Enum

Private Type GUID
 Data1    As Long
 Data2    As Integer
 Data3    As Integer
 Data4(7) As Byte
End Type

Private Type PICTDESC
 Size     As Long
 Type     As Long
 hBmp     As Long
 hPal     As Long
 Reserved As Long
End Type

Private Type GdiplusStartupInput
 GdiplusVersion           As Long
 DebugEventCallback       As Long
 SuppressBackgroundThread As Long
 SuppressExternalCodecs   As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GdipLoadImageFromStream Lib "gdiplus.dll" (ByVal Stream As IUnknown, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hdc As Long, GpGraphics As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)

Private Const PATCOPY = &HF00021
Private Const PICTYPE_BITMAP = 1

Public Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
wMsg As Long, ByVal wParam As Integer, lParam As Any) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetThreadLocale Lib "Kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "Kernel32" Alias "GetLocaleInfoA" _
(ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As _
Long) As Long
Public Declare Function SetLocaleInfo Lib "Kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, _
ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As _
Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal _
uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function SetCurrentDirectory Lib "Kernel32" Alias "SetCurrentDirectoryA" _
(ByVal lpPathName As String) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, _
ByVal fDeleteOnResume As Long, ByRef ppstr As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32.dll" (ByVal lpStream As IUnknown, _
ByVal lSize As Long, ByVal fRunMode As Long, ByRef riid As GUID, ByRef lplpObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, _
ByRef pclsid As GUID) As Long

Public ImgCancella As IPictureDisp
Public ImgConferma As IPictureDisp
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_STHOUSAND = &HF
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_GETDROPPEDSTATE = &H157
Public Const LB_SETSEL = &H185
Public Const WM_MOUSEMOVE = &H200
Public Const WM_PAINT = &HF
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_ERASEBKGND = &H14
Public Const WM_MOUSEWHEEL = &H20A
Public Const SM_CXVSCROLL = 2
Const WS_VSCROLL = &H200000
Const GWL_STYLE = (-16)

Enum TipoDocumentoStampa
 FatturaDifferita = 0
 FatturaImmediata = 1
 NotaCredito = 2
 ReportBilancio = 3
 ReportFattureClienti = 4
 ReportFattureFornitori = 5
End Enum

Public Type CoordinateMouse
 x As Single
 y As Single
End Type

Public conn As ADODB.Connection, EseguiBackup As Boolean
Public PercorsoApp$, LargSB&
Public LCID As Integer, SepMigCorr As String, SepDecCorr As String, StampaConf As Boolean

Private GDIToken As Long
Private Const GUID_IPICTURE As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
Public Function MakeLong(ByVal wLow As Integer, ByVal wHigh As Integer) As Long
MakeLong = wHigh * &H10000 + wLow
End Function
Public Sub InitGDIPlus()
Dim gdipInit As GdiplusStartupInput
    
gdipInit.GdiplusVersion = 1
GdiplusStartup GDIToken, gdipInit, ByVal 0&
End Sub
Public Sub FreeGDIPlus()
GdiplusShutdown GDIToken
End Sub
Public Function CaricaImmagineDaRisorsa(IdRisorsa As String, Optional Width As Long = -1, _
Optional Height As Long = -1, Optional ByVal BackColor As Long = vbWhite, Optional RetainRatio As _
Boolean = False) As IPicture
Dim hdc     As Long
Dim s       As IUnknown
Dim hBitmap As Long
Dim hGraphics As Long
Dim Img     As Long

Dim ArrDatiImg() As Byte: ArrDatiImg = LoadResData(IdRisorsa, "CUSTOM")
If CreateStreamOnHGlobal(ArrDatiImg(LBound(ArrDatiImg)), False, s) = 0 Then
 If GdipLoadImageFromStream(s, Img) = 0 Then

  If Width = -1 Or Height = -1 Then
   GdipGetImageWidth Img, Width
   GdipGetImageHeight Img, Height
  End If
    
  InizializzaDC hdc, hBitmap, BackColor, Width, Height
  
  GdipCreateFromHDC hdc, Graphics
  GdipDrawImageRectI Graphics, Img, 0, 0, Width, Height
  GdipDisposeImage Img
    
  hBitmap = SelectObject(hdc, hBitmap)
  DeleteDC hdc
  
  Set CaricaImmagineDaRisorsa = CreaImmagine(hBitmap)
 End If
End If
End Function
Private Sub InizializzaDC(hdc As Long, hBitmap As Long, BackColor As Long, Width As Long, Height As _
Long)
Dim hBrush As Long
hdc = CreateCompatibleDC(ByVal 0&)
hBitmap = CreateCompatibleBitmap(GetDC(0&), Width, Height)
hBitmap = SelectObject(hdc, hBitmap)
hBrush = CreateSolidBrush(BackColor)
hBrush = SelectObject(hdc, hBrush)
PatBlt hdc, 0, 0, Width, Height, PATCOPY
DeleteObject SelectObject(hdc, hBrush)
End Sub
Private Function CreaImmagine(hBitmap As Long) As IPicture
Dim IID_IDispatch As GUID
Dim pic           As PICTDESC
    
IID_IDispatch.Data1 = &H20400
IID_IDispatch.Data4(0) = &HC0
IID_IDispatch.Data4(7) = &H46
       
pic.Size = Len(pic)
pic.Type = PICTYPE_BITMAP
pic.hBmp = hBitmap

OleCreatePictureIndirect pic, IID_IDispatch, True, CreaImmagine
End Function
Public Function IdDocumento(ByVal NumDoc As String) As String
If NumDoc <> "" Then
 Dim NumDoc1$
 NumDoc1 = Split(NumDoc, "/")(0)
 IdDocumento = String(4 - Len(NumDoc1), "0") & NumDoc
End If
End Function
Public Function NumeroDocumento(ByVal IdDoc As String)
Dim NumDoc$
NumDoc = Mid(IdDoc, 6)
Do While Left(NumDoc, 1) = "0"
 NumDoc = Mid(NumDoc, 2)
Loop
NumeroDocumento = NumDoc
End Function
Public Function MouseInGriglia(Griglia As MSFlexGrid, cm As CoordinateMouse) As Boolean
MouseInGriglia = cm.y > (Griglia.RowPos(0) + Griglia.RowHeight(0)) And cm.y <= _
(Griglia.RowPos(Griglia.Rows - 1) + Griglia.RowHeight(Griglia.Rows - 1)) And _
cm.x <= (Griglia.ColPos(Griglia.Cols - 1) + Griglia.ColWidth(Griglia.Cols - 1))
End Function
Public Function ControlloCarIns(ControlloTesto As TextBox, ByVal CarIns$, ByVal NumCifreInt%, ByVal _
NumCifreDec%) As Boolean
Dim PosVirgola%: PosVirgola = InStr(ControlloTesto, ",")
If CarIns = "," And (PosVirgola Or ControlloTesto.SelStart = 0) Then
 Exit Function
End If
If IsNumeric(CarIns) Then
 Dim NumCifre%, NumCifreAmmesse%
 If PosVirgola Then
  If ControlloTesto.SelStart < PosVirgola Then
   If ControlloTesto.SelStart <> 0 And Left(ControlloTesto.Text, 1) = "0" Then
    Exit Function
   End If
   NumCifreAmmesse = NumCifreInt: NumCifre = PosVirgola
  Else
   NumCifreAmmesse = NumCifreDec: NumCifre = Len(ControlloTesto) - PosVirgola
  End If
 Else
  If ControlloTesto.SelStart <> 0 And Left(ControlloTesto.Text, 1) = "0" Then
   Exit Function
  End If
  NumCifreAmmesse = NumCifreInt: NumCifre = Len(ControlloTesto)
 End If
 If NumCifre + 1 > NumCifreAmmesse Then
  Exit Function
 End If
End If
ControlloCarIns = True
End Function
Sub Main()
If App.PrevInstance Then
 MsgBox "Attenzione, Fattura Pro è gia in esecuzione !", vbOKOnly + vbExclamation + _
 vbApplicationModal, "Fattura Pro"
 Exit Sub
Else
 PercorsoApp = App.Path
 If Right(PercorsoApp, 1) <> "\" Then PercorsoApp = PercorsoApp & "\"
 If Dir(PercorsoApp & "FatturaPro.mdb") <> "" Then
  Set conn = New ADODB.Connection
  conn.CursorLocation = adUseClient
  conn.IsolationLevel = adXactChaos
  conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PercorsoApp & "FatturaPro.mdb;" _
  & "Jet OLEDB:Database Password=omuiakon14"

  Set rsImpostazioni = New ADODB.Recordset
  rsImpostazioni.Open "Impostazioni", conn, adOpenDynamic
  rsImpostazioni.MoveFirst
  EseguiBackup = rsImpostazioni("backup")
  InitGDIPlus
  SetCurrentDirectory PercorsoApp
  Set ImgCancella = CaricaImmagineDaRisorsa("CANCELLA", , , vbWhite, False)
  Set ImgConferma = CaricaImmagineDaRisorsa("CONFERMA", , , vbWhite, False)
  LargSB = GetSystemMetrics(SM_CXVSCROLL)
  Dim Data As String, ret As Integer, DataLen As Long
  LCID = GetThreadLocale: ret = GetLocaleInfo(LCID, LOCALE_STHOUSAND, Data, DataLen)
  If ret <> 0 Then
   DataLen = ret: Data = Space(DataLen)
   ret = GetLocaleInfo(LCID, LOCALE_STHOUSAND, Data, DataLen)
   SepMigCorr = Left(Data, DataLen - 1)
   If SepMigCorr <> "." Then SetLocaleInfo LCID, LOCALE_STHOUSAND, "."
  End If
  Data = "": DataLen = 0
  ret = GetLocaleInfo(LCID, LOCALE_SDECIMAL, Data, DataLen)
  If ret <> 0 Then
   DataLen = ret: Data = Space(DataLen)
   ret = GetLocaleInfo(LCID, LOCALE_SDECIMAL, Data, DataLen)
   SepDecCorr = Left(Data, DataLen - 1)
   If SepDecCorr <> "," Then SetLocaleInfo LCID, LOCALE_SDECIMAL, ","
  End If
  Load FatturaProMDI
 Else
  MsgBox "Attenzione, il database programma non è stato trovato !" & vbNewLine & _
  "Il programma verrà chiuso.", vbExclamation, "Fattura Pro"
 End If
End If
End Sub
Public Function FormVisibile(NomeForm As String) As Boolean
Dim fv As Boolean
For Each f In Forms
 If f.Name = NomeForm And f.Visible Then
  fv = True: Exit For
 End If
Next f
FormVisibile = fv
End Function
Public Function ScrollBarVisibile(CtlHwnd As Long) As Boolean
ScrollBarVisibile = GetWindowLong(CtlHwnd, GWL_STYLE) And WS_VSCROLL
End Function
Public Sub ScrollGriglia(Griglia As MSFlexGrid, ByVal Rotazione As Long)
With Griglia
 If Rotazione > 0 Then
  .TopRow = IIf(.TopRow <> 1, .TopRow - 1, 1)
 Else
  .TopRow = IIf(.TopRow <> .Rows - 1, .TopRow + 1, .Rows - 1)
 End If
End With
End Sub
Public Function EliminaSpazi(testo As String) As String
Dim TestoMod$, Car$, LunTesto
testo = Trim(testo): LunTesto = Len(testo)

For i = 1 To LunTesto
 Car = Mid(testo, i, 1): TestoMod = TestoMod & Car
 If Car = " " Then
  i = i + 1
  Do
   Car = Mid(testo, i, 1)
   If Car <> " " Then
    TestoMod = TestoMod & Car
   Else
    i = i + 1
   End If
  Loop While Car = " "
 End If
Next i
EliminaSpazi = TestoMod
End Function
Public Function FormattaNumero(ByVal n As String) As String
Dim PosSep As Integer, LunNum As Integer, NumDecimali As Integer
PosSep = InStr(1, n, ","): LunNum = Len(n)
If PosSep <> 0 And LunNum - PosSep < 2 Then
 FormattaNumero = FormatNumber(CDbl(n), 2): Exit Function
ElseIf PosSep <> 0 Then
 NumDecimali = LunNum - PosSep
Else
 NumDecimali = 2
End If
FormattaNumero = FormatNumber(CDbl(n), NumDecimali)
End Function
Public Function Arrotonda(ByVal n As Double) As String
Dim out As String, PosSep As Integer, dec As Integer: NumMod = CStr(n)
PosSep = InStr(1, NumMod, ",")
If PosSep <> 0 And Len(NumMod) - PosSep >= 3 Then
 If Mid(NumMod, PosSep + 3, 1) = "5" And Mid(NumMod, PosSep + 2, 1) <> "9" Then
  NumMod = Mid(NumMod, 1, PosSep + 1) & CInt(Mid(NumMod, PosSep + 2, 1)) + 1
 ElseIf Mid(NumMod, PosSep + 3, 1) = "5" And Mid(NumMod, PosSep + 2, 1) = "9" Then
  dec = CInt(Mid(NumMod, PosSep + 1, 2)) + 1
  NumMod = "" & (CInt(Mid(NumMod, 1, PosSep - 1)) + (dec \ 100)) & "," _
  & IIf(dec = 100, "00", dec)
 Else
  NumMod = FormatNumber(n, 2)
 End If
Else
 NumMod = FormatNumber(n, 2)
End If
Arrotonda = NumMod
End Function

