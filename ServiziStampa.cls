VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServiziStampa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type LOGFONT
 lfHeight As Long
 lfWidth As Long
 lfEscapement As Long
 lfOrientation As Long
 lfWeight As Long
 lfItalic As Byte
 lfUnderline As Byte
 lfStrikeOut As Byte
 lfCharSet As Byte
 lfOutPrecision As Byte
 lfClipPrecision As Byte
 lfQuality As Byte
 lfPitchAndFamily As Byte
 lfFaceName As String * 32
End Type

Private Type Size
 x As Long
 y As Long
End Type

Private Type CoordDati
 NumVoce As Integer
 PrimaVoce As Boolean
End Type

Dim Logo As IPictureDisp, OffsetX As Single, OffsetY As Single, Anteprima As Boolean
Dim ColonneVoci As Variant, cv() As CoordDati, LargColonneVoci As Variant, StampaY As Integer, _
AltezzaRiga As Integer, Stampante As Object, cd As CoordDati
Dim NumNota As Integer, NumVoce As Integer, CaricaPagine As Boolean, PagCorr As Integer
Public TipoDoc As TipoDocumentoStampa, RepScadenze As ReportScadenzario, ReportBilancio As Collection, _
TotaliDoc As Collection
Public rsVociDoc As ADODB.Recordset, rsDoc As ADODB.Recordset, rsCliente As ADODB.Recordset, _
InfoDitta As Variant, ElencoModPag As Variant, LuogoConsegna As Variant
Dim InizioStampa As Single, FineStampa As Single, AreaStampabileX As Single, AreaStampabileY As Single, _
LimiteCorpo As Integer, LargMMPix As Single
Const PHYSICALOFFSETX = 112: Const PHYSICALOFFSETY = 113
Const LOGPIXELSX = 88: Const LOGPIXELSY = 90
Private lf As LOGFONT: Private PrintDC As Long
Private FontPiccolo As Long, FontCorrente As Long
Private Const StrCopAnalog = "Copia analogica (PDF) della fattura elettronica. L'origin" _
& "ale in formato XML è disponibile nell'area riservata del sito web Fatture e Corrispe" _
& "ttivi dell'Agenzia delle Entrate"
Private Const StrObbContr = "Assolve gli obblighi di cui all'articolo 62, comma 1, del " _
& "decreto legge 24 Gennaio 2012, n. 1, convertito, con modificazioni, dalla legge 24 Marzo " _
& "2012 n. 27"
Private Const StrOrgContr = "Prodotto proveniente da Agricoltura e Zootecnia Biologica, " _
& "controllata da ECOGRUPPO ITALIA S.R.L. Cod. Op. H786"
Dim LargPixMM As Single

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, _
ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" _
(lpLogFont As LOGFONT) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" _
(ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As _
Long
Public Sub OffsetStampante()
OffsetX = -25.4 * Abs(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX) / _
GetDeviceCaps(Printer.hdc, LOGPIXELSX))
OffsetY = -25.4 * Abs(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY) / _
GetDeviceCaps(Printer.hdc, LOGPIXELSY))
End Sub
Public Sub ImpostaAnteprima(AnteprimaStampa As Boolean)
Anteprima = AnteprimaStampa: Call ImpostaDimensioniPagina
If Anteprima Then
 Load AnteprimaDoc: AnteprimaDoc.Inizializza Me
 Set Stampante = AnteprimaDoc.PicPrinter
 Stampante.Width = AreaStampaX: Stampante.Height = AreaStampaY
 OffsetX = 0: OffsetY = 0
Else
 Set Stampante = Printer
End If
LargMMPix = Math.Round(Stampante.ScaleX(1, vbPixels, vbMillimeters), 2)
End Sub
Public Property Get OffsetStampaX() As Single
OffsetStampaX = OffsetX
End Property
Public Property Get OffsetStampaY() As Single
OffsetStampaY = OffsetY
End Property
Public Sub StampaDocDifferito(NumPag As Integer)
Dim IdNota$, NumVoce%, PrimaVoce As Boolean, ColonneVoci As Variant
Dim DataDoc As String, NumDoc As String
Call CreaFontPiccolo
NumVoce = 1
StampaY = 82: LimiteCorpo = 230
ColonneVoci = Array("descr", "lotto", "um", "qnt", "prezzo", "totale", "sconto", "iva")
If rsVociDoc.RecordCount Then
 rsVociDoc.MoveFirst
End If
If NumPag = -1 Then
 NumNota = 1: NumVoce = 0: PagCorr = 1
 If Anteprima Then
  ReDim cv(0) As CoordDati
  cv(0).NumVoce = 1: CaricaPagine = True
 End If
Else
 NumVoce = cv(NumPag - 1).NumVoce: PrimaVoce = cv(NumPag - 1).PrimaVoce
 rsVociDoc.Move NumVoce - 1, adBookmarkFirst: CaricaPagine = False
End If
Call DisegnaModulo(PagCorr): Stampante.Font.Size = 8: Stampante.Font.Bold = True
Stampante.CurrentY = StampaY + OffsetY
IdNota = "0"
While Not rsVociDoc.EOF
 If rsVociDoc("IdNota") <> IdNota Then
  PrimaVoce = True: IdNota = rsVociDoc("IdNota")
 Else
  PrimaVoce = False
 End If
 Dim RigheDescrizione() As String
 RigheDescrizione = DisponiSuRighe(UCase(rsVociDoc("descr")), 0)
 If StampaY + (UBound(RigheDescrizione) * 3 + 5) > 226 Then
  If Not Anteprima Then
   Stampante.NewPage: PagCorr = PagCorr + 1: Call DisegnaModulo(PagCorr)
   StampaY = 82: Stampante.Font.Size = 8: Stampante.Font.Bold = True
  ElseIf NumPag = -1 Then
   PagCorr = PagCorr + 1: ReDim Preserve cv(PagCorr - 1) As CoordDati
   cv(PagCorr - 1).NumVoce = NumVoce: cv(PagCorr - 1).PrimaVoce = PrimaVoce
   StampaY = 82
  Else: Exit Sub
  End If
 End If
 XColonne = InizioStampa
 If StampaRecord() And PrimaVoce Then
  NumDoc = NumeroDocumento(rsVociDoc("IdNota"))
  DataDoc = rsVociDoc("data")
  Stampante.CurrentY = StampaY + OffsetY
  Stampante.CurrentX = (LargColonneVoci(0) - Stampante.TextWidth(NumDoc)) / 2 + XColonne + OffsetX
  Stampante.Print NumDoc
  Stampante.CurrentX = (LargColonneVoci(1) - Stampante.TextWidth(DataDoc)) / 2 + XColonne + LargColonneVoci(0) + OffsetX
  Stampante.CurrentY = StampaY + OffsetY: Stampante.Print DataDoc
 End If
 For i = 0 To UBound(ColonneVoci)
  Dim CampoVoce$, CifreDec%
  CampoVoce = IIf(IsNull(rsVociDoc(ColonneVoci(i))), "", rsVociDoc(ColonneVoci(i)))
  If CampoVoce <> "" Then
   If IsNumeric(CStr(rsVociDoc(ColonneVoci(i)))) Then
    If ColonneVoci(i) = "qnt" Then
     CifreDec = 3
    ElseIf ColonneVoci(i) <> "iva" Then
     CifreDec = 2
    End If
    If rsVociDoc(ColonneVoci(i)) <> 0 Then
     CampoVoce = FormatNumber(rsVociDoc(ColonneVoci(i)), CifreDec)
    Else
     CampoVoce = ""
    End If
   Else
    CampoVoce = UCase(rsVociDoc(ColonneVoci(i)))
   End If
   Stampante.CurrentY = StampaY + OffsetY
   If i = 0 Then
    For r = 0 To UBound(RigheDescrizione)
     XColonne = InizioStampa + LargColonneVoci(0) + LargColonneVoci(1)
     If StampaRecord() Then
      Stampante.CurrentX = XColonne + 1.5 + OffsetX: Stampante.Print RigheDescrizione(r)
     End If
     If r <> UBound(RigheDescrizione) Then
      AltezzaRiga = AltezzaRiga + 3: Stampante.CurrentY = StampaY + AltezzaRiga + OffsetY
     End If
    Next r
   Else
    XColonne = InizioStampa
    For j = 0 To i + 1
     XColonne = XColonne + LargColonneVoci(j)
    Next j
    If StampaRecord() Then
     Stampante.CurrentX = (LargColonneVoci(i + 2) - Stampante.TextWidth(CampoVoce)) / 2 + XColonne + OffsetX
     Stampante.Print CampoVoce
    End If
   End If
  End If
 Next i
 StampaY = StampaY + AltezzaRiga + 5: AltezzaRiga = 0
 rsVociDoc.MoveNext: NumVoce = NumVoce + 1
Wend
If Anteprima Then
 If CaricaPagine Then
  AnteprimaDoc.Pagine = PagCorr: AnteprimaDoc.PagineDoc.Caption = "di " & PagCorr
 End If
Else
 Call StampaTotali: Stampante.EndDoc
End If
Exit Sub
CancellaDocumentoStampante:
If Stampante Is Printer Then
 Printer.KillDoc
 MsgBox "Attenzione, si è verificato un errore durante la stampa !", vbExclamation, _
 "Fattura Pro"
Else
 MsgBox "Attenzione, si è verificato un errore durante la creazione dell'anteprima di stampa !", _
 vbExclamation, "Fattura Pro"
End If
Exit Sub
End Sub
Private Function CalcolaLarghezzaReport() As Single
PosCampi = Array(15.5, 15.5, 68, 23, 8, 16, 15, 15, 14, 7)
For i = 0 To UBound(PosCampi)
 CalcolaLarghezzaReport = CalcolaLarghezzaReport + PosCampi(i)
Next
End Function
Public Sub StampaDoc(ByVal NumPag As Integer)
Dim ColonneVoci As Variant
Call CreaFontPiccolo
StampaY = 82: LimiteCorpo = 230
ColonneVoci = Array("descr", "lotto", "um", "qnt", "prezzo", "totale", "sconto", "iva")
Dim cd As CoordDati, completato As Boolean
If rsVociDoc.RecordCount Then
 rsVociDoc.MoveFirst
End If
If NumPag = -1 Then
 NumVoce = 0: NumDdt = 0: PagCorr = 1
 If Anteprima Then
  ReDim cv(0) As CoordDati: completato = False
  cv(0).NumVoce = NumVoce
  CaricaPagine = True
 End If
Else
 PagCorr = NumPag: NumPag = NumPag - 1: CaricaPagine = False
 If NumPag <= UBound(cv) Then
  NumVoce = cv(NumPag).NumVoce
  rsVociDoc.Move NumVoce, adBookmarkFirst
 End If
End If
Call DisegnaModulo(PagCorr)
Stampante.Font.Size = 8: Stampante.Font.Bold = True
Call StampaVoci(ColonneVoci)

If Anteprima Then
 If CaricaPagine Then
  AnteprimaDoc.Pagine = UBound(cv) + 1
  AnteprimaDoc.PagineDoc.Caption = "di " & AnteprimaDoc.Pagine
 End If
Else
 Call StampaTotali: Stampante.EndDoc
End If
Exit Sub
CancellaDocumentoStampante:
If Stampante Is Printer Then
 Stampante.KillDoc
 MsgBox "Attenzione, si è verificato un errore durante la stampa !", vbExclamation, _
 "Fattura Pro"
Else
 MsgBox "Attenzione, si è verificato un errore durante la creazione dell'anteprima di stampa !", _
 vbExclamation, "Fattura Pro"
End If
Exit Sub
End Sub
Private Sub StampaVoci(ByVal ColonneVoci As Variant)
Dim RigheDescrizione() As String
While Not rsVociDoc.EOF
 RigheDescrizione = DisponiSuRighe(UCase(rsVociDoc(ColonneVoci(0))), 0)
 If StampaY + (UBound(RigheDescrizione) * 3 + 5) > 225 Then
  If Not Anteprima Then
   PagCorr = PagCorr + 1: Stampante.NewPage: Call DisegnaModulo(PagCorr)
   StampaY = 82: Stampante.Font.Size = 8: Stampante.Font.Bold = True
  ElseIf CaricaPagine Then
   ReDim Preserve cv(PagCorr) As CoordDati: cv(PagCorr).NumVoce = rsVociDoc.AbsolutePosition
   StampaY = 82: PagCorr = PagCorr + 1
  Else: Exit Sub
  End If
 End If
 XColonne = InizioStampa
 If StampaRecord() And rsVociDoc.AbsolutePosition = 1 Then
  If TipoDoc <> NotaCredito Then
   If rsDoc("numddt") <> "" Then
    Stampante.CurrentY = StampaY + OffsetY
    NumDdt = rsDoc("numddt")
    Stampante.CurrentY = StampaY + OffsetY
    DataDdt = Format$(rsDoc("dataddt"), "dd-mm-yyyy")
    Stampante.CurrentX = (LargColonneVoci(0) - Stampante.TextWidth(NumDdt)) / 2 + XColonne + OffsetX
    Stampante.Print NumDdt
    Stampante.CurrentX = (LargColonneVoci(1) - Stampante.TextWidth(DataDdt)) / 2 + XColonne + LargColonneVoci(0) + OffsetX
    Stampante.CurrentY = StampaY + OffsetY: Stampante.Print DataDdt
   End If
  Else
   Dim StrDocRif$, AnnoRifDoc$
   AnnoRifDoc = Split(rsDoc("RifDoc"), "-")(1)
   StrDocRif = "Riferimento ns. fattura n. " & Split(rsDoc("RifDoc"), "-")(0) & " del " & AnnoRifDoc
   Stampante.CurrentY = StampaY + OffsetY
   Stampante.CurrentX = XColonne + LargColonneVoci(0) + LargColonneVoci(1) + 1.5 + OffsetX
   Stampante.Print StrDocRif
   StampaY = StampaY + 5
  End If
 End If
 Dim AttrVoce$, CifreDec%
 For i = 0 To UBound(ColonneVoci)
  AttrVoce = "": CifreDec = 0
  If Not IsNull(rsVociDoc(ColonneVoci(i))) Then
   If IsNumeric(CStr(rsVociDoc(ColonneVoci(i)))) Then
    If i = 3 Then
     CifreDec = 3
    ElseIf ColonneVoci(i) <> "iva" Then
     CifreDec = 2
    End If
    If rsVociDoc(ColonneVoci(i)) <> 0 Then
     AttrVoce = FormatNumber(rsVociDoc(ColonneVoci(i)), CifreDec)
    End If
   Else
    AttrVoce = UCase(rsVociDoc(ColonneVoci(i)))
   End If
  End If
  If AttrVoce <> "" Then
   Stampante.CurrentY = StampaY + OffsetY
   If i = 0 Then
    For r = 0 To UBound(RigheDescrizione)
     XColonne = InizioStampa + LargColonneVoci(0) + LargColonneVoci(1)
     If StampaRecord() Then
      Stampante.CurrentX = XColonne + 1.5 + OffsetX: Stampante.Print RigheDescrizione(r)
     End If
     If r <> UBound(RigheDescrizione) Then
      AltezzaRiga = AltezzaRiga + 3: Stampante.CurrentY = StampaY + AltezzaRiga + OffsetY
     End If
    Next r
   Else
    XColonne = InizioStampa
    For j = 0 To i + 1
     XColonne = XColonne + LargColonneVoci(j)
    Next j
    If StampaRecord() Then
     Stampante.CurrentX = (LargColonneVoci(i + 2) - Stampante.TextWidth(AttrVoce)) / 2 + XColonne + OffsetX
     Stampante.Print AttrVoce
    End If
   End If
  End If
 Next i
 NumVoce = NumVoce + 1: StampaY = StampaY + AltezzaRiga + 5: AltezzaRiga = 0
 rsVociDoc.MoveNext
Wend
End Sub
Private Function CaricaLogo(ByVal ImgLogo As Variant) As String
Dim stm As New ADODB.Stream
stm.Type = adTypeBinary
stm.Open
stm.Write ImgLogo
If stm.Size > 0 Then
 FileLogo = App.Path & "\logo." & InfoDitta(12, 0)
 stm.SaveToFile FileLogo, adSaveCreateOverWrite
 CaricaLogo = FileLogo
End If
End Function
Private Sub StampaIntestazioneDitta()
Dim NomeAzienda$, TipoAzienda$, Ditta$, FileLogo$, OffsetLogo As Single, LargLogo, AltLogo
Dim LargIntestazione As Integer: LargIntestazione = 100 - 3.8
Stampante.Font.Size = 10: Stampante.Font.Bold = True
NomeAzienda = InfoDitta(0, 0)
Ditta = InfoDitta(1, 0)
FileLogo = CaricaLogo(InfoDitta(11, 0))
If FileLogo <> "" Then
 Set Logo = LoadPicture(FileLogo)
 Kill FileLogo
Else: Set Logo = LoadResPicture(101, vbResBitmap)
End If
Stampante.CurrentX = Abs(3.8 + OffsetX)
LargLogo = Logo.Width / 100: AltLogo = Logo.Height / 100

OffsetLogo = (LargIntestazione - LargLogo) / 2 + (Abs(3.8 + OffsetX))
Stampante.PaintPicture Logo, OffsetLogo, 6 + OffsetY, LargLogo, AltLogo
StampaMultiLinea NomeAzienda, LargIntestazione, 3.8, 6 + AltLogo + 2.7, True
Stampante.Font.Bold = False
If Ditta <> NomeAzienda Then
 Stampante.CurrentY = Stampante.CurrentY + 0.6
 Ditta = "di " & Ditta
 Stampante.CurrentX = (LargIntestazione - Stampante.TextWidth(Ditta)) / 2 + Abs(3.8 + _
 OffsetX)
 Stampante.Print Ditta
End If
Stampante.Font.Size = 8: Stampante.CurrentY = 39.8 + OffsetY
Stampante.CurrentX = Abs(3.8 + OffsetX)
Stampante.Print "Sede Legale: " & InfoDitta(4, 0)
Stampante.CurrentY = Stampante.CurrentY + 0.4
Stampante.CurrentX = Abs(3.8 + OffsetX)
Stampante.Print "Sede Aziendale: " & InfoDitta(5, 0)
Stampante.CurrentY = Stampante.CurrentY + 0.4
Stampante.CurrentX = Abs(3.8 + OffsetX)
Dim Tel$, Fax$
Tel = InfoDitta(8, 0)
Fax = InfoDitta(9, 0)
If Tel = Fax Or Fax = "" Then
 Stampante.Print "Tel/Fax: " & Tel
Else
 Stampante.Print "Tel: " & Tel & " - Fax: " & Fax
End If
Stampante.CurrentY = Stampante.CurrentY + 0.4
Stampante.CurrentX = Abs(3.8 + OffsetX)
Dim Piva$, CodFisc$
Piva = InfoDitta(2, 0)
Stampante.Print "P. Iva: " & Piva
CodFisc = InfoDitta(3, 0)
If CodFisc <> "" Then
 Stampante.CurrentX = Abs(3.8 + OffsetX)
 Stampante.CurrentY = Stampante.CurrentY + 0.4
 Stampante.Print "Cod. Fiscale: " & CodFisc
End If
End Sub
Private Sub StampaMultiLinea(ByVal testo As String, LarghezzaCol As Integer, ByVal StampaX As Single, _
ByVal StampaY As Single, Optional ByVal Centrato As Boolean = False)
Dim riga As String, LunghezzaTesto As Integer, LunghezzaRiga As Integer
Dim NumRighe As Integer, FineParola As Integer, PosCorrente As Integer
LunghezzaTesto = Len(testo)
For i = 1 To LunghezzaTesto
 LunghezzaRiga = i - PosCorrente
 riga = Mid(testo, PosCorrente + 1, LunghezzaRiga)
 If Stampante.TextWidth(riga) > LarghezzaCol Then
  If FineParola Then LunghezzaRiga = FineParola - PosCorrente - 1
  If Centrato Then
   Stampante.CurrentX = (LarghezzaCol - Stampante.TextWidth(riga)) / 2 + Abs(StampaX + OffsetX)
  Else
   Stampante.CurrentX = Abs(StampaX + OffsetX)
  End If
  Stampante.CurrentY = StampaY + OffsetY
  Stampante.Print riga
  NumRighe = NumRighe + 1
  If NumRighe > 2 Then Exit Sub
  StampaY = StampaY + 4: PosCorrente = PosCorrente + LunghezzaRiga
 End If
 If Mid(testo, i, 1) = " " Then FineParola = i
Next i

LunghezzaRiga = LunghezzaTesto - PosCorrente
If Centrato Then
 Stampante.CurrentX = (LarghezzaCol - Stampante.TextWidth(Mid(testo, PosCorrente + 1, LunghezzaRiga))) / 2 _
 + Abs(StampaX + OffsetX)
Else
 Stampante.CurrentX = Abs(StampaX + OffsetX)
End If
Stampante.CurrentY = StampaY + OffsetY
Stampante.Print Mid(testo, PosCorrente + 1, LunghezzaRiga)
End Sub
Private Sub DisegnaModulo(ByVal NumPag As Integer)
Dim NumDoc$, ModPag$, CoordBan$, InfoAggiuntiva%, i%, cy, CodFisc
Dim NomiColonneVoci As Variant, CampiInfo As Variant, LargCampiInfo As Variant, XColonne As Single, ValCampiInfo As Variant
Call StampaIntestazioneDitta
NumDoc = NumeroDocumento(rsDoc("IdDoc"))
CampiInfo = Array("N. Pagina", "Condizioni di Pagamento", "Coordinate Bancarie", "N. Documento", _
"Tipo Documento", "Data")
'3.8,22,64,132,158,188,FineStampa
LargCampiInfo = Array(18.2, 42, 56, 26, 30, FineStampa - (172.2 + InizioStampa))
XColonne = InizioStampa
If TipoDoc <> NotaCredito Then
 ModPag = IIf(Not rsDoc.EOF, ElencoModPag(rsDoc("modpag")), "")
End If
CoordBan = IIf(IsNull(InfoDitta(7, 0)), "", InfoDitta(7, 0))
ValCampiInfo = Array(NumPag, ModPag, CoordBan, NumDoc, "", Format$(rsDoc("data"), "dd-mm-yyyy"))
NomiColonneVoci = Array("N. Ddt", "Data Ddt", "Descrizione", "Lotto", "UM", "Quantità", "Prezzo", _
"Importo", "Sconto", "Iva")

Stampante.Font.Size = 9: Stampante.Font.Bold = False
If InStr(rsCliente("Ditta"), "IL TOCCO S.R.L.") <> 0 Then
 InfoAggiuntiva = 1
End If
cy = 11
If InfoAggiuntiva = 1 Then
 cy = 16
End If
Stampante.Line (100 + OffsetX, cy + OffsetY)-(FineStampa + OffsetX, cy + OffsetY)
Stampante.Line (FineStampa + OffsetX, cy + OffsetY)-(FineStampa + OffsetX, 60 + OffsetY)
Stampante.Line (100 + OffsetX, cy + OffsetY)-(100 + OffsetX, 60 + OffsetY)
Stampante.Line (100 + OffsetX, 50 + OffsetY)-(FineStampa + OffsetX, 50 + OffsetY)
Stampante.Line (InizioStampa + OffsetX, 60 + OffsetY)-(FineStampa + OffsetX, 73 + OffsetY), RGB(0, 0, 0), B
Stampante.CurrentX = 102 + OffsetX: Stampante.CurrentY = cy + 2 + OffsetY
Stampante.Print "Spett.": Stampante.Font.Size = 8: Stampante.Font.Bold = True
cy = cy + 2.25

If Not rsCliente.EOF Then
 Dim Loc$
 StampaMultiLinea UCase(rsCliente("Ditta")), 83, 114, cy
 Stampante.CurrentX = 114 + OffsetX: cy = cy + 10: Stampante.CurrentY = cy + OffsetY
 Stampante.Print UCase(rsCliente("Indirizzo"))

 Stampante.CurrentX = 114 + OffsetX: cy = cy + 4: Stampante.CurrentY = cy + OffsetY
 If rsCliente("Cap") <> "" Then
  Loc = rsCliente("Cap") & " " & UCase(rsCliente("Loc"))
 Else
  Loc = UCase(rsCliente("Loc"))
 End If
 Stampante.Print Loc
 Stampante.CurrentX = 180 + OffsetX: Stampante.CurrentY = cy + OffsetY
 If rsCliente("Prov") <> "" Then
  Stampante.Print UCase(rsCliente("Prov"))
 Else
  Stampante.Print UCase(rsCliente("Stato"))
 End If
End If

If InfoAggiuntiva = 1 Then
 cy = cy + 7: Stampante.CurrentX = 102 + OffsetX: Stampante.CurrentY = cy + OffsetY
 Stampante.Print "Sede Operativa : VIA S.CATERINA 42/B ACIREALE CT"
Else
 cy = cy + 6
 Stampante.Line (100 + OffsetX, cy + OffsetY)-(FineStampa + OffsetX, cy + OffsetY)
 cy = cy + 1
 Stampante.CurrentX = 102 + OffsetX: Stampante.CurrentY = cy + OffsetY
 Stampante.Font.Size = 9: Stampante.Font.Bold = False: Stampante.Print "Luogo di Consegna"
 cy = cy + 5: Stampante.Font.Size = 8: Stampante.Font.Bold = True
 If IsArray(LuogoConsegna) Then
  Stampante.CurrentX = 108 + OffsetX: Stampante.CurrentY = cy + OffsetY
  Stampante.Print UCase(LuogoConsegna(2, 0)): cy = cy + 4
  Stampante.CurrentX = 108 + OffsetX: Stampante.CurrentY = cy + OffsetY
  Stampante.Print UCase(LuogoConsegna(3, 0)) & "     " & UCase(LuogoConsegna(4, 0)) _
  & " " & LuogoConsegna(5, 0) & " (" & LuogoConsegna(6, 0) & ")"
 End If
End If
Stampante.CurrentX = 102 + OffsetX: Stampante.CurrentY = 51 + OffsetY
Stampante.Font.Size = 9: Stampante.Font.Bold = False
Stampante.Print "P. IVA": Stampante.Font.Size = 8: Stampante.Font.Bold = True
Stampante.CurrentX = 111 + OffsetX: Stampante.CurrentY = 55 + OffsetY
Stampante.Print rsCliente("PartitaIVA"): Stampante.Font.Size = 9: Stampante.Font.Bold = False
Stampante.Line (142 + OffsetX, 50 + OffsetY)-(142 + OffsetX, 60 + OffsetY)
Stampante.CurrentX = 144 + OffsetX: Stampante.CurrentY = 51 + OffsetY
Stampante.Print "Cod. Fiscale": Stampante.Font.Size = 8: Stampante.Font.Bold = True
Stampante.CurrentX = 150 + OffsetX: Stampante.CurrentY = 55 + OffsetY
CodFisc = rsCliente("CodFiscale")
If CodFisc <> "" Then Stampante.Print CodFisc
Stampante.Font.Size = 9: Stampante.Font.Bold = False
ValCampiInfo(4) = IIf(TipoDoc <> NotaCredito, "Fattura", "Nota di Credito")
For i = 0 To UBound(CampiInfo)
 If i <> UBound(CampiInfo) Then
  Stampante.Line (XColonne + LargCampiInfo(i) + OffsetX, 60 + OffsetY)-(XColonne + LargCampiInfo(i) + OffsetX, 73 + _
  OffsetY)
 End If
 Stampante.Font.Size = 9: Stampante.Font.Bold = False: Stampante.CurrentY = 61 + OffsetY
 Select Case i
  Case 0, 3 To 5
  Stampante.CurrentX = (LargCampiInfo(i) - Stampante.TextWidth(CampiInfo(i))) / 2 + XColonne + OffsetX
  Stampante.Print CampiInfo(i): Stampante.CurrentY = 67.5 + OffsetY
  Stampante.Font.Size = 8: Stampante.Font.Bold = True
  Stampante.CurrentX = (LargCampiInfo(i) - Stampante.TextWidth(ValCampiInfo(i))) / 2 + XColonne + OffsetX
  Case 1, 2
  Stampante.CurrentX = XColonne + 2 + OffsetX: Stampante.Print CampiInfo(i)
  Stampante.CurrentY = 67.5 + OffsetY
  Stampante.Font.Size = 8: Stampante.Font.Bold = True
  If i = 1 Then
   Stampante.CurrentX = XColonne + LargCampiInfo(i) - 2 - Stampante.TextWidth(ValCampiInfo(i)) + OffsetX
  Else
   Stampante.CurrentX = XColonne + 2 + OffsetX
  End If
 End Select
 Stampante.Print ValCampiInfo(i)
 XColonne = XColonne + LargCampiInfo(i)
Next i
Stampante.Font.Size = 9: Stampante.Font.Bold = False
Stampante.Line (InizioStampa + OffsetX, 75 + OffsetY)-(FineStampa + OffsetX, 236 + OffsetY), RGB(0, 0, 0), B
Stampante.Line (InizioStampa + OffsetX, 81 + OffsetY)-(FineStampa + OffsetX, 81 + OffsetY)
XColonne = InizioStampa
For i = 0 To UBound(NomiColonneVoci)
 If i <> UBound(NomiColonneVoci) Then
  Stampante.Line (XColonne + LargColonneVoci(i) + OffsetX, 75 + OffsetY)-(XColonne + LargColonneVoci(i) + OffsetX, _
  81 + OffsetY)
 End If
 Stampante.CurrentY = 76 + OffsetY
 Select Case i
  Case 0, 1, 3 To 9
  Stampante.CurrentX = (LargColonneVoci(i) - Stampante.TextWidth(NomiColonneVoci(i))) / 2 + XColonne + OffsetX
  Case 2
  Stampante.CurrentX = XColonne + 1.5 + OffsetX
 End Select
 Stampante.Print NomiColonneVoci(i)
 XColonne = XColonne + LargColonneVoci(i)
Next i
Call StampaEsenzioneIva: Stampante.Font.Size = 9: Stampante.Font.Bold = False
Stampante.Line (InizioStampa + OffsetX, 238 + OffsetY)-(88 + OffsetX, 270 + OffsetY), RGB(0, 0, 0), B
Stampante.Line (InizioStampa + OffsetX, 244 + OffsetY)-(88 + OffsetX, 244 + OffsetY)
Stampante.Line (26 + OffsetX, 238 + OffsetY)-(26 + OffsetX, 270 + OffsetY)
Stampante.CurrentX = 9.1 + OffsetX: Stampante.CurrentY = 239 + OffsetY: Stampante.Print "Aliquota"
Stampante.Line (46 + OffsetX, 238 + OffsetY)-(46 + OffsetX, 270 + OffsetY)
Stampante.CurrentX = 31.5 + OffsetX: Stampante.CurrentY = 239 + OffsetY: Stampante.Print "Totale"
Stampante.Line (68 + OffsetX, 238 + OffsetY)-(68 + OffsetX, 270 + OffsetY)
Stampante.CurrentX = 49.3 + OffsetX: Stampante.CurrentY = 239 + OffsetY: Stampante.Print "Imponibile"
Stampante.CurrentX = 72 + OffsetX: Stampante.CurrentY = 239 + OffsetY: Stampante.Print "Imposta"
Stampante.Line (142 + OffsetX, 238 + OffsetY)-(FineStampa + OffsetX, 245 + OffsetY), RGB(0, 0, 0), B
Stampante.CurrentX = 144 + OffsetX: Stampante.CurrentY = 239 + OffsetY: Stampante.Print "Totale Imponibile"
Stampante.Line (142 + OffsetX, 248 + OffsetY)-(FineStampa + OffsetX, 255 + OffsetY), RGB(0, 0, 0), B
Stampante.CurrentX = 144 + OffsetX: Stampante.CurrentY = 249 + OffsetY: Stampante.Print "Totale Iva"
Stampante.Line (142 + OffsetX, 258 + OffsetY)-(FineStampa + OffsetX, 265 + OffsetY), RGB(0, 0, 0), B
Stampante.CurrentX = 144 + OffsetX: Stampante.CurrentY = 259 + OffsetY: Stampante.Font.Bold = True
Stampante.Print "Totale Documento"
Dim r As Long, StrSize As Size: FontCorrente = SelectObject(PrintDC, FontPiccolo)
GetTextExtentPoint32 PrintDC, StrCopAnalog, Len(StrCopAnalog), StrSize
r = TextOut(PrintDC, ConvMillInPix((AreaStampaX - ConvPixInMill(StrSize.x, LOGPIXELSX)) / 2 + _
OffsetX, LOGPIXELSX), ConvMillInPix(276 + OffsetY, LOGPIXELSY), StrCopAnalog, Len(StrCopAnalog))
GetTextExtentPoint32 PrintDC, StrObbContr, Len(StrObbContr), StrSize
r = TextOut(PrintDC, ConvMillInPix((AreaStampaX - ConvPixInMill(StrSize.x, LOGPIXELSX)) / 2 + _
OffsetX, LOGPIXELSX), ConvMillInPix(282 + OffsetY, LOGPIXELSY), StrObbContr, Len(StrObbContr))
GetTextExtentPoint32 PrintDC, StrOrgContr, Len(StrOrgContr), StrSize
r = TextOut(PrintDC, ConvMillInPix((AreaStampaX - ConvPixInMill(StrSize.x, LOGPIXELSX)) / 2 + _
OffsetX, LOGPIXELSX), ConvMillInPix(287 + OffsetY, LOGPIXELSY), StrOrgContr, Len(StrOrgContr))
r = SelectObject(PrintDC, FontCorrente)
If Anteprima Then Call StampaTotali
End Sub
Private Sub StampaTotali()
Dim EsenteIva$, TotDoc$, TotImp$, TotIva$
TotImp = FormatNumber(rsDoc("TotImp"), 2)
TotIva = FormatNumber(rsDoc("TotIva"), 2)
TotDoc = FormatNumber(rsDoc("TotDoc"), 2)
If Not IsNull(rsDoc("esenteiva")) Then EsenteIva = rsDoc("esenteiva")
Stampante.Font.Size = 8: Stampante.CurrentX = FineStampa - 2 - Stampante.TextWidth(TotImp) + OffsetX
Stampante.CurrentY = 239 + OffsetY: Stampante.Print TotImp
Stampante.CurrentX = FineStampa - 2 - Stampante.TextWidth(TotIva) + OffsetX
Stampante.CurrentY = 249 + OffsetY: Stampante.Print TotIva
Stampante.CurrentX = FineStampa - 2 - Stampante.TextWidth(TotDoc) + OffsetX
Stampante.CurrentY = 259 + OffsetY: Stampante.Print TotDoc: PosY = 245
If EsenteIva = "" Then
 Dim Totale$, Imponibile$, Iva$
 For Each TotaleIva In TotaliDoc
  If TotaleIva.Imponibile <> 0 Then
   Totale = FormatNumber(TotaleIva.Totale, 2)
   Imponibile = FormatNumber(TotaleIva.Imponibile, 2)
   Iva = FormatNumber(TotaleIva.Iva, 2)
   Stampante.CurrentX = 24 - Stampante.TextWidth(TotaleIva.Aliquota) + OffsetX
   Stampante.CurrentY = PosY + OffsetY: Stampante.Print TotaleIva.Aliquota
   Stampante.CurrentX = 44 - Stampante.TextWidth(Totale) + OffsetX
   Stampante.CurrentY = PosY + OffsetY: Stampante.Print Totale
   Stampante.CurrentX = 66 - Stampante.TextWidth(Imponibile) + OffsetX
   Stampante.CurrentY = PosY + OffsetY: Stampante.Print Imponibile
   Stampante.CurrentX = 86 - Stampante.TextWidth(Iva) + OffsetX
   Stampante.CurrentY = PosY + OffsetY: Stampante.Print Iva: PosY = PosY + 5
  End If
 Next
End If
End Sub
Public Sub AbilitaControlliAnteprima(abilita As Boolean)
AnteprimaDoc.PagPrec.Enabled = abilita: AnteprimaDoc.PagSucc.Enabled = abilita
AnteprimaDoc.Stampa.Enabled = abilita
End Sub
Private Sub DisegnaReportFatture(rsDitta As ADODB.Recordset, ByVal LarghezzaReport As Integer, _
ByVal LarghezzeColonne As Variant, ByVal Tipo As Byte, ByVal NumPag As Integer)
Call StampaIntestazioneDitta
Dim NomeDitta$, AltezzaInt%, StampaX As Single, FineReport As Single
Stampante.CurrentX = 112 + OffsetY: Stampante.CurrentY = 25 + OffsetX
Stampante.Font.Size = 9: Stampante.Font.Bold = False
Stampante.Print "Spett.": Stampante.Line (109 + OffsetX, 21 + OffsetY)-(201.8 + OffsetX, 21 + OffsetY)
Stampante.Line (201.8 + OffsetX, 21 + OffsetY)-(201.8 + OffsetX, 60 + OffsetY)
Stampante.Line (109 + OffsetX, 21 + OffsetY)-(109 + OffsetX, 60 + OffsetY)
Stampante.Line (109 + OffsetX, 50 + OffsetY)-(201.8 + OffsetX, 50 + OffsetY)
Stampante.Line (109 + OffsetX, 60 + OffsetY)-(201.8 + OffsetX, 60 + OffsetY)
Stampante.Font.Size = 8: Stampante.Font.Bold = True
NomeDitta = UCase(rsDitta("ditta"))
Stampante.CurrentX = 126 + OffsetX: Stampante.CurrentY = 25.25 + OffsetY
If Stampante.TextWidth(NomeDitta) > 72 Then
 Dim riga As String, EndPos, LastEndPos: EndPos = 0
 Do
  LastEndPos = EndPos: EndPos = InStr(EndPos + 1, NomeDitta, " ")
  If EndPos = 0 Then
   Exit Do
  ElseIf Stampante.TextWidth(Left(NomeDitta, EndPos - 1)) <= 72 Then
   riga = Left(NomeDitta, EndPos - 1)
  End If
 Loop While Stampante.TextWidth(Left(NomeDitta, EndPos - 1)) <= 72
 Stampante.Print UCase(riga): Stampante.CurrentX = 126 + OffsetX
 Stampante.CurrentY = 31 + OffsetY: Stampante.Print Mid(NomeDitta, LastEndPos + 1)
Else
 Stampante.Print UCase(NomeDitta)
End If
Stampante.CurrentX = 126 + OffsetX: Stampante.CurrentY = 37 + OffsetY
Stampante.Print UCase(rsDitta("indirizzo"))
Stampante.CurrentX = 126 + OffsetX: Stampante.CurrentY = 43 + OffsetY
Cap = rsDitta("cap")
If Cap <> "" Then
 Stampante.Print Cap & " " & UCase(rsDitta("loc"))
Else
 Stampante.Print UCase(rsDitta("loc"))
End If
Dim Prov$, Stato$
Prov = rsDitta("prov")
Stato = UCase(rsDitta("stato"))
If Prov <> "" Then
 Stampante.CurrentX = 201.8 - Stampante.TextWidth(Prov) - 4 + OffsetX
 Stampante.CurrentY = 43 + OffsetY
 Stampante.Print Prov
Else
 Stampante.CurrentX = 201.8 - Stampante.TextWidth(Stato) - 4 + OffsetX
 Stampante.CurrentY = 43 + OffsetY
 Stampante.Print Stato
End If
Stampante.CurrentX = 126 + OffsetX: Stampante.CurrentY = 52 + OffsetY
Stampante.Print "P. IVA  " & rsDitta("partitaiva")
Stampante.Font.Size = 9
Dim NomiColonne As Variant
If Tipo = 0 Then
 Titolo = "REPORT FATTURE CLIENTE"
 NomiColonne = Array("N. Documento", "Data", "Importo", "Scadenza Pagamento", "Modalità di Pagamento", _
 "Stato Pagamento")
Else
 Titolo = "REPORT FATTURE FORNITORE"
 NomiColonne = Array("N. Documento", "Data", "Importo", "Scadenza Pagamento", "Modalità di Pagamento", "Stato Pagamento")
End If
StampaX = (AreaStampaX - LarghezzaReport) / 2: FineReport = StampaX + LarghezzaReport
Stampante.Line (StampaX + OffsetX, 75 + OffsetY)-(FineReport + OffsetX, 238 + OffsetY), _
RGB(0, 0, 0), B
Stampante.CurrentX = (LarghezzaReport - Stampante.TextWidth(Titolo)) / 2 + StampaX + OffsetX
Stampante.CurrentY = (8.1 - Stampante.TextHeight("A")) / 2 + 75 + OffsetY
Stampante.Print Titolo: Stampante.Font.Bold = False
AltezzaInt = (7 - Stampante.TextHeight("abc")) / 2 + 83.1 + OffsetY
Stampante.Line (StampaX + OffsetX, 83.1 + OffsetY)-(FineReport + OffsetX, 83.1 + OffsetY)
Stampante.Line (StampaX + OffsetX, 90.1 + OffsetY)-(FineReport + OffsetX, 90.1 + OffsetY)
For i = 0 To UBound(NomiColonne)
 Stampante.CurrentX = (LarghezzeColonne(i) - Stampante.TextWidth(NomiColonne(i))) / 2 + StampaX + OffsetX
 Stampante.CurrentY = AltezzaInt
 Stampante.Print NomiColonne(i): StampaX = StampaX + LarghezzeColonne(i)
 If i <> UBound(NomiColonne) Then
  Stampante.Line (StampaX + OffsetX, 83.1 + OffsetY)-(StampaX + OffsetX, 90.1 + OffsetY)
 End If
Next
Stampante.CurrentX = (AreaStampaX - Stampante.TextWidth("Pagina " & NumPag)) / 2 + OffsetX
Stampante.CurrentY = AreaStampaY - Stampante.TextHeight("a") - 2 + OffsetY
Stampante.Print "Pagina " & NumPag
End Sub
Public Sub StampaReportFatture(Tipo As Byte, NumPag As Integer)
Dim ElencoFatture As Collection, StampaX As Single, InizioReport As Single
Dim LarghezzeColonne As Variant, ColVoci As Variant, LarghezzaReport As Single
StampaY = 92.1
If NumPag = -1 Then
 ReDim cv(0) As CoordDati: PagCorr = 1
 NumVoce = PagCorr: cv(0).NumVoce = NumVoce: CaricaPagine = True
Else
 NumVoce = cv(NumPag - 1).NumVoce: CaricaPagine = False
End If
LarghezzeColonne = Array(26, 20, 20, 37, 44, 31)
For i = 0 To UBound(LarghezzeColonne)
 LarghezzaReport = LarghezzaReport + LarghezzeColonne(i)
Next
Set ElencoFatture = RepScadenze.ElencoFatture
Call DisegnaReportFatture(RepScadenze.Ditta, LarghezzaReport, LarghezzeColonne, Tipo, PagCorr)
StampaX = (AreaStampaX - LarghezzaReport) / 2
InizioReport = StampaX: FineReport = StampaX + LarghezzaReport
Do While NumVoce <= ElencoFatture.Count
 If StampaY + Stampante.TextHeight("abc") > 220 Then
  StampaY = 92.1
  If Not Anteprima Then
   Stampante.NewPage: PagCorr = PagCorr + 1: Call DisegnaReportFatture(RepScadenze.Ditta, LarghezzaReport, _
   LarghezzeColonne, Tipo, PagCorr)
  ElseIf NumPag = -1 Then
   cd.NumVoce = NumVoce: PagCorr = PagCorr + 1: ReDim Preserve cv(PagCorr - 1) As CoordDati
   cv(PagCorr - 1) = cd
  Else
   Exit Do
  End If
 End If
 Stampante.Font.Size = 8: Stampante.Font.Bold = True
 If StampaRecord() Then
  StampaX = InizioReport
  For j = 0 To UBound(ElencoFatture(NumVoce))
   Stampante.CurrentX = (LarghezzeColonne(j) - Stampante.TextWidth(ElencoFatture(NumVoce)(j))) / 2 + _
   StampaX + OffsetX
   Stampante.CurrentY = StampaY + OffsetY: StampaX = StampaX + LarghezzeColonne(j)
   Stampante.Print ElencoFatture(NumVoce)(j)
  Next
 End If
 StampaY = StampaY + Stampante.TextHeight("a") + 1: NumVoce = NumVoce + 1
Loop
Stampante.Line (InizioReport + OffsetX, 222 + OffsetY)-(FineReport + OffsetX, 222 + OffsetY)
StampaX = InizioReport + 2
Stampante.CurrentX = StampaX + OffsetX: Stampante.CurrentY = 224 + OffsetY
Stampante.Print "Totale Pagato"
Stampante.CurrentX = FineReport - 2 - Stampante.TextWidth(RepScadenze.TotalePagato) + OffsetX
Stampante.CurrentY = 224 + OffsetY: Stampante.Print RepScadenze.TotalePagato
Stampante.Line (InizioReport + OffsetX, 230 + OffsetY)-(FineReport + OffsetX, 230 + OffsetY)
Stampante.CurrentX = StampaX + OffsetX: Stampante.CurrentY = 232 + OffsetY
Stampante.Print "Totale Insoluto"
Stampante.CurrentX = FineReport - 2 - Stampante.TextWidth(RepScadenze.TotaleInsoluto) + OffsetX
Stampante.CurrentY = 232 + OffsetY
Stampante.ForeColor = IIf(RepScadenze.TotaleInsoluto <> 0, RGB(200, 0, 0), vbBlack)
Stampante.Print RepScadenze.TotaleInsoluto
If Anteprima Then
 If CaricaPagine Then
  AnteprimaDoc.PagineDoc.Caption = "di " & PagCorr: AnteprimaDoc.Pagine = PagCorr
 End If
Else: Stampante.EndDoc
End If
Exit Sub
CancellaDocumentoStampante:
If Stampante Is Printer Then
 Printer.KillDoc
 MsgBox "Attenzione, si è verificato un errore durante la stampa !", vbExclamation, _
 "Fattura Pro"
Else
 MsgBox "Attenzione, si è verificato un errore durante la creazione dell'anteprima di stampa !", _
 vbExclamation, "Fattura Pro"
End If
Exit Sub
End Sub
Private Sub StampaReportBilancio()
Dim NomiColonne As Variant, LarghezzzeColonne As Variant, LarghezzaReport As Single, Titolo$
Dim StampaX, StampaY
StampaIntestazioneDitta
Titolo = "REPORT BILANCIO"
NomiColonne = Array("Mese", "Imponibile Vendite", "Imponibile Acquisti", "IVA Vendite", "IVA Acquisti", _
"Bilancio Imponibile", "Bilancio IVA")
Stampante.Font.Size = 9: Stampante.Font.Bold = False
LarghezzeColonne = Array(Stampante.TextWidth("Dicembre") + 10, Stampante.TextWidth("Imponibile Vendite") + 10, _
Stampante.TextWidth("Imponibile Acquisti") + 6, Stampante.TextWidth("IVA Vendite") + 6, _
Stampante.TextWidth("IVA Acquisti") + 6, Stampante.TextWidth("Bilancio Imponibile") + 6, _
Stampante.TextWidth("Bilancio IVA") + 6)
For i = 0 To UBound(LarghezzeColonne)
 LarghezzaReport = LarghezzaReport + LarghezzeColonne(i)
Next
StampaX = (AreaStampaX - LarghezzaReport) / 2: FineReport = StampaX + LarghezzaReport
Stampante.Line (StampaX + OffsetX, 75 + OffsetY)-(FineReport + OffsetX, 238 + OffsetY), _
RGB(0, 0, 0), B
Stampante.Font.Bold = True
Stampante.CurrentX = (LarghezzaReport - Stampante.TextWidth(Titolo)) / 2 + StampaX + OffsetX
Stampante.CurrentY = (8.1 - Stampante.TextHeight("A")) / 2 + 75 + OffsetY
Stampante.Print Titolo
Stampante.Font.Bold = False
AltezzaInt = (7 - Stampante.TextHeight("abc")) / 2 + 83.1 + OffsetY
Stampante.Line (StampaX + OffsetX, 83.1 + OffsetY)-(FineReport + OffsetX, 83.1 + OffsetY)
Stampante.Line (StampaX + OffsetX, 90.1 + OffsetY)-(FineReport + OffsetX, 90.1 + OffsetY)
For i = 0 To UBound(NomiColonne)
 Stampante.CurrentX = (LarghezzeColonne(i) - Stampante.TextWidth(NomiColonne(i))) / 2 + StampaX + OffsetX
 Stampante.CurrentY = AltezzaInt
 Stampante.Print NomiColonne(i): StampaX = StampaX + LarghezzeColonne(i)
 If i <> UBound(NomiColonne) Then
  Stampante.Line (StampaX + OffsetX, 83.1 + OffsetY)-(StampaX + OffsetX, 90.1 + OffsetY)
 End If
Next
StampaX = (AreaStampaX - LarghezzaReport) / 2
Stampante.Line (StampaX + OffsetX, 224 + OffsetY)-(FineReport + OffsetX, 224 + OffsetY)
Stampante.Line (StampaX + OffsetX, 231 + OffsetY)-(FineReport + OffsetX, 231 + OffsetY)
Dim ArrMesi As Variant, VoceBilancio As Variant
ArrMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", _
"Ottobre", "Novembre", "Dicembre")
StampaY = 92.1
Stampante.Font.Bold = True
On Error Resume Next
For i = 0 To 12
 StampaX = (AreaStampaX - LarghezzaReport) / 2
 VoceBilancio = ReportBilancio(ArrMesi(i))
 If Err.Number = 0 Then
  Stampante.CurrentX = (LarghezzeColonne(0) - Stampante.TextWidth(ArrMesi(i))) / 2 + StampaX + OffsetX
  Stampante.CurrentY = StampaY
  StampaX = StampaX + LarghezzeColonne(0)
  Stampante.Print ArrMesi(i)
  For j = 0 To UBound(VoceBilancio)
   Stampante.CurrentX = (LarghezzeColonne(j + 1) - Stampante.TextWidth(VoceBilancio(j))) / 2 + _
   StampaX + OffsetX
   Stampante.CurrentY = StampaY + OffsetY: StampaX = StampaX + LarghezzeColonne(j + 1)
   Stampante.Print VoceBilancio(j)
  Next j
  StampaY = StampaY + Stampante.TextHeight("A") + 1
 Else
  Exit For
 End If
Next i
StampaX = (AreaStampaX - LarghezzaReport) / 2 + 2
Stampante.CurrentX = StampaX
Stampante.CurrentY = (7 - Stampante.TextHeight("A")) / 2 + 224 + OffsetY
Stampante.Print "Bilancio Imponibile"
Stampante.CurrentX = FineReport - 2 - Stampante.TextWidth(ReportBilancio("Totali")(0)) + OffsetX
Stampante.CurrentY = (7 - Stampante.TextHeight("A")) / 2 + 224 + OffsetY
Stampante.Print ReportBilancio("Totali")(0)
Stampante.CurrentX = StampaX
Stampante.CurrentY = (7 - Stampante.TextHeight("A")) / 2 + 231 + OffsetY
Stampante.Print "Bilancio IVA"
Stampante.CurrentX = FineReport - 2 - Stampante.TextWidth(ReportBilancio("Totali")(1)) + OffsetX
Stampante.CurrentY = (7 - Stampante.TextHeight("A")) / 2 + 231 + OffsetY
Stampante.Print ReportBilancio("Totali")(1)
End Sub
Public Function StampaRecord() As Boolean
StampaRecord = (Not Anteprima) Or (Not CaricaPagine) Or PagCorr = 1
End Function
Private Sub StampaEsenzioneIva()
If rsDoc("esenteiva") <> "" Then
 Stampante.Line (Abs(3.8 + OffsetX), 230 + OffsetY)-(FineStampa + OffsetX, 230 + OffsetY)
 Stampante.Font.Size = 8: Stampante.Font.Bold = True
 Stampante.CurrentX = 4.8 + OffsetX: Stampante.CurrentY = 231 + OffsetY
 Stampante.Print rsDoc("esenteiva"): LimiteCorpo = 225
End If
End Sub
Private Function DisponiSuRighe(ValColonna As String, NumCol As Integer) As String()
Dim LarghezzaCol As Integer, LunghezzaTesto As Integer, LunghezzaRiga As Integer
Dim NumRighe As Integer, FineParola As Integer, PosCorrente As Integer
Dim RigheTesto() As String
LarghezzaCol = LargColonneVoci(NumCol + 2)
LunghezzaTesto = Len(ValColonna)
For i = 1 To LunghezzaTesto
 LunghezzaRiga = i - PosCorrente
 If Stampante.TextWidth(Mid(ValColonna, PosCorrente + 1, LunghezzaRiga)) > LarghezzaCol Then
  If FineParola Then LunghezzaRiga = FineParola - PosCorrente - 1
  ReDim Preserve RigheTesto(NumRighe)
  RigheTesto(NumRighe) = Mid(ValColonna, PosCorrente + 1, LunghezzaRiga)
  NumRighe = NumRighe + 1: PosCorrente = PosCorrente + LunghezzaRiga
 End If
 If Mid(ValColonna, i, 1) = " " Then FineParola = i
Next i

LunghezzaRiga = LunghezzaTesto - PosCorrente: ReDim Preserve RigheTesto(NumRighe)
RigheTesto(NumRighe) = Trim(Mid(ValColonna, PosCorrente + 1, LunghezzaRiga))

DisponiSuRighe = RigheTesto
End Function
Public Property Get AreaStampaX() As Single
AreaStampaX = AreaStampabileX
End Property
Public Property Get AreaStampaY() As Single
AreaStampaY = AreaStampabileY
End Property
Public Sub ImpostaDimensioniPagina()
With Printer
 .ScaleMode = vbMillimeters: .Orientation = vbPRORPortrait
 .PaperSize = vbPRPSA4
End With
Call OffsetStampante
AreaStampabileX = Printer.ScaleWidth + (OffsetX * -1 * 2)
AreaStampabileY = Printer.ScaleHeight + (OffsetY * -1 * 2)
End Sub
Public Sub ImpostaStampante(ByVal Nome As String)
For Each ObjStampante In Printers
 If ObjStampante.DeviceName = Nome Then
  Set Printer = ObjStampante
  Printer.ScaleMode = vbMillimeters: Printer.Orientation = vbPRORPortrait
  Printer.PaperSize = vbPRPSA4: Exit For
 End If
Next
End Sub
Private Function ConvMillInPix(Val As Long, TipoRis As Long) As Long
Dim DPI As Long
DPI = GetDeviceCaps(PrintDC, TipoRis)
ConvMillInPix = (Val * DPI) / 25.4
End Function
Private Function ConvPixInMill(ValPixel As Long, TipoRis As Long) As Long
Dim DPI As Long
DPI = GetDeviceCaps(PrintDC, TipoRis)
ConvPixInMill = (25.4 * ValPixel) / DPI
End Function
Private Sub Class_Initialize()
ElencoModPag = Array("Rimessa Diretta", "30 Giorni", "60 Giorni")
End Sub
Private Sub Class_Terminate()
Set InfoDitta = Nothing: Set Logo = Nothing
End Sub
Private Sub CreaFontPiccolo()
PrintDC = Stampante.hdc
lf.lfEscapement = 0: lf.lfHeight = (7 * -20) / Stampante.ScaleY(1, vbPixels, vbTwips)
FontPiccolo = CreateFontIndirect(lf)
End Sub
Public Sub Stampa(Optional ByVal NumPag% = -1)
If Not Anteprima Then
 AbilitaControlliAnteprima False
End If
Stampante.ForeColor = RGB(0, 0, 0)
If TipoDoc <> ReportFattureClienti And TipoDoc <> ReportFattureFornitori Then
 InizioStampa = IIf(OffsetX <> 0 And (3.8 + OffsetX < 0), Abs(OffsetX), 3.8)
 FineStampa = IIf(OffsetX <> 0 And (3.8 + OffsetX < 0), AreaStampabileX + OffsetX - 0.1, AreaStampabileX - 3.8)
End If
LargColonneVoci = Array(15.6, 15.6, 65, 18, 10, 16, 15, 20, 14, FineStampa - (189.2 + InizioStampa))
Stampante.Font.Name = "Arial"
Select Case TipoDoc
Case 0:
 StampaDocDifferito NumPag
Case 1, 2:
 StampaDoc NumPag
Case 3:
 StampaReportBilancio
Case 4:
 StampaReportFatture 0, NumPag
Case 5:
 StampaReportFatture 1, NumPag
End Select
AbilitaControlliAnteprima True
End Sub
Public Property Set ReportScadenze(ByVal RepScad As ReportScadenzario)
Set RepScadenze = RepScad
End Property
