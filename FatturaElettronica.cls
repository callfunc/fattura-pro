VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FatturaElettronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim SigleStati As Collection
Public Function CreaFattura(rsDitta As ADODB.Recordset, rsDoc As ADODB.Recordset, _
rsVociDoc As ADODB.Recordset, TotaliDoc As Collection) As String
Dim DocXml As New DOMDocument, CodTramsittente As String, PercorsoFile As String
Dim rsInfoDitta As ADODB.Recordset
Dim VersioneXml As IXMLDOMProcessingInstruction
Dim ElemXml As IXMLDOMElement
Dim NuovaRiga As IXMLDOMNode
Dim RadiceDoc As IXMLDOMElement
Dim DocNS As IXMLDOMAttribute
Dim DocSchema As IXMLDOMAttribute
Dim Intestazione As IXMLDOMElement
Dim DatiTrasmissione As IXMLDOMElement
Dim IdTrasmittente As IXMLDOMElement
Dim CedentePrestatore As IXMLDOMElement
Dim CessionarioCommittente As IXMLDOMElement
Dim DatiAnagrafici As IXMLDOMElement
Dim Anagrafica As IXMLDOMElement
Dim Sede As IXMLDOMElement
Dim Corpo As IXMLDOMElement
Dim DatiGenerali As IXMLDOMElement
Dim DatiGeneraliDocumento As IXMLDOMElement
Dim DatiOrdineAcquisto As IXMLDOMElement
Dim IdFiscaleIva As IXMLDOMElement
Dim DatiTrasporto As IXMLDOMElement
Dim DatiAnagraficiVettore As IXMLDOMElement
Dim DatiBeniServizi As IXMLDOMElement
Dim DettaglioLinee As IXMLDOMElement
Dim DatiPagamento As IXMLDOMElement
Dim DettaglioPagamento As IXMLDOMElement

Dim CodTrasmittente As String
Set rsInfoDitta = New ADODB.Recordset
rsInfoDitta.Open "SELECT * FROM InfoDitta", conn, adOpenStatic, adLockReadOnly
If Not rsInfoDitta.EOF Then
 rsInfoDitta.MoveFirst
 If Not IsNull(rsInfoDitta("CodTrasmittente")) Then
  CodTrasmittente = rsInfoDitta("CodTrasmittente")
 Else
  CodTrasmittente = "04744470875"
 End If
Else
 CodTrasmittente = "04744470875"
End If

DocXml.resolveExternals = True

Set VersioneXml = DocXml.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
DocXml.appendChild VersioneXml
Set VersioneXml = Nothing

Set RadiceDoc = DocXml.createElement("p:FatturaElettronica")
DocXml.appendChild RadiceDoc

'Set DocNS = DocXml.createAttribute("xmlns:ds")
'DocNS.nodeValue = "http://www.w3.org/2000/09/xmldsig#"
'RadiceDoc.setAttributeNode DocNS

Set DocNS = DocXml.createAttribute("xmlns:p")
DocNS.nodeValue = "http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2"
RadiceDoc.setAttributeNode DocNS

Set DocNS = DocXml.createAttribute("xmlns:xsi")
DocNS.nodeValue = "http://www.w3.org/2001/XMLSchema-instance"
RadiceDoc.setAttributeNode DocNS

Set DocSchema = DocXml.createAttribute("xsi:schemaLocation")
DocSchema.nodeValue = "http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2 " _
& "http://www.fatturapa.gov.it/export/fatturazione/sdi/fatturapa/v1.2/Schema_del_file_xml_FatturaPA_versione_1.2.xsd"
RadiceDoc.setAttributeNode DocSchema

Set NuovaRiga = DocXml.createTextNode("NewLine1")
NuovaRiga.Text = vbCrLf
RadiceDoc.appendChild NuovaRiga

Set Intestazione = DocXml.createElement("FatturaElettronicaHeader")
RadiceDoc.appendChild Intestazione

Set NuovaRiga = DocXml.createTextNode("NewLine2")
NuovaRiga.Text = vbCrLf
Intestazione.appendChild NuovaRiga

Set DatiTrasmissione = DocXml.createElement("DatiTrasmissione")
Intestazione.appendChild DatiTrasmissione

Set NuovaRiga = DocXml.createTextNode("NewLine3")
NuovaRiga.Text = vbCrLf
DatiTrasmissione.appendChild NuovaRiga

Set IdTrasmittente = DocXml.createElement("IdTrasmittente")
DatiTrasmissione.appendChild IdTrasmittente

Set NuovaRiga = DocXml.createTextNode("NewLine4")
NuovaRiga.Text = vbCrLf
IdTrasmittente.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("IdPaese")
ElemXml.Text = "IT"
IdTrasmittente.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine5")
NuovaRiga.Text = vbCrLf
IdTrasmittente.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("IdCodice")
ElemXml.Text = CodTrasmittente
IdTrasmittente.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine6")
NuovaRiga.Text = vbCrLf
IdTrasmittente.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine7")
NuovaRiga.Text = vbCrLf
DatiTrasmissione.appendChild NuovaRiga

Dim IdProgressivoDoc As String

IdProgressivoDoc = IdProgressivo()
Set ElemXml = DocXml.createElement("ProgressivoInvio")
ElemXml.Text = IdProgressivoDoc
DatiTrasmissione.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine8")
NuovaRiga.Text = vbCrLf
DatiTrasmissione.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("FormatoTrasmissione")
ElemXml.Text = "FPR12"
DatiTrasmissione.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine9")
NuovaRiga.Text = vbCrLf
DatiTrasmissione.appendChild NuovaRiga

Dim CodDest As String, PECDest As String
If IsNull(rsDitta("CodDest")) Then
 CodDest = "0000000"
Else
 CodDest = rsDitta("CodDest")
End If

Set ElemXml = DocXml.createElement("CodiceDestinatario")
ElemXml.Text = CodDest
DatiTrasmissione.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine10")
NuovaRiga.Text = vbCrLf
DatiTrasmissione.appendChild NuovaRiga

If CodDest = "0000000" Then
 Set ElemXml = DocXml.createElement("PECDestinatario")
 If Not IsNull(rsDitta("pec")) Then
  ElemXml.Text = rsDitta("pec")
 Else
  ElemXml.Text = ""
 End If
 DatiTrasmissione.appendChild ElemXml
 
 Set NuovaRiga = DocXml.createTextNode("NewLine11")
 NuovaRiga.Text = vbCrLf
 DatiTrasmissione.appendChild NuovaRiga
End If

Set NuovaRiga = DocXml.createTextNode("NewLine12")
NuovaRiga.Text = vbCrLf
Intestazione.appendChild NuovaRiga

Set CedentePrestatore = DocXml.createElement("CedentePrestatore")
Intestazione.appendChild CedentePrestatore

Set NuovaRiga = DocXml.createTextNode("NewLine13")
NuovaRiga.Text = vbCrLf
CedentePrestatore.appendChild NuovaRiga

Set DatiAnagrafici = DocXml.createElement("DatiAnagrafici")
CedentePrestatore.appendChild DatiAnagrafici

Set NuovaRiga = DocXml.createTextNode("NewLine14")
NuovaRiga.Text = vbCrLf
DatiAnagrafici.appendChild NuovaRiga

Set IdFiscaleIva = DocXml.createElement("IdFiscaleIva")
DatiAnagrafici.appendChild IdFiscaleIva

Set NuovaRiga = DocXml.createTextNode("NewLine15")
NuovaRiga.Text = vbCrLf
IdFiscaleIva.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("IdPaese")
ElemXml.Text = "IT"
IdFiscaleIva.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine16")
NuovaRiga.Text = vbCrLf
IdFiscaleIva.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("IdCodice")
ElemXml.Text = rsInfoDitta("PartitaIva")
IdFiscaleIva.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine17")
NuovaRiga.Text = vbCrLf
IdFiscaleIva.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine18")
NuovaRiga.Text = vbCrLf
DatiAnagrafici.appendChild NuovaRiga

Set Anagrafica = DocXml.createElement("Anagrafica")
DatiAnagrafici.appendChild Anagrafica

Set NuovaRiga = DocXml.createTextNode("NewLine19")
NuovaRiga.Text = vbCrLf
Anagrafica.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("Denominazione")
ElemXml.Text = rsInfoDitta("Ditta")
Anagrafica.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine20")
NuovaRiga.Text = vbCrLf
Anagrafica.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine21")
NuovaRiga.Text = vbCrLf
DatiAnagrafici.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("RegimeFiscale")
ElemXml.Text = "RF01"
DatiAnagrafici.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine22")
NuovaRiga.Text = vbCrLf
DatiAnagrafici.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine23")
NuovaRiga.Text = vbCrLf
CedentePrestatore.appendChild NuovaRiga

Set Sede = DocXml.createElement("Sede")
CedentePrestatore.appendChild Sede

Set NuovaRiga = DocXml.createTextNode("NewLine24")
NuovaRiga.Text = vbCrLf
Sede.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("Indirizzo")
ElemXml.Text = rsInfoDitta("SedeLegale")
Sede.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine25")
NuovaRiga.Text = vbCrLf
Sede.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("CAP")
ElemXml.Text = "94010"
Sede.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine26")
NuovaRiga.Text = vbCrLf
Sede.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("Comune")
ElemXml.Text = "Acireale"
Sede.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine27")
NuovaRiga.Text = vbCrLf
Sede.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("Nazione")
ElemXml.Text = "IT"
Sede.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine28")
NuovaRiga.Text = vbCrLf
Sede.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine29")
NuovaRiga.Text = vbCrLf
CedentePrestatore.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine30")
NuovaRiga.Text = vbCrLf
Intestazione.appendChild NuovaRiga

Set CessionarioCommittente = DocXml.createElement("CessionarioCommittente")
Intestazione.appendChild CessionarioCommittente

Set NuovaRiga = DocXml.createTextNode("NewLine31")
NuovaRiga.Text = vbCrLf
CessionarioCommittente.appendChild NuovaRiga

Set DatiAnagrafici = DocXml.createElement("DatiAnagrafici")
CessionarioCommittente.appendChild DatiAnagrafici

Set NuovaRiga = DocXml.createTextNode("NewLine32")
NuovaRiga.Text = vbCrLf
DatiAnagrafici.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("CodiceFiscale")
ElemXml.Text = rsDitta("PartitaIva")
DatiAnagrafici.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine33")
NuovaRiga.Text = vbCrLf
DatiAnagrafici.appendChild NuovaRiga

Set Anagrafica = DocXml.createElement("Anagrafica")
DatiAnagrafici.appendChild Anagrafica

Set NuovaRiga = DocXml.createTextNode("NewLine34")
NuovaRiga.Text = vbCrLf
Anagrafica.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("Denominazione")
ElemXml.Text = rsDitta("Ditta")
Anagrafica.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine35")
NuovaRiga.Text = vbCrLf
Anagrafica.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine36")
NuovaRiga.Text = vbCrLf
DatiAnagrafici.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine37")
NuovaRiga.Text = vbCrLf
CessionarioCommittente.appendChild NuovaRiga

Set Sede = DocXml.createElement("Sede")
CessionarioCommittente.appendChild Sede

Set NuovaRiga = DocXml.createTextNode("NewLine38")
NuovaRiga.Text = vbCrLf
Sede.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("Indirizzo")
ElemXml.Text = rsDitta("Indirizzo")
Sede.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine39")
NuovaRiga.Text = vbCrLf
Sede.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("CAP")
ElemXml.Text = rsDitta("Cap")
Sede.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine40")
NuovaRiga.Text = vbCrLf
Sede.appendChild NuovaRiga

Dim ComuneDest As String
If Not IsNull(rsDitta("Comune")) Then
 Set ElemXml = DocXml.createElement("Comune")
 ElemXml.Text = rsDitta("Comune")
 Sede.appendChild ElemXml
 
 Set NuovaRiga = DocXml.createTextNode("NewLine41")
 NuovaRiga.Text = vbCrLf
 Sede.appendChild NuovaRiga
End If

Set ElemXml = DocXml.createElement("Provincia")
ElemXml.Text = rsDitta("Prov")
Sede.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine42")
NuovaRiga.Text = vbCrLf
Sede.appendChild NuovaRiga

Dim NazCommittente As String
NazCommittente = rsDitta("Stato")

Set ElemXml = DocXml.createElement("Nazione")
If NazCommittente = "" Then
 ElemXml.Text = "IT"
Else
 On Error Resume Next
 ElemXml.Text = SigleStati(NazCommittente)
End If
Sede.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine43")
NuovaRiga.Text = vbCrLf
Sede.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine44")
NuovaRiga.Text = vbCrLf
CessionarioCommittente.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine45")
NuovaRiga.Text = vbCrLf
Intestazione.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine46")
NuovaRiga.Text = vbCrLf
RadiceDoc.appendChild NuovaRiga

Set Corpo = DocXml.createElement("FatturaElettronicaBody")
RadiceDoc.appendChild Corpo

Set NuovaRiga = DocXml.createTextNode("NewLine47")
NuovaRiga.Text = vbCrLf
Corpo.appendChild NuovaRiga

Set DatiGenerali = DocXml.createElement("DatiGenerali")
Corpo.appendChild DatiGenerali

Set NuovaRiga = DocXml.createTextNode("NewLine48")
NuovaRiga.Text = vbCrLf
DatiGenerali.appendChild NuovaRiga

Set DatiGeneraliDocumento = DocXml.createElement("DatiGeneraliDocumento")
DatiGenerali.appendChild DatiGeneraliDocumento

Set NuovaRiga = DocXml.createTextNode("NewLine49")
NuovaRiga.Text = vbCrLf
DatiGeneraliDocumento.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("TipoDocumento")
ElemXml.Text = "TD01"
DatiGeneraliDocumento.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine50")
NuovaRiga.Text = vbCrLf
DatiGeneraliDocumento.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("Divisa")
ElemXml.Text = "EUR"
DatiGeneraliDocumento.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine51")
NuovaRiga.Text = vbCrLf
DatiGeneraliDocumento.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("Data")
ElemXml.Text = rsDoc("Data")
DatiGeneraliDocumento.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine52")
NuovaRiga.Text = vbCrLf
DatiGeneraliDocumento.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("Numero")
ElemXml.Text = NumeroDocumento(rsDoc("IdDoc"))
DatiGeneraliDocumento.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine53")
NuovaRiga.Text = vbCrLf
DatiGeneraliDocumento.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("ImportoTotaleDocumento")
ElemXml.Text = Replace(FormatNumber(rsDoc("TotDoc"), 2), ",", ".")
DatiGeneraliDocumento.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine54")
NuovaRiga.Text = vbCrLf
DatiGeneraliDocumento.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine55")
NuovaRiga.Text = vbCrLf
DatiGenerali.appendChild NuovaRiga

Set DatiTrasporto = DocXml.createElement("DatiTrasporto")
DatiGenerali.appendChild DatiTrasporto

Set NuovaRiga = DocXml.createTextNode("NewLine56")
NuovaRiga.Text = vbCrLf
DatiTrasporto.appendChild NuovaRiga

Set DatiAnagraficiVettore = DocXml.createElement("DatiAnagraficiVettore")
DatiTrasporto.appendChild DatiAnagraficiVettore

Set NuovaRiga = DocXml.createTextNode("NewLine57")
NuovaRiga.Text = vbCrLf
DatiAnagraficiVettore.appendChild NuovaRiga

Dim CodFiscVettore As String, NazVettor As String, NomeVettore As String
CodFiscVettore = rsInfoDitta("PartitaIva")
NomeVettore = rsInfoDitta("Ditta")
NazVettore = "IT"

If Not IsNull(rsDoc("IdVettore")) Then
 Dim rsVettore As ADODB.Recordset
 rsVettore.Open "SELECT * FROM Vettori WHERE Id = " & rsDoc("IdVettore"), conn, adOpenStatic, _
 adLockReadOnly
 If Not rsVettore.EOF Then
  rsVettore.MoveFirst
  CodFiscVettore = rsVettore("PartitaIva")
  NomeVettore = rsVettore("Ditta")
  NazVettore = rsVettore("Stato")
  If NazVettore = "" Then
   NazVettore = "IT"
  Else
   On Error Resume Next
   NazVettore = SigleStati(NazVettore)
  End If
 End If
 rsVettore.Close
End If

Set IdFiscaleIva = DocXml.createElement("IdFiscaleIva")
DatiAnagraficiVettore.appendChild IdFiscaleIva

Set NuovaRiga = DocXml.createTextNode("NewLine58")
NuovaRiga.Text = vbCrLf
IdFiscaleIva.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("IdPaese")
ElemXml.Text = NazVettore
IdFiscaleIva.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine59")
NuovaRiga.Text = vbCrLf
IdFiscaleIva.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("IdCodice")
ElemXml.Text = CodFiscVettore
IdFiscaleIva.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine60")
NuovaRiga.Text = vbCrLf
IdFiscaleIva.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine61")
NuovaRiga.Text = vbCrLf
DatiAnagraficiVettore.appendChild NuovaRiga

Set Anagrafica = DocXml.createElement("Anagrafica")
DatiAnagraficiVettore.appendChild Anagrafica

Set NuovaRiga = DocXml.createTextNode("NewLine62")
NuovaRiga.Text = vbCrLf
Anagrafica.appendChild NuovaRiga

Set ElemXml = DocXml.createElement("Denominazione")
ElemXml.Text = NomeVettore
Anagrafica.appendChild ElemXml

Set NuovaRiga = DocXml.createTextNode("NewLine63")
NuovaRiga.Text = vbCrLf
Anagrafica.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine64")
NuovaRiga.Text = vbCrLf
DatiAnagraficiVettore.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine65")
NuovaRiga.Text = vbCrLf
DatiTrasporto.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine66")
NuovaRiga.Text = vbCrLf
DatiGenerali.appendChild NuovaRiga

Set NuovaRiga = DocXml.createTextNode("NewLine67")
NuovaRiga.Text = vbCrLf
Corpo.appendChild NuovaRiga

Set DatiBeniServizi = DocXml.createElement("DatiBeniServizi")
Corpo.appendChild DatiBeniServizi

Set NuovaRiga = DocXml.createTextNode("NewLine68")
NuovaRiga.Text = vbCrLf
DatiBeniServizi.appendChild NuovaRiga

Dim NumeroLinea As Integer, NumeroRighe As Integer

NumeroRighe = 69
rsVociDoc.MoveFirst

While Not rsVociDoc.EOF
 Set DettaglioLinee = DocXml.createElement("DettaglioLinee")
 DatiBeniServizi.appendChild DettaglioLinee
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DettaglioLinee.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1
 
 Set ElemXml = DocXml.createElement("NumeroLinea")
 NumeroLinea = NumeroLinea + 1
 ElemXml.Text = NumeroLinea
 DettaglioLinee.appendChild ElemXml
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DettaglioLinee.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1
 
 Set ElemXml = DocXml.createElement("Descrizione")
 ElemXml.Text = rsVociDoc("Descr")
 DettaglioLinee.appendChild ElemXml
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DettaglioLinee.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1
 
 If rsVociDoc("UM") <> "" Then
  Set ElemXml = DocXml.createElement("UnitaMisura")
  ElemXml.Text = rsVociDoc("UM")
  DettaglioLinee.appendChild ElemXml
  
  Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
  NuovaRiga.Text = vbCrLf
  DettaglioLinee.appendChild NuovaRiga
 
  NumeroRighe = NumeroRighe + 1
 End If

 Set ElemXml = DocXml.createElement("Quantita")
 ElemXml.Text = Replace(FormatNumber(rsVociDoc("Qnt"), 2), ",", ".")
 DettaglioLinee.appendChild ElemXml
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DettaglioLinee.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1
 
 Set ElemXml = DocXml.createElement("PrezzoUnitario")
 ElemXml.Text = Replace(FormatNumber(rsVociDoc("Prezzo"), 2), ",", ".")
 DettaglioLinee.appendChild ElemXml
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DettaglioLinee.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1
 
 Set ElemXml = DocXml.createElement("PrezzoTotale")
 ElemXml.Text = Replace(FormatNumber(rsVociDoc("Totale"), 2), ",", ".")
 DettaglioLinee.appendChild ElemXml
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DettaglioLinee.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1
 
 Set ElemXml = DocXml.createElement("AliquotaIva")
 ElemXml.Text = Replace(FormatNumber(CDbl(rsVociDoc("Iva")), 2), ",", ".")
 DettaglioLinee.appendChild ElemXml
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DettaglioLinee.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DatiBeniServizi.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1
 
 rsVociDoc.MoveNext
Wend

For Each TotaleDoc In TotaliDoc
 Set DatiRiepilogo = DocXml.createElement("DatiRiepilogo")
 DatiBeniServizi.appendChild DatiRiepilogo

 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DatiRiepilogo.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1
 
 Set ElemXml = DocXml.createElement("AliquotaIva")
 ElemXml.Text = Replace(FormatNumber(CDbl(TotaleDoc.Aliquota), 2), ",", ".")
 DatiRiepilogo.appendChild ElemXml
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DatiRiepilogo.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1

 Set ElemXml = DocXml.createElement("ImponibileImporto")
 ElemXml.Text = Replace(FormatNumber(TotaleDoc.Imponibile, 2), ",", ".")
 DatiRiepilogo.appendChild ElemXml
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DatiRiepilogo.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1

 Set ElemXml = DocXml.createElement("Imposta")
 ElemXml.Text = Replace(FormatNumber(TotaleDoc.Iva, 2), ",", ".")
 DatiRiepilogo.appendChild ElemXml
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DatiRiepilogo.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1

 Set ElemXml = DocXml.createElement("EsigibilitaIva")
 ElemXml.Text = "D"
 DatiRiepilogo.appendChild ElemXml
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DatiRiepilogo.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1
 
 Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
 NuovaRiga.Text = vbCrLf
 DatiBeniServizi.appendChild NuovaRiga
 
 NumeroRighe = NumeroRighe + 1
Next

Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
NuovaRiga.Text = vbCrLf
Corpo.appendChild NuovaRiga
 
NumeroRighe = NumeroRighe + 1

Set NuovaRiga = DocXml.createTextNode("NewLine" & NumeroRighe)
NuovaRiga.Text = vbCrLf
RadiceDoc.appendChild NuovaRiga

Set fso = New FileSystemObject
If Not fso.FolderExists(PercorsoApp & "Fatture Elettroniche\Clienti") Then
 fso.CreateFolder PercorsoApp & "Fatture Elettroniche\Clienti"
End If
 
DocXml.Save PercorsoApp & "Fatture Elettroniche\Clienti\" & "IT" & rsInfoDitta("PartitaIva") & _
"_" & IdProgressivoDoc & ".xml"
Set DocXml = Nothing

CreaFattura = "IT" & rsInfoDitta("PartitaIva") & "_" & IdProgressivoDoc & ".xml"
End Function
Private Function IdProgressivo() As String
Dim rsProgressiviInvio As New ADODB.Recordset, NumProgressivo As Long
Dim CarIdProgressivo As Variant
CarIdProgressivo = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f", _
"g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "A", _
"B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
"W", "X", "Y", "Z")
rsProgressiviInvio.Open "SELECT ProgressivoInvio FROM FattureElettroniche", conn, adOpenStatic, adLockOptimistic
If rsProgressiviInvio.EOF Then
 NumProgressivo = 0
Else
 rsProgressiviInvio.MoveFirst
 NumProgressivo = rsProgressiviInvio("ProgressivoInvio") + 1
End If
While NumProgressivo > 62
 IdProgressivo = IdProgressivo & CarIdProgressivo((NumProgressivo Mod 62))
 NumProgressivo = NumProgressivo / 62
Wend
If rsProgressiviInvio.EOF Then
 rsProgressiviInvio.AddNew
End If
rsProgressiviInvio("ProgressivoInvio") = NumProgressivo
rsProgressiviInvio.Update
rsProgressiviInvio.Close
IdProgressivo = IdProgressivo & CarIdProgressivo((NumProgressivo Mod 62))
IdProgressivo = String(5 - Len(IdProgressivo), "0") & IdProgressivo
End Function
Public Function ImportaFattura(NomeFile As String, FormFatture As FattureFornitori, ByRef Errore As String) As Boolean
Dim DocXml As New DOMDocument
Dim rsDitte As New ADODB.Recordset
Dim Ditta As String
Dim PartitaIva As String
Dim Indirizzo As String
Dim Cap As String
Dim Comune As String
Dim Provincia As String
Dim Stato As String
Dim TipoDoc As String
Dim NumDoc As String
Dim DataDoc As String
Dim DatiRitenuta As IXMLDOMNode
Dim RigheDoc As IXMLDOMNodeList
Dim ElemXml As IXMLDOMNode
Dim ListaElemXml As IXMLDOMNodeList

DocXml.setProperty "SelectionLanguage", "XPath"
DocXml.Load NomeFile
Set ElemXml = DocXml.selectSingleNode("/*[local-name(.) = 'FatturaElettronica']/*[local-name(.) = 'FatturaElett" _
& "ronicaHeader']/*[local-name(.) = 'CedentePrestatore']/*[local-name(.) = 'DatiAnagrafici']/" _
& "*[local-name(.) = 'Anagrafica']/Denominazione")

If ElemXml Is Nothing Then
 Errore = "Campo Denominazione non trovato"
 Exit Function
End If

Ditta = ElemXml.Text

Set ElemXml = DocXml.selectSingleNode("/*[local-name(.) = 'FatturaElettronica']/*[local-name(.) = 'FatturaElett" _
& "ronicaHeader']/*[local-name(.) = 'CedentePrestatore']/*[local-name(.) = 'DatiAnagrafici']/" _
& "*[local-name(.) = 'IdFiscaleIva']/IdCodice")

PartitaIva = ElemXml.Text

Set ElemXml = DocXml.selectSingleNode("/*[local-name(.) = 'FatturaElettronica']/*[local-name(.) = 'FatturaElett" _
& "ronicaHeader']/*[local-name(.) = 'CedentePrestatore']/Sede")

If ElemXml Is Nothing Then
 Errore = "Campo Sede non trovato"
 Exit Function
End If

Set ListaElemXml = ElemXml.childNodes
For i = 0 To ListaElemXml.Length - 1
 Select Case ListaElemXml(i).nodeName
 Case "Indirizzo":
  Indirizzo = ListaElemXml(i).Text
 Case "CAP":
  Cap = ListaElemXml(i).Text
 Case "Comune":
  Comune = ListaElemXml(i).Text
 Case "Provincia":
  Provincia = ListaElemXml(i).Text
 Case "Nazione":
  On Error Resume Next
  Nazione = SigleStati(ListaElemXml(i).Text)
 End Select
Next i

Set ElemXml = DocXml.selectSingleNode("/*[local-name(.) = 'FatturaElettronica']/*[local-name(.) = 'FatturaElett" _
& "ronicaBody']/*[local-name(.) = 'DatiGenerali']/*[local-name(.) = 'DatiGeneraliDocumento']")

If ElemXml Is Nothing Then
 Errore = "Campo DatiGeneraliDocumento non trovato"
 Exit Function
End If

Set ListaElemXml = ElemXml.childNodes

For i = 0 To ListaElemXml.Length - 1
 Select Case ListaElemXml(i).nodeName
 Case "TipoDocumento":
  TipoDoc = ListaElemXml(i).Text
 Case "Numero":
  NumDoc = ListaElemXml(i).Text
 Case "Data":
  DataDoc = Replace(ListaElemXml(i).Text, "/", "-")
  DataDoc = Replace(DataDoc, ".", "-")
 End Select
Next i

If TipoDoc = "" Then
 Errore = "Campo TipoDocumento non trovato"
 Exit Function
End If

If NumDoc = "" Then
 Errore = "Campo Numero non trovato"
 Exit Function
End If

If DataDoc = "" Then
 Errore = "Campo Data non trovato"
 Exit Function
End If

Set DatiRitenuta = DocXml.selectSingleNode("/*[local-name(.) = 'FatturaElettronica']/*[local-name(.) = 'FatturaElett" _
& "ronicaBody']/*[local-name(.) = 'DatiGenerali']/*[local-name(.) = 'DatiGeneraliDocumento']/DatiRitenuta")

Dim ImportoRitenuta As String
Dim AliquotaRitenuta As String

If Not DatiRitenuta Is Nothing Then
 Set ListaElemXml = DatiRitenuta.childNodes
 For i = 0 To ListaElemXml.Length - 1
  Select Case ListaElemXml(i).nodeName
  Case "ImportoRitenuta":
   ImportoRitenuta = ListaElemXml(i).Text
  Case "AliquotaRitenuta":
   AliquotaRitenuta = ListaElemXml(i).Text
  End Select
 Next i
 
 If ImportoRitenuta = "" Then
  Errore = "Campo ImportoRitenuta non trovato"
  Exit Function
 End If
 
 If AliquotaRitenuta = "" Then
  Errore = "Campo AliquotaRitenuta non trovato"
  Exit Function
 End If
End If

Set DatiCassaPrev = DocXml.selectSingleNode("/*[local-name(.) = 'FatturaElettronica']/*[local-name(.) = 'FatturaElett" _
& "ronicaBody']/*[local-name(.) = 'DatiGenerali']/*[local-name(.) = 'DatiGeneraliDocumento']/DatiCassaPrevidenziale")

Dim RitenutaSuCassa As String
Dim AlCassa As String
Dim ImpCassa As String
Dim AliqIvaCassa As String

If Not DatiCassaPrev Is Nothing Then
 Set ListaElemXml = DatiCassaPrev.childNodes
 For i = 0 To ListaElemXml.Length - 1
  Select Case ListaElemXml(i).nodeName
  Case "Ritenuta":
   RitenutaSuCassa = ListaElemXml(i).Text
  Case "AlCassa":
   AliqCassa = ListaElemXml(i).Text
  Case "ImportoContributoCassa":
   ImpCassa = ListaElemXml(i).Text
  Case "AliquotaIva":
   AliqIvaCassa = ListaElemXml(i).Text
  End Select
 Next i
 
 If AliqCassa = "" Or IsNumeric(AliqCassa) = False Then
  Errore = "Campo AlCassa non trovato o non valido"
  Exit Function
 End If
 
 If ImpCassa = "" Or IsNumeric(ImpCassa) = False Then
  Errore = "Campo ImportoContributoCassa non trovato o non valido"
  Exit Function
 End If
 
 If AliqIvaCassa <> "" And IsNumeric(AliqIvaCassa) = False Then
  Errore = "Campo AliquotaIva non valido"
  Exit Function
 End If
End If

Set RigheDoc = DocXml.selectNodes("/*[local-name(.) = 'FatturaElettronica']/*[local-name(.) = 'FatturaElett" _
& "ronicaBody']/*[local-name(.) = 'DatiBeniServizi']/*[local-name(.) = 'DettaglioLinee']")

If RigheDoc.Length = 0 Then
 Errore = "Campo DettaglioLinee non trovato"
 Exit Function
End If

Dim Descrizione As String
Dim Quantita As String
Dim UnitaMisura As String
Dim AliquotaIva As String
Dim PrezzoUnitario As String
Dim PrezzoTotale As String
Dim TotImp As Double
Dim TotIva As Double
Dim TotDoc As Double

For i = 0 To RigheDoc.Length - 1
 Set ListaElemXml = RigheDoc(i).childNodes
 For j = 0 To ListaElemXml.Length - 1
  Select Case ListaElemXml(j).nodeName
  Case "Descrizione":
   Descrizione = ListaElemXml(j).Text
  Case "Quantita":
   Quantita = Replace(ListaElemXml(j).Text, ".", ",")
  Case "AliquotaIva":
   AliquotaIva = Replace(ListaElemXml(j).Text, ".", ",")
  Case "PrezzoUnitario":
   PrezzoUnitario = Replace(ListaElemXml(j).Text, ".", ",")
  Case "PrezzoTotale":
   PrezzoTotale = Replace(ListaElemXml(j).Text, ".", ",")
  End Select
 Next j
 
 If Descrizione = "" Then
  Errore = "Campo Descrizione non trovato"
  Exit Function
 End If

 If Quantita = "" Or IsNumeric(Quantita) = False Then
  Errore = "Campo Quantità non trovato o non valido"
  Exit Function
 End If

 If PrezzoUnitario = "" Or IsNumeric(PrezzoUnitario) = False Then
  Errore = "Campo PrezzoUnitario non trovato o non valido"
  Exit Function
  End If

 If PrezzoTotale = "" Or IsNumeric(PrezzoTotale) = False Then
  Errore = "Campo PrezzoTotale non trovato o non valido"
  Exit Function
 End If
 
 Quantita = FormatNumber(CDbl(Quantita), 2)
 PrezzoUnitario = FormatNumber(CDbl(PrezzoUnitario), 2)
 PrezzoTotale = FormatNumber(CDbl(PrezzoTotale), 2)
 
 TotImp = TotImp + (CDbl(PrezzoUnitario) * CDbl(Quantita))
 TotIva = TotIva + (CDbl(PrezzoUnitario) * CDbl(Quantita) * (CDbl(AliquotaIva) / 100))
Next i

rsDitte.Open "SELECT * FROM Fornitori WHERE Ditta = '" & Replace(Ditta, "'", "''") & "'", conn, _
adOpenDynamic, adLockOptimistic
If rsDitte.EOF Then
 rsDitte.AddNew
End If
rsDitte("Ditta") = Ditta
rsDitte("PartitaIva") = PartitaIva
rsDitte("Indirizzo") = Indirizzo
rsDitte("Loc") = Comune
rsDitte("Cap") = Cap
rsDitte("Prov") = Provincia
rsDitte("Stato") = Nazione
rsDitte("Rimosso") = False
rsDitte.Update

With FormFatture
 .CreaNuovoRecord True
 .TxtNumFattura.Text = NumDoc
 .Data.Text = DataDoc
 .TxtDitta.Text = Ditta
 .TxtDitta.Tag = rsDitte("Id")
 .TxtPartIva.Text = PartitaIva
 .TotDoc.Text = FormatNumber(TotImp + TotIva, 2)
 .TotImp.Text = FormatNumber(TotImp, 2)
 .TotIva.Text = FormatNumber(TotIva, 2)
 .TotNetto.Text = .TotDoc.Text
 .ElencoVoci.Rows = 1
 rsDitte.Close
 For i = 0 To RigheDoc.Length - 1
  .ElencoVoci.AddItem ""
  .ElencoVoci.RowHeight(.ElencoVoci.Rows - 1) = 315
  For j = 0 To ListaElemXml.Length - 1
   Select Case ListaElemXml(j).nodeName
   Case "Descrizione":
    .ElencoVoci.TextMatrix(i + 1, 0) = Replace(ListaElemXml(j).Text, ".", ",")
   Case "UnitaMisura":
    .ElencoVoci.TextMatrix(i + 1, 1) = Replace(ListaElemXml(j).Text, ".", ",")
   Case "Quantita":
    .ElencoVoci.TextMatrix(i + 1, 2) = Replace(ListaElemXml(j).Text, ".", ",")
   Case "AliquotaIva":
    .ElencoVoci.TextMatrix(i + 1, 4) = Replace(ListaElemXml(j).Text, ".", ",")
   Case "PrezzoUnitario":
    .ElencoVoci.TextMatrix(i + 1, 3) = Replace(ListaElemXml(j).Text, ".", ",")
   Case "PrezzoTotale":
    .ElencoVoci.TextMatrix(i + 1, 5) = Replace(ListaElemXml(j).Text, ".", ",")
   End Select
  Next j
 Next i
 .ElencoVoci.AddItem ""
 
 If Not DatiRitenuta Is Nothing Then
  .TxtRitenuta.Text = FormatNumber(ImportoRitenuta, 2)
  .TxtAliqRit.Text = FormatNumber(AliquotaRitenuta, 2)
  .TotNetto.Text = FormatNumber(CDbl(.TotDoc.Text) - CDbl(.TxtRitenuta.Text), 2)
  .ChkCalcRit.Value = 1
 End If
 
 If Not DatiCassaPrev Is Nothing Then
  If RitenutaSuCassa = "SI" Then
   .ChkRitCassa.Value = 1
  End If
  .TxtAliqCassaPro.Text = FormatNumber(AliqCassa, 2)
  .TxtCassaPro.Text = FormatNumber(ImpCassa, 2)
  If AliqIvaCassa <> "" Then
   .TxtAliqIva.Text = FormatNumber(AliqIvaCassa, 2)
  End If
 End If
End With

Set DocXml = Nothing
Dim fso As New FileSystemObject, CartellaFile As String
CartellaFile = Mid(NomeFile, 1, InStrRev(NomeFile, "\"))

If CartellaFile <> PercorsoApp & "Fatture Elettroniche\Fornitori" Then
 If Not fso.FolderExists(PercorsoApp & "Fatture Elettroniche\Fornitori") Then
  fso.CreateFolder PercorsoApp & "Fatture Elettroniche\Fornitori"
 End If
 fso.CopyFile NomeFile, PercorsoApp & "Fatture Elettroniche\Fornitori\" & Mid(NomeFile, InStrRev(NomeFile, "\") + 1)
End If
End Function
Private Sub Class_Initialize()
Set SigleStati = New Collection
With SigleStati
.Add "AD", "ANDORRA"
.Add "AE", "EMIRATI ARABI UNITI"
.Add "AG", "ANTIGUA / BARBUDA"
.Add "AI", "ANGUILLA"
.Add "AL", "ALBANIA"
.Add "AM", "ARMENIA"
.Add "AN", "ANTILLE OLANDESI"
.Add "AO", "ANGOLA"
.Add "AR", "ARGENTINA"
.Add "AS", "SAMOA AMERICANE"
.Add "AT", "AUSTRIA"
.Add "AU", "AUSTRALIA"
.Add "AW", "ARUBA"
.Add "AZ", "AZERBAIJAN"
.Add "BA", "BOSNIA"
.Add "BB", "BARBADOS"
.Add "BD", "BANGLADESH"
.Add "BE", "BELGIO"
.Add "BF", "BURKINA FASO"
.Add "BG", "BULGARIA"
.Add "BH", "BAHREIN"
.Add "BI", "BURUNDI"
.Add "BJ", "BENIN"
.Add "BM", "BERMUDE"
.Add "BN", "BRUNEI"
.Add "BO", "BOLIVIA"
.Add "BR", "BRASILE"
.Add "BS", "BAHAMAS"
.Add "BT", "BUTHAN"
.Add "BU", "BIRMANIA"
.Add "BV", "BOUVET"
.Add "BW", "BOTSWANA"
.Add "BY", "BIELORUSSIA"
.Add "BZ", "BELIZE"
.Add "CA", "CANADA"
.Add "CC", "ISOLE COCOSKEELING"
.Add "CF", "REPUBBLICA CENTROAFRICANA"
.Add "CG", "CONGO"
.Add "CH", "SVIZZERA"
.Add "CI", "COSTA D'AVORIO"
.Add "CK", "ISOLA DI COOK"
.Add "CL", "CILE"
.Add "CM", "CAMERUN"
.Add "CN", "CINA"
.Add "CO", "COLOMBIA"
.Add "CR", "COSTA RICA"
.Add "CU", "CUBA"
.Add "CV", "CAPO VERDE"
.Add "CX", "ISOLA CHRISTMAS"
.Add "CY", "CIPRO"
.Add "CZ", "REPUBBLICA CECA"
.Add "DE", "GERMANIA"
.Add "DJ", "GIBUTI"
.Add "DK", "DANIMARCA"
.Add "DM", "DOMINIQUE"
.Add "DO", "REPUBBLICA DOMINICANA"
.Add "DZ", "ALGERIA"
.Add "EC", "EQUADOR"
.Add "EE", "ESTONIA"
.Add "EG", "EGITTO"
.Add "ER", "ERITREA"
.Add "ES", "SPAGNA"
.Add "ET", "ETIOPIA"
.Add "FI", "FINLANDIA"
.Add "FJ", "ISOLE FIJI"
.Add "FK", "ISOLE FALKLAND"
.Add "FM", "MICRONESIA"
.Add "FO", "ISOLE FAROER"
.Add "FR", "FRANCIA"
.Add "GA", "GABON"
.Add "GB", "GRAN BRETAGNA"
.Add "GD", "GRENADE"
.Add "GE", "GEORGIA"
.Add "GF", "GUYANA FRANCESE"
.Add "GH", "GHANA"
.Add "GI", "GIBILTERRA"
.Add "GL", "GROENLANDIA"
.Add "GM", "GAMBIA"
.Add "GN", "GUINEA"
.Add "GP", "GUADALUPE"
.Add "GQ", "GUINEA EQUATORIALE"
.Add "GR", "GRECIA"
.Add "GS", "GEORGIA DEL SUD"
.Add "GT", "GUATEMALA"
.Add "GU", "GUAM"
.Add "GW", "GUINEA BISSAU"
.Add "GY", "GUYANA"
.Add "HK", "HONG KONG"
.Add "HM", "ISOLE HEART E MAC DONALD"
.Add "HN", "HONDURAS"
.Add "HR", "CROAZIA"
.Add "HT", "HAITI"
.Add "HU", "UNGHERIA"
.Add "ID", "INDONESIA"
.Add "IE", "IRLANDA"
.Add "IL", "ISRAELE"
.Add "IN", "INDIA"
.Add "IQ", "IRAQ"
.Add "IR", "IRAN"
.Add "IS", "ISLANDA"
.Add "IT", "ITALIA"
.Add "JM", "GIAMAICA"
.Add "JO", "GIORDANIA"
.Add "JP", "GIAPPONE"
.Add "KE", "KENIA"
.Add "KG", "KIRGHIZISTAN"
.Add "KH", "CAMBOGIA"
.Add "KI", "KIRIBATI"
.Add "KM", "COMORE"
.Add "KN", "ST.KITTS E NEVIS"
.Add "KP", "COREA DEL NORD"
.Add "KR", "COREA DEL SUD"
.Add "KW", "KUWAIT"
.Add "KY", "ISOLE CAIMANE"
.Add "KZ", "KAZAKISTAN"
.Add "LA", "LAOS"
.Add "LB", "LIBANO"
.Add "LC", "SANTA LUCIA"
.Add "LI", "LIECHTENSTEIN"
.Add "LK", "SRI LANKA"
.Add "LR", "LIBERIA"
.Add "LS", "LESOTHO"
.Add "LT", "LITUANIA"
.Add "LU", "LUSSEMBURGO"
.Add "LV", "LETTONIA"
.Add "LY", "LIBIA"
.Add "MA", "MAROCCO"
.Add "MC", "MONACO"
.Add "MD", "MOLDAVIA"
.Add "MG", "MADAGASCAR"
.Add "MH", "ISOLE MARSHALL"
.Add "MK", "MACEDONIA"
.Add "ML", "MALI"
.Add "MM", "MYANMAR"
.Add "MN", "MONGOLIA"
.Add "MO", "MACAO"
.Add "MP", "ISOLE MARIANNE SETTENTRIONALI"
.Add "MQ", "MARTINICA"
.Add "MR", "MAURITANIA"
.Add "MS", "MONTSERRAT"
.Add "MT", "MALTA"
.Add "MU", "MAURIZIUS"
.Add "MV", "MALDIVE"
.Add "MW", "MALAWI"
.Add "MX", "MESSICO"
.Add "MY", "MALESIA"
.Add "MZ", "MOZAMBICO"
.Add "NA", "NAMIBIA"
.Add "NC", "NUOVA CALEDONIA"
.Add "NE", "NIGER"
.Add "NF", "ISOLA DI NORFOLK"
.Add "NG", "NIGERIA"
.Add "NI", "NICARAGUA"
.Add "NL", "OLANDA"
.Add "NO", "NORVEGIA"
.Add "NP", "NEPAL"
.Add "NR", "NAURU"
.Add "NU", "ISOLA DI NIOUE"
.Add "NZ", "NUOVA ZELANDA"
.Add "OM", "OMAN"
.Add "PA", "PANAMA"
.Add "PE", "PERU"
.Add "PF", "POLINESIA FRANCESE"
.Add "PG", "PAPUA NUOVA GUINEA"
.Add "PH", "FILIPPINE"
.Add "PK", "PAKISTAN"
.Add "PL", "POLONIA"
.Add "PM", "SAINT PIERRE E MIQUELON"
.Add "PN", "ISOLE PITCAIRN"
.Add "PR", "PORTORICO"
.Add "PT", "PORTOGALLO"
.Add "PW", "PALAU"
.Add "PY", "PARAGUAY"
.Add "QA", "QATAR"
.Add "RE", "REUNION"
.Add "RO", "ROMANIA"
.Add "RU", "RUSSIA"
.Add "RW", "RUANDA"
.Add "SA", "ARABIA SAUDITA"
.Add "SB", "ISOLE SALOMONE"
.Add "SC", "SEYCHELLES"
.Add "SD", "SUDAN"
.Add "SE", "SVEZIA"
.Add "SG", "SINGAPORE"
.Add "SH", "ISOLA DI SANT'ELENA"
.Add "SI", "SLOVENIA"
.Add "SJ", "ISOLE SVALBARD E JAN MAYEN"
.Add "SK", "SLOVACCHIA"
.Add "SL", "SIERRA LEONE"
.Add "SM", "SAN MARINO"
.Add "SN", "SENEGAL"
.Add "SO", "SOMALIA"
.Add "SR", "SURINAME"
.Add "ST", "SAINT THOMAS E PRINCE"
.Add "SV", "EL SALVADOR"
.Add "SY", "SIRIA"
.Add "SZ", "SWAZILAND"
.Add "TC", "ISOLE DI TURKS E CAIQUES"
.Add "TD", "CIAD"
.Add "TG", "TOGO"
.Add "TH", "TAILANDIA"
.Add "TJ", "TAGICHISTAN"
.Add "TK", "TOKELAU"
.Add "TM", "TURKMENISTAN"
.Add "TN", "TUNISIA"
.Add "TO", "TONGA"
.Add "TP", "TIMOR ORIENTALE"
.Add "TR", "TURCHIA"
.Add "TT", "TRINIDAD E TOBAGO"
.Add "TV", "ISOLA DI TUVALU"
.Add "TW", "TAIWAN"
.Add "TZ", "TANZANIA"
.Add "UA", "UCRAINA"
.Add "UG", "UGANDA"
.Add "US", "STATI UNITI"
.Add "UY", "URUGUAY"
.Add "UZ", "UZBECHISTAN"
.Add "VA", "CITTA' DEL VATICANO"
.Add "VC", "SAINT VINCENT E GRANADINE"
.Add "VE", "VENEZUELA"
.Add "VG", "ISOLE VERGINI BRITANNICHE"
.Add "VI", "ISOLE VERGINI AMERICANE"
.Add "VN", "VIETNAM"
.Add "VU", "VANAUTU"
.Add "WF", "ISOLE WALLIS E FORTUNA"
.Add "WS", "SAMOA"
.Add "YD", "YEMEN DEL SUD"
.Add "YE", "YEMEN"
.Add "YT", "MAYOTTE"
.Add "ZM", "ZAMBIA"
.Add "ZR", "ZAIRE"
.Add "ZW", "ZIMBABWE"
End With
End Sub
