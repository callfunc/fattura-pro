VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form AnteprimaDoc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0033CCFF&
   Caption         =   "Anteprima Documento"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15180
   Icon            =   "Anteprima.frx":0000
   MinButton       =   0   'False
   ScaleHeight     =   149.49
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   267.759
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12000
      Top             =   2715
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   26
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anteprima.frx":4072
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   360
      Left            =   1320
      TabIndex        =   9
      Top             =   7410
      Width           =   8310
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4710
      Left            =   10710
      TabIndex        =   8
      Top             =   1995
      Width           =   360
   End
   Begin VB.CommandButton Stampa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3375
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Anteprima.frx":45AF
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Stampa"
      Top             =   195
      UseMaskColor    =   -1  'True
      Width           =   660
   End
   Begin VB.TextBox NumPag 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1620
      TabIndex        =   4
      Top             =   420
      Width           =   480
   End
   Begin VB.CommandButton PagSucc 
      Height          =   300
      Left            =   2685
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Anteprima.frx":4947
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton PagPrec 
      Height          =   300
      Left            =   495
      MaskColor       =   &H00D8E9EC&
      Picture         =   "Anteprima.frx":505D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.PictureBox PicContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   270
      ScaleHeight     =   87.577
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   180.711
      TabIndex        =   0
      Top             =   2100
      Width           =   10275
      Begin VB.PictureBox PicPrinter 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3795
         Left            =   1140
         ScaleHeight     =   66.94
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   116.417
         TabIndex        =   7
         Top             =   555
         Width           =   6600
      End
   End
   Begin VB.Label PagineDoc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2175
      TabIndex        =   5
      Top             =   450
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pagina"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1005
      TabIndex        =   3
      Top             =   450
      Width           =   540
   End
End
Attribute VB_Name = "AnteprimaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MargX = 5: Const MargY = 5
Const PaddX = 10: Const PaddY = 10
Dim DistMargX As Single, DistMargY As Single, SS As ServiziStampa
Dim WithEvents ssc As SmartSubClass, CaricamentoForm As Boolean
Attribute ssc.VB_VarHelpID = -1
Public Pagine As Integer
Private Sub Form_Load()
CaricamentoForm = True: Set ssc = New SmartSubClass
ssc.SubClassHwnd PicPrinter.hWnd, True: ssc.SubClassHwnd PicContainer.hWnd, True
NumPag.Text = "1": HScroll1.Min = 0: VScroll1.Min = 0
CaricamentoForm = False
End Sub
Private Sub Form_Resize()
RiposizionaControlli
End Sub
Private Sub RiposizionaControlli()
If Me.WindowState <> 1 And FatturaProMDI.WindowState <> 1 Then
 Dim LarghezzaCont As Single, AltezzaCont As Single
 LarghezzaCont = Me.ScaleWidth - (VScroll1.Width + MargX + 5)
 AltezzaCont = Me.ScaleHeight - (Stampa.Top + Stampa.Height + 10 + HScroll1.Height + 5)
 If LarghezzaCont > 0 Then
  PicContainer.Move MargX, Stampa.Top + Stampa.Height + 10, LarghezzaCont, PicContainer.Height
 End If
 If AltezzaCont > 0 Then
  PicContainer.Move MargX, Stampa.Top + Stampa.Height + 10, PicContainer.Width, AltezzaCont
 End If
 PicPrinter.Width = SS.AreaStampaX: PicPrinter.Height = SS.AreaStampaY
 DistMargX = ((PicContainer.ScaleWidth - PaddX * 2) - PicPrinter.Width) / 2
 DistMargY = ((PicContainer.ScaleHeight - PaddY * 2) - PicPrinter.Height) / 2
 If DistMargX < 0 Then DistMargX = 0
 If DistMargY < 0 Then DistMargY = 0
 DistMargX = DistMargX + PaddX: DistMargY = DistMargY + PaddY
 PicPrinter.Left = (HScroll1.Value * -1) + DistMargX
 PicPrinter.Top = (VScroll1.Value * -1) + DistMargY
 VScroll1.Left = PicContainer.Left + PicContainer.Width
 VScroll1.Top = PicContainer.Top: VScroll1.Height = PicContainer.Height
 HScroll1.Left = PicContainer.Left: HScroll1.Top = PicContainer.Top + PicContainer.Height
 HScroll1.Min = 0: VScroll1.Min = 0: HScroll1.Width = PicContainer.Width
 If PicPrinter.Height > (PicContainer.ScaleHeight - PaddY * 2) Then
  VScroll1.Max = PicPrinter.Height - (PicContainer.ScaleHeight - PaddY * 2)
 Else
  VScroll1.Max = 0
 End If
 If PicPrinter.Width > (PicContainer.ScaleWidth - PaddX * 2) Then
  HScroll1.Max = PicPrinter.Width - (PicContainer.ScaleWidth - PaddX * 2)
 Else
  HScroll1.Max = 0
 End If
End If
End Sub
Public Sub Inizializza(ClasseStampa As ServiziStampa)
Set SS = ClasseStampa
End Sub
Private Sub Form_Unload(Cancel As Integer)
ssc.SubClassHwnd PicPrinter.hWnd, False
ssc.SubClassHwnd PicContainer.hWnd, False
Set AnteprimaDoc = Nothing
End Sub
Private Sub HScroll1_Change()
PicPrinter.Left = (HScroll1.Value * -1) + DistMargX
End Sub
Private Sub HScroll1_GotFocus()
PicPrinter.SetFocus
End Sub
Private Sub HScroll1_Scroll()
PicPrinter.Left = (HScroll1.Value * -1) + DistMargX
End Sub
Private Sub PagPrec_Click()
If Val(NumPag) > 1 Then
 NumPag = CInt(NumPag) - 1
End If
PicPrinter.SetFocus
End Sub
Private Sub PagSucc_Click()
If Val(NumPag) < Pagine Then
 NumPag = CInt(NumPag) + 1
End If
PicPrinter.SetFocus
End Sub
Private Sub NumPag_Change()
If Not CaricamentoForm Then
 If Val(NumPag) > 0 And Val(NumPag) <= Pagine Then
  PicPrinter.Cls: Call SS.Stampa(CInt(NumPag))
 End If
End If
End Sub
Private Sub NumPag_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 And KeyAscii <> 8) Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub ssc_NewMessage(ByVal hWnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
If uMsg = WM_MOUSEWHEEL Then
 Dim Rotazione As Long: Rotazione = wParam / 65536
 Call AggiornaPosizioneBarra(Rotazione)
End If
End Sub
Private Sub AggiornaPosizioneBarra(ByVal Rotazione As Long)
Dim PosBarra As Long
PosBarra = VScroll1.Value + (5 * IIf(Rotazione < 0, 1, -1))
If PosBarra >= 0 And PosBarra <= VScroll1.Max Then
 VScroll1.Value = PosBarra
ElseIf PosBarra > VScroll1.Max Then
 VScroll1.Value = VScroll1.Max
Else
 VScroll1.Value = 0
End If
End Sub
Private Sub Stampa_Click()
OpzioniStampa.Inizializza SS: OpzioniStampa.Show vbModal
If OpzioniStampa.Scelta = "Stampa" Then
 SS.ImpostaAnteprima False: SS.Stampa -1
 SS.ImpostaAnteprima True
End If
End Sub
Private Sub VScroll1_Change()
PicPrinter.Top = (VScroll1.Value * -1) + DistMargY
End Sub
Private Sub VScroll1_GotFocus()
PicPrinter.SetFocus
End Sub
Private Sub VScroll1_Scroll()
PicPrinter.Top = (VScroll1.Value * -1) + DistMargY
End Sub
