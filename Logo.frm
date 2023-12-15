VERSION 5.00
Begin VB.Form Logo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0033CCFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox LogoDitta 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2385
      Left            =   300
      Picture         =   "Logo.frx":0000
      ScaleHeight     =   2385
      ScaleWidth      =   2580
      TabIndex        =   3
      Top             =   315
      Width           =   2580
   End
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   3150
      Top             =   255
   End
   Begin VB.Image LogoProg 
      Height          =   675
      Left            =   4260
      Picture         =   "Logo.frx":4FED
      Top             =   315
      Width           =   675
   End
   Begin VB.Label LblInfoProg 
      Alignment       =   2  'Center
      BackColor       =   &H0033CCFF&
      Caption         =   "Realizzato da: Antonio Maugeri"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00007800&
      Height          =   300
      Left            =   2940
      TabIndex        =   2
      Top             =   2085
      Width           =   3405
   End
   Begin VB.Label LblNomeProg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fattura Pro 1.0"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00007800&
      Height          =   435
      Left            =   3450
      TabIndex        =   1
      Top             =   1065
      Width           =   2340
   End
   Begin VB.Label LblTipoProg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0033CCFF&
      Caption         =   "Software per la fatturazione"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00007800&
      Height          =   270
      Left            =   3300
      TabIndex        =   0
      Top             =   1545
      Width           =   2700
   End
End
Attribute VB_Name = "Logo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ColoreBordo As Long
Private Sub Form_Load()
Me.DrawWidth = 1: ColoreBordo = RGB(0, 120, 0)
Me.Line (0, 0)-(Me.Width - 15, Me.Height - 15), ColoreBordo, B
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Logo = Nothing
End Sub
Private Sub Timer1_Timer()
If IsObject(Printer) Then
 On Error Resume Next
 Printer.ScaleMode = vbMillimeters: Printer.Orientation = vbPRORPortrait
 Printer.PaperSize = vbPRPSA4: Printer.ColorMode = vbPRCMColor
 Printer.Duplex = vbPRDPSimplex
 If Err.Number = 0 Then
  StampaConf = True
 End If
 On Error GoTo 0
End If
FatturaProMDI.Visible = True: FatturaProMDI.WindowState = vbMaximized
Timer1.Enabled = False
Me.Hide
Dim rsDatiDitta As New ADODB.Recordset
rsDatiDitta.Open "SELECT * FROM InfoDitta", conn, adOpenDynamic, adLockOptimistic
If Not rsDatiDitta.EOF Then
 Dim CampiObbligatori As Variant, InserisciInfoDitta As Boolean
 CampiObbligatori = Array("Azienda", "Ditta", "SedeLegale", "SedeAziendale", "PartitaIVA")
 For i = 0 To UBound(CampiObbligatori)
  If rsDatiDitta(CampiObbligatori(i)) = "" Or IsNull(rsDatiDitta(CampiObbligatori(i))) Then
   InserisciDatiDitta = True: Exit For
  End If
 Next i
 If InserisciDatiDitta Then
  Set InfoDitta.rsInfoDitta = rsDatiDitta
  InfoDitta.Show vbModal
 End If
Else
 Set InfoDitta.rsInfoDitta = rsDatiDitta
 InfoDitta.Show vbModal
End If
End Sub
