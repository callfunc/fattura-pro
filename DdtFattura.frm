VERSION 5.00
Begin VB.Form DdtFattura 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ddt Fattura"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4785
   Icon            =   "DdtFattura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnAggVettore 
      Height          =   345
      Left            =   4125
      Picture         =   "DdtFattura.frx":4072
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   420
   End
   Begin VB.ComboBox CmbVettori 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   225
      TabIndex        =   6
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox TxtNumDoc 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   225
      TabIndex        =   2
      Top             =   525
      Width           =   1830
   End
   Begin VB.TextBox TxtDataDoc 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   225
      TabIndex        =   1
      Top             =   1230
      Width           =   1845
   End
   Begin VB.CommandButton BtnSalva 
      Caption         =   "Salva"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1845
      TabIndex        =   0
      Top             =   3585
      Width           =   1110
   End
   Begin VB.Label LblVettore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vettore"
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
      Left            =   225
      TabIndex        =   5
      Top             =   1665
      Width           =   570
   End
   Begin VB.Label LblNumDdt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero:"
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
      Left            =   225
      TabIndex        =   4
      Top             =   270
      Width           =   705
   End
   Begin VB.Label LblDataDdt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
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
      Left            =   225
      TabIndex        =   3
      Top             =   975
      Width           =   405
   End
End
Attribute VB_Name = "DdtFattura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NumDoc As String
Public DataDoc As String
Public Vettore As Long
Public SalvaModifiche As Boolean
Dim rsVettori As ADODB.Recordset
Private Sub BtnAggVettore_Click()
Set Vettori.FormChiamante = DdtFattura
Vettori.Show vbModal
rsVettori.Requery
If Not rsVettori.EOF Then
 CaricaVettori
End If
End Sub
Private Sub BtnSalva_Click()
SalvaDdt
End Sub
Private Sub SalvaDdt()
Dim d As Variant
If TxtNumDoc.Text = "" Then
 MsgBox "Inserire il numero del ddt !", vbExclamation, "Fattura Pro"
 TxtNumDoc.SetFocus: TxtNumDoc.SelStart = 0: TxtNumDoc.SelLength = Len(TxtNumDoc.Text)
 Exit Sub
End If

If IsDate(TxtDataDoc.Text) Then
 d = Split(TxtDataDoc.Text, "-")
 If UBound(d) = 2 Then
  If Len(d(0)) < 2 Then d(0) = "0" & d(0)
  If Len(d(1)) < 2 Then d(1) = "0" & d(1)
  TxtDataDoc.Text = d(0) & "-" & d(1) & "-" & d(2)
 Else: e = True
 End If
Else: e = True
End If

If e Then
 MsgBox "Attenzione, inserire una data valida !", vbExclamation, "Fattura Pro"
 TxtDataDoc.SetFocus: TxtDataDoc.SelStart = 0: TxtDataDoc.SelLength = Len(TxtDataDoc.Text)
 Exit Sub
End If

NumDoc = TxtNumDoc.Text
DataDoc = TxtDataDoc.Text
If CmbVettori.ListIndex <> -1 Then
 Vettore = CmbVettori.ItemData(CmbVettori.ListIndex)
End If
SalvaModifiche = True
Me.Hide
End Sub
Private Sub Form_Load()
TxtNumDoc.Text = NumDoc
TxtDataDoc.Text = DataDoc
Set rsVettori = New ADODB.Recordset
rsVettori.Open "SELECT Id, Ditta FROM Vettori", conn, adOpenDynamic, adLockOptimistic
If Not rsVettori.EOF Then
 CaricaVettori
End If
End Sub
Private Sub CaricaVettori()
rsVettori.MoveFirst
While Not rsVettori.EOF
 With CmbVettori
 .AddItem rsVettori("Ditta")
 .ItemData(.NewIndex) = rsVettori("Id")
 If rsVettori("Id") = Vettore Then
  .ListIndex = .NewIndex
 End If
 End With
 rsVettori.MoveNext
Wend
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
 Cancel = 1
 Me.Hide
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set DdtFattura = Nothing
End Sub
Private Sub TxtDataDoc_KeyPress(KeyAscii As Integer)
If InStr("0123456789-" & vbBack, Chr(KeyAscii)) = 0 Then
 KeyAscii = 0
End If
End Sub

