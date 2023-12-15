VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Backup 
   BackColor       =   &H0033CCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Dati"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerConn 
      Interval        =   100
      Left            =   4020
      Top             =   225
   End
   Begin MSComctlLib.ProgressBar Preload 
      Height          =   330
      Left            =   525
      TabIndex        =   1
      Top             =   930
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label LblProgress 
      Alignment       =   2  'Center
      BackColor       =   &H0033CCFF&
      Caption         =   "Completato al 0 %"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   1380
      Width           =   4740
   End
   Begin VB.Label LblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0033CCFF&
      Caption         =   "Backup in corso..."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00007800&
      Height          =   270
      Left            =   1350
      TabIndex        =   0
      Top             =   360
      Width           =   2070
   End
End
Attribute VB_Name = "Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WM_PAINT = &HF
Private Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type

Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, _
ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, _
ByVal lContext As Long) As Long
Private Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" _
(ByVal hFtpSession As Long, ByVal sBuff As String, ByVal Access As Long, ByVal Flags As Long, _
ByVal Context As Long) As Long
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
(ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function ReadFile Lib "Kernel32" (ByVal hFile As Long, lpBuffer As Any, _
ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function InternetWriteFile Lib "wininet.dll" _
(ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToWite As Long, _
dwNumberOfBytesWritten As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, _
ByVal nHeight As Long) As Long

Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_FLAG_PASSIVE = &H8000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const FTP_TRANSFER_TYPE_BINARY = &H2
Private Const GENERIC_WRITE = &H40000000
Const BUFFER_READ = 2048
Dim WithEvents ssc As SmartSubClass, rsImpostazioni As ADODB.Recordset
Attribute ssc.VB_VarHelpID = -1
Private Sub Form_Load()
Preload.Max = 100
Preload.Value = 0
Set ssc = New SmartSubClass
ssc.SubClassHwnd Preload.hWnd, True
Set rsImpostazioni = New ADODB.Recordset
rsImpostazioni.Open "Impostazioni", conn, adOpenDynamic, adLockOptimistic
rsImpostazioni.MoveFirst
End Sub
Public Sub DisegnaPreload(ByVal hWnd As Long)
  
Dim hdc As Long, hdc_mem As Long, hbmp_mem As Long, hPen As Long
Dim hOld As Long, PosCorr As Long, PixCorr As Long
Dim Rosso As Long, Verde As Long, Blu As Long
    
Dim RetControllo As RECT, RetCaricamento As RECT
    
GetWindowRect hWnd, RetControllo
    
RetControllo.Right = RetControllo.Right - RetControllo.Left
RetControllo.Bottom = RetControllo.Bottom - RetControllo.Top
RetControllo.Left = 0
RetControllo.Top = 0

hdc = GetWindowDC(hWnd)

RetCaricamento.Right = RetControllo.Right
RetCaricamento.Bottom = RetControllo.Bottom
RetCaricamento.Left = 1
RetCaricamento.Top = 1

hdc_mem = CreateCompatibleDC(hdc)
hbmp_mem = CreateCompatibleBitmap(hdc, RetCaricamento.Right - RetCaricamento.Left, _
RetCaricamento.Bottom - RetCaricamento.Top)

SelectObject hdc_mem, hbmp_mem

FillRect hdc_mem, RetControllo, CreateSolidBrush(RGB(240, 240, 240))
FillRect hdc_mem, RetCaricamento, CreateSolidBrush(RGB(200, 200, 200))

With Preload
 PosCorr = .Value * (RetCaricamento.Right - RetCaricamento.Left) / (.Max - .Min)
End With

For PixCorr = RetCaricamento.Left To PosCorr
 Rosso = 0
 Verde = 50 + (PixCorr * (160 - 50) / (RetCaricamento.Right - RetCaricamento.Left))
 Blu = 0
        
 hPen = CreatePen(PS_SOLID, 1, RGB(Rosso, Verde, Blu))
 hOld = SelectObject(hdc_mem, hPen)
        
 MoveToEx hdc_mem, PixCorr, RetCaricamento.Top, 0
 LineTo hdc_mem, PixCorr, RetCaricamento.Bottom
        
 SelectObject hdc_mem, hOld
 DeleteObject hPen
    
Next PixCorr

BitBlt hdc, RetCaricamento.Left, RetCaricamento.Top, RetCaricamento.Right - RetCaricamento.Left, _
RetCaricamento.Bottom - RetCaricamento.Top, hdc_mem, RetCaricamento.Left, RetCaricamento.Top, vbSrcCopy
DeleteDC hdc_mem
ReleaseDC hWnd, hdc
ValidateRect hWnd, 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
ssc.SubClassHwnd Preload.hWnd, False
End Sub
Private Sub ssc_NewMessage(ByVal hWnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
If uMsg = WM_PAINT Then
 DisegnaPreload hWnd
End If
End Sub
Private Sub TimerConn_Timer()
TimerConn.Enabled = False
Dim BackupCompletato As Boolean
While Not BackupCompletato
 Dim int_open As Long, int_conn As Long, descr_file As Long, dim_file As Long, byte_inviati As Long, _
 perc_compl%, byte_scritti As Long, BUFFER_WRITE As Long, Scelta%
 Dim arr_file(BUFFER_READ) As Byte, ftp_file As Long
 BUFFER_WRITE = 2048

 int_open = InternetOpen("FatturaPro", 0, vbNullString, vbNullString, 0)
 int_conn = InternetConnect(int_open, "ftp.casalgismondo.it", 21, "casalgismondo", "34n.EMGHN!", _
 INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
 If int_conn = 0 Then 'And (Err.LastDllError = 12029 Or Err.LastDllError = 12007)
  MsgBox "Backup fallito. Connessione al server non riuscita !", vbExclamation, "Fattura Pro"
  Unload Me
  Exit Sub
 End If
 dim_file = FileLen(App.Path & "\FatturaPro.mdb")
 descr_file = FreeFile

 ftp_file = FtpOpenFile(int_conn, "/www.casalgismondo.it/fp/FatturaPro.mdb", GENERIC_WRITE, _
 FTP_TRANSFER_TYPE_BINARY, INTERNET_FLAG_NO_CACHE_WRITE)
 If ftp_file = 0 Then
  Exit Sub
 End If

 Open PercorsoApp & "FatturaPro.mdb" For Binary As #descr_file
 Preload.Value = 0
 LblProgress.Caption = "Completato al 0 %"
 
 Do While byte_inviati < dim_file
  Get #descr_file, , arr_file

  If dim_file - byte_inviati < BUFFER_WRITE Then
   BUFFER_WRITE = dim_file - byte_inviati
  End If
  If InternetWriteFile(ftp_file, arr_file(0), BUFFER_WRITE, byte_scritti) = 0 Then
   LblProgress.ForeColor = RGB(255, 0, 0)
   LblProgress.Caption = "Errore. Backup fallito !"
   
   Scelta = MsgBox("Il backup dei dati non è riuscito !. Riprovare ?", vbYesNo + vbQuestion, "Fattura Pro")
   If Scelta = vbNo Then
    BackupCompletato = True
   Else
    BackupCompletato = False
   End If
   Exit Do
  ElseIf byte_scritti <> BUFFER_WRITE Then
   LblProgress.ForeColor = RGB(255, 0, 0)
   LblProgress.Caption = "Errore. Backup fallito !"
   Scelta = MsgBox("Il backup dei dati non è riuscito !. Riprovare ?", vbYesNo + vbQuestion, "Fattura Pro")
   If Scelta = vbNo Then
    BackupCompletato = True
   Else
    BackupCompletato = False
   End If
   Exit Do
  Else
   byte_inviati = byte_inviati + BUFFER_WRITE
   perc_compl = Round(byte_inviati * 100 / dim_file)
   Preload.Value = perc_compl
   LblProgress.Caption = "Completato al " & perc_compl & " %"
   DoEvents
  End If
 Loop
 If byte_inviati = dim_file Then
  BackupCompletato = True
  rsImpostazioni("Backup") = False
  rsImpostazioni.Update
 End If
 Close #descr_file
 InternetCloseHandle int_open
Wend
rsImpostazioni.Close
Unload Me
End Sub
