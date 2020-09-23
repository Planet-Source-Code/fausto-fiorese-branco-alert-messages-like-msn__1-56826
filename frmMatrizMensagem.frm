VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMatrizMensagem 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMensagem 
      Interval        =   1000
      Left            =   1260
      Top             =   0
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   495
      Top             =   -180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   181
      ImageHeight     =   116
      MaskColor       =   14737632
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMatrizMensagem.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMatrizMensagem.frx":0B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMatrizMensagem.frx":1745
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMatrizMensagem.frx":22E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMatrizMensagem.frx":2EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMatrizMensagem.frx":3973
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgClose 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2160
      Picture         =   "frmMatrizMensagem.frx":3DB5
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblMensagem 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   2625
   End
   Begin VB.Image imgAcao 
      Height          =   1740
      Left            =   0
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "frmMatrizMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Primeira As Boolean
Public IndiceMensagem As Integer
Public tp_Mensagem As Integer
Public tp_Fechamento As Integer '0 = Automatico; 1 = Manual
Public tp_Beep As Boolean
Public tp_EstiloFechamento As Integer '0 = Nenhum; 1 = P/ Baixo; 2 = P/ Direita
Public tp_EstiloAbertura As Integer '0 = Nenhum; 1 = P/ Baixo; 2 = P/ Direita


Private Declare Function Beep Lib "kernel32" (ByVal lngFreq As Long, ByVal lngDuration As Long) As Long

Sub FechaMensagem()

    Dim i As Integer
    Dim Tamanho As Double
    
    Select Case tp_EstiloFechamento
           Case 1
               Tamanho = Me.Height
               For i = 0 To Tamanho
                   Me.Height = Tamanho - i
                   DoEvents
               Next i
           Case 2
               Tamanho = Me.Width
               For i = Me.Left To (Me.Left + Tamanho)
                   Me.Left = i
                   DoEvents
               Next i
    End Select
    
    Unload Me
    MensagemAtiva(IndiceMensagem) = False
End Sub

Sub AbreMensagem()
    
    Dim i As Integer
    Dim Tamanho As Double
    
    Select Case tp_EstiloAbertura
           Case 1
               
               Tamanho = Me.Height
               Me.Height = 0
               Me.Show
               For i = 0 To Tamanho
                   Me.Height = i
                   DoEvents
               Next i
           Case 2
               Tamanho = Me.Left
               Me.Left = Me.Left + Me.Width
               Me.Show
               For i = Me.Left To Tamanho Step -1
                   Me.Left = i
                   DoEvents
               Next i
           Case Else
               Me.Show
    End Select
    
 
End Sub
Private Sub Form_Load()
    
    Dim Result
    
    'If tp_Mensagem = 1 Then
    '   lblMensagem.BackColor = RGB(185, 250, 255)
    '   Me.BackColor = RGB(185, 250, 255)
    'End If
    
    
    Select Case tp_Mensagem
           Case 1
               imgAcao.Picture = IMG.ListImages(1).Picture
           Case 2
               imgAcao.Picture = IMG.ListImages(3).Picture
           Case 3
               imgAcao.Picture = IMG.ListImages(2).Picture
           Case 4
               imgAcao.Picture = IMG.ListImages(4).Picture
           Case 5
               imgAcao.Picture = IMG.ListImages(5).Picture
    End Select
           
    'imgClose.Picture = IMG.ListImages(42).Picture
    imgClose.Picture = Nothing
    imgClose.Height = 380
    imgClose.Width = 380
    imgClose.Top = 0
    imgClose.Left = Me.Width - imgClose.Width
    
    
    Me.Left = Screen.Width - Me.Width
    Me.Top = Screen.Height - (Me.Height * (IndiceMensagem + 1))
    tmrMensagem.Interval = 1000 * qt_TempoPermanencia
    tmrMensagem.Enabled = True
    Primeira = False
    If tp_Beep = True Then
       Call Beep(460, 60)
       Call Beep(260, 60)
    End If
    
    
    'Result = CreateRoundRectRgn(0, 0, Me.Width / 2, (Me.Height / 2), 40, 40)
    'Result = SetWindowRgn(Me.hwnd, Result, True)
    
End Sub


Private Sub imgClose_Click()
    Call FechaMensagem
End Sub

Private Sub tmrMensagem_Timer()
    
    If Primeira = False Then
       Primeira = True
    Else
       If tp_Fechamento = 1 Then
          tmrMensagem.Enabled = False
          Exit Sub
       End If
       Call FechaMensagem
    End If
    
End Sub


