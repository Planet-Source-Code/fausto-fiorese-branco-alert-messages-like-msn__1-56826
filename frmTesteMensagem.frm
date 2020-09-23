VERSION 5.00
Begin VB.Form frmTesteMensagem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensagens de Alerta Estilo MSN"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Type Open"
      Height          =   1320
      Left            =   4725
      TabIndex        =   18
      Top             =   900
      Width           =   2490
      Begin VB.OptionButton optAbreCima 
         Caption         =   "of Top"
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   585
         Width           =   1230
      End
      Begin VB.OptionButton optAbreNenhum 
         Caption         =   "None"
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   270
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton optAbreDireita 
         Caption         =   "of Right"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   945
         Width           =   1230
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Type Close"
      Height          =   1320
      Left            =   2070
      TabIndex        =   14
      Top             =   900
      Width           =   2535
      Begin VB.OptionButton optDireita 
         Caption         =   "To Right"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   945
         Width           =   1230
      End
      Begin VB.OptionButton optNenhum 
         Caption         =   "None"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   270
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton optCima 
         Caption         =   "To Top"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   585
         Width           =   1230
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Beep"
      Height          =   870
      Left            =   4680
      TabIndex        =   11
      Top             =   0
      Width           =   2535
      Begin VB.OptionButton optSim 
         Caption         =   "Yes"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton optNao 
         Caption         =   "No"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   585
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Close"
      Height          =   870
      Left            =   2070
      TabIndex        =   8
      Top             =   0
      Width           =   2535
      Begin VB.OptionButton optManual 
         Caption         =   "Manual"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   585
         Width           =   1230
      End
      Begin VB.OptionButton optAutomatico 
         Caption         =   "Automatic"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   270
         Value           =   -1  'True
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type"
      Height          =   1905
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   1950
      Begin VB.OptionButton optTipo1 
         Caption         =   "Alert"
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   1770
      End
      Begin VB.OptionButton optTipo2 
         Caption         =   "Information"
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   540
         Width           =   1770
      End
      Begin VB.OptionButton optTipo3 
         Caption         =   "Erro"
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   855
         Width           =   1770
      End
      Begin VB.OptionButton optTipo4 
         Caption         =   "Question"
         Height          =   285
         Left            =   90
         TabIndex        =   4
         Top             =   1170
         Width           =   1770
      End
      Begin VB.OptionButton optTipo5 
         Caption         =   "None"
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   1485
         Value           =   -1  'True
         Width           =   1770
      End
   End
   Begin VB.TextBox txtMensagem 
      Height          =   1590
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2295
      Width           =   5955
   End
   Begin VB.CommandButton cmdTeste 
      Caption         =   "OK"
      Height          =   510
      Left            =   6075
      TabIndex        =   0
      Top             =   3375
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Message:"
      Height          =   240
      Left            =   90
      TabIndex        =   22
      Top             =   2070
      Width           =   1815
   End
End
Attribute VB_Name = "frmTesteMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdTeste_Click()

    Dim tp_Fechamento As Integer
    Dim tp_EstiloFechamento As Integer
    Dim tp_EstiloAbertura As Integer
    Dim tp_Beep As Boolean
    
    If optAutomatico.Value = True Then
       tp_Fechamento = 0
    End If
    If optManual.Value = True Then
       tp_Fechamento = 1
    End If
    
    If optSim.Value = True Then
       tp_Beep = True
    Else
       tp_Beep = False
    End If
    
    If optNenhum.Value = True Then
       tp_EstiloFechamento = 0
    End If
    If optCima.Value = True Then
       tp_EstiloFechamento = 1
    End If
    If optDireita.Value = True Then
       tp_EstiloFechamento = 2
    End If
    
    If optAbreNenhum.Value = True Then
       tp_EstiloAbertura = 0
    End If
    If optAbreCima.Value = True Then
       tp_EstiloAbertura = 1
    End If
    If optAbreDireita.Value = True Then
       tp_EstiloAbertura = 2
    End If
    
    
    If optTipo1.Value = True Then
       Call msgAlerta(txtMensagem.Text, 1, tp_Fechamento, Me, tp_Beep, tp_EstiloFechamento, tp_EstiloAbertura)
    End If
    If optTipo2.Value = True Then
       Call msgAlerta(txtMensagem.Text, 2, tp_Fechamento, Me, tp_Beep, tp_EstiloFechamento, tp_EstiloAbertura)
    End If
    If optTipo3.Value = True Then
       Call msgAlerta(txtMensagem.Text, 3, tp_Fechamento, Me, tp_Beep, tp_EstiloFechamento, tp_EstiloAbertura)
    End If
    If optTipo4.Value = True Then
       Call msgAlerta(txtMensagem.Text, 4, tp_Fechamento, Me, tp_Beep, tp_EstiloFechamento, tp_EstiloAbertura)
    End If
    If optTipo5.Value = True Then
       Call msgAlerta(txtMensagem.Text, 5, tp_Fechamento, Me, tp_Beep, tp_EstiloFechamento, tp_EstiloAbertura)
    End If
    

End Sub


Private Sub Form_Unload(Cancel As Integer)

Cancel = -1

Me.WindowState = vbMinimized

End Sub


