Attribute VB_Name = "mdlMensagem"
Option Explicit
Global Const winding = 2
Global Const alternate = 1
Global Const rgn_or = 2


Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function CreateEllipticRgn& Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

Public qt_MensagensAtivas As Double
Public frmMensagem(10) As New frmMatrizMensagem
Public MensagemAtiva(10) As Boolean

Public Const qt_TempoPermanencia = 5

Sub msgAlerta(ds_Mensagem As String, tp_Mensagem As Integer, tp_Fechamento As Integer, frmOrigem As Form, tp_Beep As Boolean, tp_EstiloFechamento As Integer, tp_EstiloAbertura As Integer)

    Dim IndiceMensagem As Integer
    
    For IndiceMensagem = 0 To 9
        If MensagemAtiva(IndiceMensagem) = False Then
           frmMensagem(IndiceMensagem).tp_Beep = tp_Beep
           frmMensagem(IndiceMensagem).tp_Mensagem = tp_Mensagem
           frmMensagem(IndiceMensagem).tp_Fechamento = tp_Fechamento
           frmMensagem(IndiceMensagem).tp_EstiloAbertura = tp_EstiloAbertura
           frmMensagem(IndiceMensagem).tp_EstiloFechamento = tp_EstiloFechamento
           frmMensagem(IndiceMensagem).IndiceMensagem = IndiceMensagem
           frmMensagem(IndiceMensagem).lblMensagem.Caption = ds_Mensagem
           frmMensagem(IndiceMensagem).AbreMensagem
           MensagemAtiva(IndiceMensagem) = True
           frmOrigem.SetFocus
           Exit For
        End If
        DoEvents
    Next IndiceMensagem

End Sub


