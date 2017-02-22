VERSION 5.00
Begin VB.Form frmChecaIE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Checagem do IE por Estado"
   ClientHeight    =   3570
   ClientLeft      =   1425
   ClientTop       =   1980
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4350
   Begin VB.Frame fraIE2 
      Caption         =   "Inscrição Estadual 2 - ERRADO!"
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   60
      TabIndex        =   6
      Top             =   2400
      Width           =   4215
      Begin VB.CommandButton cmdAcheEstado 
         Caption         =   "&Ache Estado"
         Height          =   315
         Left            =   2880
         TabIndex        =   8
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txtChecaIE2 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Text            =   "76105790"
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label lblEstado 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   660
         Width           =   2895
      End
      Begin VB.Label lblIE2 
         Caption         =   "Digite um I.E.:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame fraIE1 
      Caption         =   "Inscrição Estadual 1 - CERTO!"
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4215
      Begin VB.TextBox txtChecaIE 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Text            =   "250768135"
         Top             =   300
         Width           =   1575
      End
      Begin VB.CommandButton cmdChecaIE 
         Caption         =   "&Checa IE"
         Height          =   315
         Left            =   2880
         TabIndex        =   3
         Top             =   660
         Width           =   1215
      End
      Begin VB.ComboBox cboEstado 
         Height          =   315
         ItemData        =   "frmChecaIE.frx":0000
         Left            =   1200
         List            =   "frmChecaIE.frx":0066
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   660
         Width           =   1575
      End
      Begin VB.CheckBox chkIsento 
         Caption         =   "&Isento"
         Height          =   315
         Left            =   2880
         TabIndex        =   1
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblChecaIE 
         Caption         =   "Digite um I.E.:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Label lblExplicacao 
      Caption         =   $"frmChecaIE.frx":019F
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   4095
   End
End
Attribute VB_Name = "frmChecaIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ---------------
' Dois Exemplos:
' 250768135 - SC
' 76105790  - RJ
' ---------------

Dim objChecaIE As New clsChecaIE

Private Sub chkIsento_Click()

    If chkIsento.Value Then
        txtChecaIE.Text = "ISENTO"
        txtChecaIE.Locked = True
    Else
        txtChecaIE.Text = ""
        txtChecaIE.Locked = False
    End If

End Sub

Private Sub cmdAcheEstado_Click()
Dim bytEstado As Byte

    If txtChecaIE2.Text = "" Then
        MsgBox "POR FAVOR, PREENCHA UMA INSCRIÇÃO ESTADUAL 2!", vbCritical, "PREENCHA UMA I.E.2"
        Exit Sub
    End If

    bytEstado = objChecaIE.AchaEstado(txtChecaIE2.Text)

    If bytEstado < 99 Then
        lblEstado.Caption = strUFe(bytEstado)
    Else
        lblEstado.Caption = "NÃO ENCONTRADO!"
    End If
    
End Sub

Private Sub Form_Load()
    cboEstado.ListIndex = 23
End Sub

Private Sub cmdChecaIE_Click()
    Dim blnChecaIE As Boolean
    Dim strUF As String
    
    'Não precisa desta checagem caso a combobox for "Style = 2 - Dropdown List".
'    If cboEstado.Text = "" Then
'        MsgBox "POR FAVOR, ESCOLHA UM ESTADO!", vbCritical, "ESCOLHA UM ESTADO"
'        Exit Sub
'    End If

    If txtChecaIE.Text = "" Then
        MsgBox "POR FAVOR, PREENCHA UMA INSCRIÇÃO ESTADUAL!", vbCritical, "PREENCHA UMA I.E."
        Exit Sub
    End If

    strUF = ConverteEstado(cboEstado.Text)
    blnChecaIE = objChecaIE.ChecaInscrE(strUF, txtChecaIE.Text)
    MsgBox ("A checagem foi: " & blnChecaIE)
End Sub

Private Function ConverteEstado(strUFESTADO As String) As String

    For bytNumero = 0 To UBound(strUFe)
        
        If strUFESTADO = strUFe(bytNumero) Then
            ConverteEstado = strUF(bytNumero)
        End If
        
    Next

End Function

Private Sub txtChecaIE_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtChecaIE2_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
