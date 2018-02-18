VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmValores 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   5340
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FrmValores.frx":0000
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   6105
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   1710
         Left            =   4200
         ScaleHeight     =   1650
         ScaleWidth      =   1620
         TabIndex        =   13
         Top             =   1680
         Width           =   1680
         Begin VB.Label LPre 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   60
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1440
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   1560
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   4365
         ScaleHeight     =   1035
         ScaleWidth      =   1035
         TabIndex        =   12
         Top             =   480
         Width           =   1095
         Begin VB.Image Image1 
            Height          =   1215
            Left            =   0
            Picture         =   "FrmValores.frx":0234
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   1095
         End
      End
      Begin VB.CommandButton BValor 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   10
         Left            =   1560
         TabIndex        =   11
         Top             =   4050
         Width           =   2475
      End
      Begin VB.CommandButton BValor 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   3
         Left            =   2880
         TabIndex        =   10
         Top             =   2865
         Width           =   1155
      End
      Begin VB.CommandButton BValor 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   2
         Left            =   1560
         TabIndex        =   9
         Top             =   2865
         Width           =   1155
      End
      Begin VB.CommandButton BValor 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   7
         Left            =   240
         TabIndex        =   8
         Top             =   465
         Width           =   1155
      End
      Begin VB.CommandButton BValor 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   2865
         Width           =   1155
      End
      Begin VB.CommandButton BValor 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   6
         Left            =   2880
         TabIndex        =   6
         Top             =   1665
         Width           =   1155
      End
      Begin VB.CommandButton BValor 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   5
         Left            =   1560
         TabIndex        =   5
         Top             =   1665
         Width           =   1155
      End
      Begin VB.CommandButton BValor 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   1665
         Width           =   1155
      End
      Begin VB.CommandButton BValor 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   9
         Left            =   2880
         TabIndex        =   3
         Top             =   465
         Width           =   1155
      End
      Begin VB.CommandButton BValor 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   8
         Left            =   1560
         TabIndex        =   2
         Top             =   465
         Width           =   1155
      End
      Begin VB.CommandButton BValor 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   4050
         Width           =   1170
      End
   End
End
Attribute VB_Name = "FrmValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pos As Integer
Private Sub BValor_GotFocus(Index As Integer)
If Index = 10 Then
    LPre = "00"
Else
    LPre = Index
End If
End Sub

Private Sub BValor_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
If KeyCode = 33 Then
    If pos < 10 Then pos = pos + 1
End If

If KeyCode = 34 Then
    If pos > 0 Then pos = pos - 1
End If

BValor(pos).SetFocus

End Sub

Private Sub BValor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 10 Then
    LPre = "00"
Else
    LPre = Index
End If
End Sub

Private Sub Form_Load()
Error_skin Me

pos = 0

Valor = 230
Call Aplicar_Transparencia(Me.hwnd, CByte(Valor))
End Sub

Private Sub Label1_Click()

End Sub

