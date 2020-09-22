VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDetalhe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalhe"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox rtbTexto 
      Height          =   4035
      Left            =   1170
      TabIndex        =   7
      Top             =   1560
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7117
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDetalhe.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   7935
   End
   Begin VB.TextBox txtTask 
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   630
      Width           =   7935
   End
   Begin VB.TextBox txtPackage 
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   180
      Width           =   7935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text:"
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   1530
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   1110
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Task:"
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   660
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Package:"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   210
      Width           =   720
   End
End
Attribute VB_Name = "frmDetalhe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    rtbTexto.RightMargin = 45000
End Sub


Public Sub ShowModal(strPackage As String, strTask As String, _
                strProperty As String, strText As String, strSearchString As String)
    
    
    Dim lngPos As Long
    
    txtPackage = strPackage
    txtTask = strTask
    txtProperty = strProperty
    rtbTexto = strText
    
    lngPos = 0
    Do
        lngPos = rtbTexto.Find(strSearchString, lngPos + rtbTexto.SelLength)
        rtbTexto.SelColor = vbBlue
        If lngPos = -1 Then Exit Do
    Loop
    rtbTexto.SelStart = 0
    
    Show vbModal

End Sub

