VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flowchart Assistant 2004"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1260
      TabIndex        =   3
      Top             =   1785
      Width           =   990
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "UPX"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1290
      TabIndex        =   4
      Top             =   1110
      Width           =   390
   End
   Begin VB.Label Label2 
      Caption         =   $"FrmAbout.frx":000C
      Height          =   1185
      Left            =   285
      TabIndex        =   2
      Top             =   510
      Width           =   3045
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FlowChart Assistant 2004"
      BeginProperty Font 
         Name            =   "Tiranti Solid LET"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8000&
      Height          =   405
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   3270
   End
   Begin VB.Label Label1 
      Caption         =   "FlowChart Assistant 2004"
      BeginProperty Font 
         Name            =   "Tiranti Solid LET"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   60
      Width           =   3270
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Label3_Click()
   MsgBox "http://upx.sourceforge.net"
End Sub
