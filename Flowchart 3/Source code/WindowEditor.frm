VERSION 5.00
Begin VB.Form WindowEditor 
   Caption         =   "Window Editor"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   Picture         =   "WindowEditor.frx":0000
   ScaleHeight     =   339
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   506
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   30
      TabIndex        =   3
      Top             =   2925
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   15
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2550
      Width           =   885
   End
   Begin VB.PictureBox PicContainer 
      BackColor       =   &H8000000C&
      Height          =   5025
      Left            =   945
      ScaleHeight     =   4965
      ScaleWidth      =   6570
      TabIndex        =   0
      Top             =   30
      Width           =   6630
      Begin VB.PictureBox PicFrm 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         DrawWidth       =   2
         Height          =   4155
         Left            =   15
         ScaleHeight     =   277
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   323
         TabIndex        =   1
         Top             =   15
         Width           =   4845
         Begin VB.PictureBox PicRes 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   105
            Index           =   0
            Left            =   2865
            ScaleHeight     =   9.333
            ScaleMode       =   0  'User
            ScaleWidth      =   6
            TabIndex        =   15
            Top             =   2580
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.VScrollBar VS 
            Height          =   195
            Index           =   0
            Left            =   1650
            TabIndex        =   14
            Top             =   510
            Visible         =   0   'False
            Width           =   165
         End
         Begin VB.HScrollBar HS 
            Height          =   165
            Index           =   0
            Left            =   1230
            TabIndex        =   13
            Top             =   525
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.ListBox Lb 
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.ComboBox Cb 
            Height          =   315
            Index           =   0
            Left            =   390
            TabIndex        =   11
            Text            =   "Combo2"
            Top             =   465
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Option1"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   10
            Top             =   525
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.CheckBox Chk 
            Caption         =   "Check1"
            Height          =   225
            Index           =   0
            Left            =   1905
            TabIndex        =   9
            Top             =   195
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "Command1"
            Height          =   195
            Index           =   0
            Left            =   1530
            TabIndex        =   8
            Top             =   210
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Frame Fr 
            Caption         =   "Frame1"
            Height          =   210
            Index           =   0
            Left            =   1200
            TabIndex        =   7
            Top             =   180
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.TextBox Tb 
            Height          =   285
            Index           =   0
            Left            =   855
            TabIndex        =   6
            Text            =   "Text2"
            Top             =   135
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.PictureBox Pic 
            Height          =   225
            Index           =   0
            Left            =   180
            ScaleHeight     =   165
            ScaleWidth      =   180
            TabIndex        =   4
            Top             =   150
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            DrawMode        =   6  'Mask Pen Not
            Height          =   195
            Left            =   2610
            Top             =   195
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label Lbl 
            Caption         =   "Label1"
            Height          =   165
            Index           =   0
            Left            =   540
            TabIndex        =   5
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "WindowEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Under Development. Check For Updates Or Program It Yourself
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Const HorRes = 0
Const DiagRes = 1
Const VertRes = 2

Dim iIndex%, selL&, selT&
Dim SmX&, SmY&, OL&, OT&
Enum CountBy
   ByRows = 1
   ByColums = 0
End Enum

Private Sub Form_Activate()
    Dim Style&
    'Hehe this is what you've been looking of. Changes the style of the picbox to a form!
    Style = GetWindowLong(PicFrm.hwnd, -16)
    SetWindowLong PicFrm.hwnd, -16, Style Or &H40000 Or &H800000 Or &H400000
    PicFrm.Width = PicFrm.Width + 15
    For i = 1 To 7
      Load PicRes(i)
    Next
    PicRes(0).MousePointer = vbSizeWE
    PicRes(1).MousePointer = vbSizeWE
    PicRes(2).MousePointer = vbSizeNS
    PicRes(3).MousePointer = vbSizeNS
    PicRes(4).MousePointer = vbSizeNWSE
    PicRes(5).MousePointer = vbSizeNESW
    PicRes(6).MousePointer = vbSizeNESW
    PicRes(7).MousePointer = vbSizeNWSE
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  iIndex = DrawButtons(Me, X, Y, 25, 25, 2, 6, 4, 3, ByRows, 2, 2)
End Sub

Function DrawButtons(DrawDest As Object, X As Single, Y As Single, ButWidth&, ButHeight&, Cols&, Rows&, SpaceH&, SpaceV&, IndexCountType As CountBy, Optional BeginOffcetX&, Optional BeginOffcetY&) As Integer
  Dim XX&, YY&, Col&, Row&, hIndex&, vIndex&
  
  Col = (X - BeginOffcetX) \ (ButWidth + SpaceH)
  XX = Col * (ButWidth + SpaceH) + BeginOffcetX
  
  Row = (Y - BeginOffcetY) \ (ButHeight + SpaceV)
  YY = Row * (ButHeight + SpaceV) + BeginOffcetY
  
  If IndexCountType = 0 Then
     DrawButtons = Col * Rows + Row
  Else
     DrawButtons = Col + Row * Cols
  End If
  
  If Row < Rows And Col < Cols Then
    DrawDest.Cls
    DrawDest.Line (XX, YY)-(XX + ButWidth, YY), &H606060
    DrawDest.Line (XX, YY)-(XX, YY + ButHeight), &H606060
    DrawDest.Line (XX, YY + ButHeight)-(XX + ButWidth, YY + ButHeight), &HFFFFFF
    DrawDest.Line (XX + ButWidth, YY)-(XX + ButWidth, YY + ButHeight), &HFFFFFF
  End If
End Function

Private Sub Form_Resize()
PicContainer.Width = Me.Width / 15 - 72
PicContainer.Height = Me.Height / 15 - 31
End Sub

Sub ResizeHandles(Obj As Object)
    PicRes(0).Move Obj.Left - 10, Obj.Top + Obj.Height / 2 - 4
    PicRes(1).Move Obj.Left + Obj.Width, Obj.Top + Obj.Height / 2 - 4
    PicRes(2).Move Obj.Left + Obj.Width / 2 - 4, Obj.Top - 8
    PicRes(3).Move Obj.Left + Obj.Width / 2 - 4, Obj.Top + Obj.Height
    PicRes(4).Move Obj.Left - 8, Obj.Top - 8
    PicRes(5).Move Obj.Left + Obj.Width, Obj.Top - 8
    PicRes(6).Move Obj.Left - 8, Obj.Top + Obj.Height
    PicRes(7).Move Obj.Left + Obj.Width, Obj.Top + Obj.Height
    For i = 0 To 7
       PicRes(i).Visible = True
    Next
End Sub

Private Sub Pic_Click(Index As Integer)
ResizeHandles Pic(Index)
End Sub

Private Sub PicFrm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   selL = X
   selT = Y
   Shape1.Width = 0
   Shape1.Height = 0
   Shape1.Visible = True
End Sub

Private Sub PicFrm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Shape1.Move Min(X, selL), Min(Y, selT), Abs(X - selL), Abs(Y - selT)
End Sub
Function Min(ByVal a1 As Long, ByVal a2 As Long) As Long
   If a1 < a2 Then
      Min = a1
   Else
      Min = a2
   End If
End Function
Private Sub PicFrm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Obj As Object
  Shape1.Visible = False
  If iIndex > 0 Then
  Select Case iIndex
     Case 1: Set Obj = Pic
     Case 2: Set Obj = Lbl
     Case 3: Set Obj = Tb
     Case 4: Set Obj = Fr
     Case 5: Set Obj = Cmd
     Case 6: Set Obj = Chk
     Case 7: Set Obj = Opt
     Case 8: Set Obj = Cb
     Case 9: Set Obj = Lb
     Case 10: Set Obj = HS
     Case 11: Set Obj = VS
  End Select
  On Error Resume Next
  Obj(0).Move Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height
  Obj(0).Visible = True
  iIndex = 0
  End If
End Sub

Private Sub PicRes_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  SmX& = X: SmY& = Y
  OL = PicRes(0).Left: OT = PicRes(0).Top
  For i = 0 To 7
      PicRes(i).Visible = False
  Next
End Sub

Private Sub PicRes_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim L&, T&, W&, H&
  If Button <> 0 Then
     On Error Resume Next
     L = Pic(0).Left
     T = Pic(0).Top
     W = Pic(0).Width
     H = Pic(0).Height
     PicFrm.Cls
     Select Case Index
        Case 2
              Pic(0).Visible = False
              PicFrm.Line (L, T + Y + SmY - OT)-(L + W, T + H), , B
     End Select
  End If
End Sub
