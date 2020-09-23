VERSION 5.00
Begin VB.Form FrmPreview 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Print Preview"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   FillColor       =   &H00E0E0E0&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmPreview.frx":0000
      Left            =   735
      List            =   "FrmPreview.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   45
      Width           =   1620
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   300
      Left            =   30
      TabIndex        =   1
      Top             =   45
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   7575
      Left            =   30
      ScaleHeight     =   501
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   552
      TabIndex        =   0
      Top             =   420
      Width           =   8340
   End
End
Attribute VB_Name = "FrmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Pi As Currency = 3.14159265358979
Const LeftSide As Byte = 1
Const UpSide As Byte = 2
Const RightSide As Byte = 3
Const DownSide As Byte = 4

Sub DrawData(Dest As Object, OffsetX&, OffsetY&)
Dim C%, W&, H&, x&, y&, Tx$, ti$, X1&, X2&, Y1&, Y2&, L1&, L2&, T1&, T2&, W1&, W2&, H1&, H2&
Dim QSS(3) As POINTAPI
Dim QSS2(5) As POINTAPI
Dim T As Currency
Dest.ScaleMode = 3
'Handle Lines
For C = 1 To FlowChart.FlowShape1.Count - 1
      If ConLne(C, 0) <> 0 Then
         L1 = FlowChart.FlowShape1(C).Left
         W1 = FlowChart.FlowShape1(C).Width
         T1 = FlowChart.FlowShape1(C).Top
         H1 = FlowChart.FlowShape1(C).Height
         
         L2 = FlowChart.FlowShape1(ConLne(C, 0)).Left
         W2 = FlowChart.FlowShape1(ConLne(C, 0)).Width
         T2 = FlowChart.FlowShape1(ConLne(C, 0)).Top
         H2 = FlowChart.FlowShape1(ConLne(C, 0)).Height
         
         If FlowChart.FlowShape1(C).Title <> "" Or FlowChart.FlowShape1(C).Shape <> 0 Then
           Select Case SuitableSide(L1 + W1 / 2, T1 + H1 / 2, W1, H1, L2 + W2 / 2, T2 + H2 / 2)
              Case LeftSide:    X1 = L1:            Y1 = T1 + H1 / 2
              Case RightSide:   X1 = L1 + W1:       Y1 = T1 + H1 / 2
              Case UpSide:      X1 = L1 + W1 / 2:   Y1 = T1
              Case DownSide:    X1 = L1 + W1 / 2:   Y1 = T1 + H1
           End Select
         Else
           X1 = L1 + W1 / 2
           Y1 = T1 + H1 / 2
         End If
         If FlowChart.FlowShape1(ConLne(C, 0)).Title <> "" Or FlowChart.FlowShape1(ConLne(C, 0)).Shape <> 0 Then
           Select Case SuitableSide(L2 + W2 / 2, T2 + H2 / 2, W2, H2, L1 + W1 / 2, T1 + H1 / 2)
              Case LeftSide:    X2 = L2:            Y2 = T2 + H2 / 2
              Case RightSide:   X2 = L2 + W2:       Y2 = T2 + H2 / 2
              Case UpSide:      X2 = L2 + W2 / 2:   Y2 = T2
              Case DownSide:    X2 = L2 + W2 / 2:   Y2 = T2 + H2
           End Select
         Else
           X2 = L2 + W2 / 2
           Y2 = T2 + H2 / 2
         End If
         Dest.Line (X1 + OffsetX, Y1 + OffsetY)-(X2 + OffsetX, Y2 + OffsetY)
         Arrow X1 + OffsetX, Y1 + OffsetY, X2 + OffsetX, Y2 + OffsetY, Dest
      End If
      If ConLne(C, 1) <> 0 Then
         L1 = FlowChart.FlowShape1(C).Left
         W1 = FlowChart.FlowShape1(C).Width
         T1 = FlowChart.FlowShape1(C).Top
         H1 = FlowChart.FlowShape1(C).Height
         
         L2 = FlowChart.FlowShape1(ConLne(C, 1)).Left
         W2 = FlowChart.FlowShape1(ConLne(C, 1)).Width
         T2 = FlowChart.FlowShape1(ConLne(C, 1)).Top
         H2 = FlowChart.FlowShape1(ConLne(C, 1)).Height
         
         If FlowChart.FlowShape1(C).Title <> "" And FlowChart.FlowShape1(C).Shape <> 0 Then
           Select Case SuitableSide(L1 + W1 / 2, T1 + H1 / 2, W1, H1, L2 + W2 / 2, T2 + H2 / 2)
              Case LeftSide:    X1 = L1:            Y1 = T1 + H1 / 2
              Case RightSide:   X1 = L1 + W1:       Y1 = T1 + H1 / 2
              Case UpSide:      X1 = L1 + W1 / 2:   Y1 = T1
              Case DownSide:    X1 = L1 + W1 / 2:   Y1 = T1 + H1
           End Select
         Else
           X1 = L1 + W1 / 2
           Y1 = T1 + H1 / 2
         End If
         If FlowChart.FlowShape1(ConLne(C, 1)).Title <> "" And FlowChart.FlowShape1(ConLne(C, 1)).Shape <> 0 Then
           Select Case SuitableSide(L2 + W2 / 2, T2 + H2 / 2, W2, H2, L1 + W1 / 2, T1 + H1 / 2)
              Case LeftSide:    X2 = L2:            Y2 = T2 + H2 / 2
              Case RightSide:   X2 = L2 + W2:       Y2 = T2 + H2 / 2
              Case UpSide:      X2 = L2 + W2 / 2:   Y2 = T2
              Case DownSide:    X2 = L2 + W2 / 2:   Y2 = T2 + H2
           End Select
         Else
           X2 = L2 + W2 / 2
           Y2 = T2 + H2 / 2
         End If
         Dest.Line (X1 + OffsetX, Y1 + OffsetY)-(X2 + OffsetX, Y2 + OffsetY)
         Arrow X1 + OffsetX, Y1 + OffsetY, X2 + OffsetX, Y2 + OffsetY, Dest
         
         Dest.Print " True"
      End If
Next

'Handle Shapes
For C = 1 To FlowChart.FlowShape1.Count - 1
      W = FlowChart.FlowShape1(C).Width
      H = FlowChart.FlowShape1(C).Height
      x = FlowChart.FlowShape1(C).Left + OffsetX
      y = FlowChart.FlowShape1(C).Top + OffsetY
      Tx = FlowChart.FlowShape1(C).Caption
      ti = FlowChart.FlowShape1(C).Title
      
      Select Case FlowChart.FlowShape1(C).Shape
         Case 0
            'Don't Draw connecting poles
            If FlowChart.FlowShape1(C).Title <> "" Then
               Dest.Circle (W / 2 - 1 + x, H / 2 - 1 + y), W / 2 - 1, , , , (H / (W + 2))
            End If
         Case 1, 5
            QSS(0).x = x: QSS(0).y = y
            QSS(1).x = W + x: QSS(1).y = y
            QSS(2).x = W + x: QSS(2).y = H + y
            QSS(3).x = x: QSS(3).y = H + y
            Polygon Dest.hdc, QSS(0), 4
         Case 2
            QSS(0).x = H / 2 + x: QSS(0).y = y
            QSS(1).x = W + x: QSS(1).y = y
            QSS(2).x = W - H / 2 + x: QSS(2).y = H + y
            QSS(3).x = x: QSS(3).y = H + y
            Polygon Dest.hdc, QSS(0), 4
         Case 3
            QSS(0).x = W / 2 + x: QSS(0).y = y
            QSS(1).x = W + x: QSS(1).y = H / 2 + y
            QSS(2).x = W / 2 + x: QSS(2).y = H - 1 + y
            QSS(3).x = x: QSS(3).y = H / 2 + y
            Polygon Dest.hdc, QSS(0), 4
         Case 4
            QSS2(0).x = H / 2 + x: QSS2(0).y = y
            QSS2(1).x = W - H / 2 + x: QSS2(1).y = y
            QSS2(2).x = W + x: QSS2(2).y = H / 2 + y
            QSS2(3).x = W - H / 2 + x: QSS2(3).y = H + y
            QSS2(4).x = H / 3 + x: QSS2(4).y = H + y
            QSS2(5).x = x: QSS2(5).y = H / 2 + y
            Polygon Dest.hdc, QSS2(0), 6
        End Select
        Select Case FlowChart.FlowShape1(C).Shape
          Case 0:
             'Show title only
             Dest.CurrentX = x + (W - Dest.TextWidth(ti)) / 2  'Center horisontally
             Dest.CurrentY = y + (H - Dest.TextHeight(ti)) / 2 'Center vertically
             Dest.Print ti
          Case 1, 3, 4, 5
             'Show text only
             Dest.CurrentX = x + (W - Dest.TextWidth(Tx)) / 2 'Center horisontally
             Dest.CurrentY = y + (H - Dest.TextHeight(Tx)) / 2 'Center vertically
             Dest.Print Tx
          Case 2
             'Show title and text
             Dest.CurrentX = x + (W - Dest.TextWidth(ti)) / 2  'Center horisontally
             Dest.CurrentY = y + (H - Dest.TextHeight(Tx) - Dest.TextHeight(ti)) / 2 'Center vertically
             Dest.Print ti
             Dest.CurrentX = x + (W - Dest.TextWidth(Tx)) / 2  'Center horisontally
             Dest.Print Tx
        End Select
Next

End Sub

Private Sub Command1_Click()
   Printer.Print ""
   StretchBlt Printer.hdc, 0, 0, Picture1.ScaleWidth * (Combo1.ListIndex + 1), Picture1.ScaleHeight * (Combo1.ListIndex + 1), Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, &HCC0020
   Printer.EndDoc
End Sub

Private Sub Form_Load()
  Picture1.AutoRedraw = True
 ' Picture1.FillStyle = vbSolid
 ' Picture1.FillColor = vbWhite
 ' Picture1.FillStyle = vbSolid
'  Picture1.ForeColor = vbBlue
  DrawData Picture1, 0, 0
  Picture1.Refresh
  Combo1.ListIndex = 0
End Sub

Sub Arrow(X1&, Y1&, X2&, Y2&, Dest As Object)
   Dim T As Currency
   If X1 - X2 = 0 Then
      If Y2 < Y1 Then
         T = Pi / 2
      Else
         T = 3 * Pi / 2
      End If
   Else
      T = Atn((Y1 - Y2) / (X1 - X2))
   End If
   If X2 > X1 Then
       Dest.Line ((X1 + X2) / 2, (Y1 + Y2) / 2)-((X1 + X2) / 2 - Cos(T + Pi / 4) * 10, (Y1 + Y2) / 2 - Sin(T + Pi / 4) * 10)
       Dest.Line ((X1 + X2) / 2, (Y1 + Y2) / 2)-((X1 + X2) / 2 - Cos(T - Pi / 4) * 10, (Y1 + Y2) / 2 - Sin(T - Pi / 4) * 10)
   Else
       Dest.Line ((X1 + X2) / 2, (Y1 + Y2) / 2)-((X1 + X2) / 2 + Cos(T + Pi / 4) * 10, (Y1 + Y2) / 2 + Sin(T + Pi / 4) * 10)
       Dest.Line ((X1 + X2) / 2, (Y1 + Y2) / 2)-((X1 + X2) / 2 + Cos(T - Pi / 4) * 10, (Y1 + Y2) / 2 + Sin(T - Pi / 4) * 10)
   End If
End Sub
