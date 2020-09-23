VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Water Ripple Test"
   ClientHeight    =   11370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   758
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   487
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicBackRender 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   0
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Polygon Lib "GDI32" (ByVal hdc As Long, lpPoint As PointApi, ByVal nCount As Long) As Long

Private Type PointApi
    X As Long
    Y As Long
End Type

Dim Started As Boolean
Dim CurT As Integer

Dim W_WIDTH As Integer
Dim W_HEIGHT As Integer

Dim RenderArr() As Long

Const PartDamp As Double = 1.5

Dim FrameNo As Long

Dim CurPX As Double
Dim CurPY As Double

Dim FWX As Integer
Dim FWY As Integer

Dim RNDW As Integer

Dim MultFact As Integer
Dim MultFactX As Integer
Dim MultFactY As Integer

Dim OffSetX As Integer
Dim OffSetY As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("x") Then
        'exit
        End
    End If
End Sub

Private Sub Form_Load()
    Randomize Timer
    
    CoolDraw = True
    Refraction = True
    
    Me.WindowState = 2
    Me.Show
    Me.FillStyle = vbSolid
    DoEvents

    RNDW = -1 'Increase to 0 or 1 to get better effect, though increases CPU usage lots!!!
    MultFact = 1
    DoEvents
    
    W_WIDTH = 80 * 2 ^ RNDW
    W_HEIGHT = 64 * 2 ^ RNDW
    
    ReDim RenderArr(0 To W_WIDTH - 1, 0 To W_HEIGHT - 1, 0 To 1) As Long
    
    DoEvents
    Started = True
    
    FWX = (Screen.Width / 15) / W_WIDTH
    FWY = (Screen.Height / 15) / W_HEIGHT
    
    CurPX = W_WIDTH / 2
    CurPY = W_HEIGHT / 2
    
    MultFactX = 3 / 2 ^ (RNDW - 1)
    MultFactY = 3.75 / 2 ^ (RNDW - 1)
        
    OffSetX = (Screen.Width / 15) / 2 - W_WIDTH * MultFactX
    OffSetY = (Screen.Height / 15) / 2 + 200
    
    Me.FontName = "Verdana"
    
    PicBackRender.Width = Screen.Width / 15
    PicBackRender.Height = Screen.Height / 15
    
    Dim X1 As Integer
    Dim Y1 As Integer
    Dim RSqrd As Double
    
    For X1 = 0 To W_WIDTH - 1
        For Y1 = 0 To W_HEIGHT - 1
            RSqrd = Sqr((CurPX - X1) * (CurPX - X1) + (CurPY - Y1) * (CurPY - Y1))
            
            If RSqrd <> 0 Then
                RenderArr(X1, Y1, 0) = (255 / W_WIDTH * RSqrd)
            End If
            
            Col = (255 / W_WIDTH * RSqrd)
            If Col > 255 Then
                Col = 255
            End If
            
            RenderArr(X1, Y1, 1) = RGB(0, Col, 255 - Col)
        Next
    Next
    
    Call Render
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Col As Long
    
    Col = GetPixel(PicBackRender.hdc, X, Y)
    If Col <> 0 And W_WIDTH <> 0 Then
        Dim RSqrd As Double
        Dim X1 As Integer, Y1 As Integer
        'On Grid
        
        CurPY = Int(Col / W_WIDTH)
        CurPX = Col - W_WIDTH * CurPY
        
        For X1 = 0 To W_WIDTH - 1
            For Y1 = 0 To W_HEIGHT - 1
                RSqrd = Sqr((CurPX - X1) * (CurPX - X1) + (CurPY - Y1) * (CurPY - Y1))
                
                If RSqrd <> 0 Then
                    RenderArr(X1, Y1, 0) = (255 / W_WIDTH * RSqrd)
                End If
                
                Col = (255 / W_WIDTH * RSqrd)
                If Col > 255 Then
                    Col = 255
                End If
                
                RenderArr(X1, Y1, 1) = RGB(0, Col, 255 - Col)
            Next
        Next
        
        CurT = 1 - CurT
        Call Render
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Make sure the program closes properly
    Started = False
    DoEvents
    End
End Sub

Sub Render()
    Dim point(0 To 5) As PointApi
    
    Dim X As Integer, Y As Integer
    Dim X1 As Integer, Y1 As Integer
    Dim X2 As Integer, Y2 As Integer
    
    'Do the main code to sort out the effect
    

    
    Me.Cls
    PicBackRender.Cls
    
    For X = 0 To W_WIDTH - 2
        X1 = X * MultFactX
        X2 = (X + 1) * MultFactX
        For Y = 0 To W_HEIGHT - 2
            Y1 = Y * MultFactX
            Y2 = (Y + 1) * MultFactX
                
            'Render to front buffer
            FrmMain.ForeColor = RenderArr(X, Y, 1)
            FrmMain.FillColor = RenderArr(X, Y, 1)
            
            point(0).X = OffSetX + X1 + Y1
            point(0).Y = OffSetY + Y1 - X1 - RenderArr(X, Y, 0) * MultFact
            point(1).X = OffSetX + X2 + Y1
            point(1).Y = OffSetY + Y1 - X2 - RenderArr(X + 1, Y, 0) * MultFact
            point(2).X = OffSetX + X2 + Y2
            point(2).Y = OffSetY + Y2 - X2 - RenderArr(X + 1, Y + 1, 0) * MultFact
            point(3).X = OffSetX + X1 + Y2
            point(3).Y = OffSetY + Y2 - X1 - RenderArr(X, Y + 1, 0) * MultFact
            point(4).X = point(0).X
            point(4).Y = point(0).Y
            
            Call Polygon(FrmMain.hdc, point(0), 4)
            
            
            'Render to the back buffer
            
            PicBackRender.FillColor = Y * W_WIDTH + X
            PicBackRender.ForeColor = Y * W_WIDTH + X
            
            point(0).X = OffSetX + X1 + Y1
            point(0).Y = OffSetY + Y1 - X1 - RenderArr(X, Y, 0) * MultFact
            point(1).X = OffSetX + X2 + Y1
            point(1).Y = OffSetY + Y1 - X2 - RenderArr(X + 1, Y, 0) * MultFact
            point(2).X = OffSetX + X2 + Y2
            point(2).Y = OffSetY + Y2 - X2 - RenderArr(X + 1, Y + 1, 0) * MultFact
            point(3).X = OffSetX + X1 + Y2
            point(3).Y = OffSetY + Y2 - X1 - RenderArr(X, Y + 1, 0) * MultFact
            point(4).X = point(0).X
            point(4).Y = point(0).Y
            
            Call Polygon(PicBackRender.hdc, point(0), 4)
                
        Next
    Next
    
    FrameNo = FrameNo + 1
    
    'Add text
    Me.CurrentX = 10
    Me.CurrentY = 20
    Me.FontBold = True
    Me.FontSize = 18
    Me.ForeColor = vbWhite
    Me.Print "Mouse Depression Demo V" & App.Major & "." & App.Minor & "." & App.Revision
            
    Me.FontSize = 12
    Me.CurrentX = 10
    Me.CurrentY = 50
    Me.Print "By Kevin Pfister"
            
    Me.FontBold = False
    
    Me.CurrentX = 10
    Me.CurrentY = 100
    Me.Print "Frame No: " & FrameNo
    
    Me.CurrentX = 10
    Me.CurrentY = 140
    Me.Print "X: " & CurPX
    
    Me.CurrentX = 10
    Me.CurrentY = 160
    Me.Print "Y: " & CurPY
    
    Me.CurrentX = 10
    Me.CurrentY = 200
    Me.Print "Press X to exit"
    
    Me.Refresh
    PicBackRender.Refresh
    
End Sub

Sub GetRgb(ByVal Color As Long, ByRef Red As Integer, ByRef Green As Integer, ByRef Blue As Integer)
    Dim LngColVal As Long
    LngColVal = Color And 255
    Red = LngColVal And 255
    LngColVal = Int(Color / 256)
    Green = LngColVal And 255
    LngColVal = Int(Color / 65536)
    Blue = LngColVal And 255
End Sub

