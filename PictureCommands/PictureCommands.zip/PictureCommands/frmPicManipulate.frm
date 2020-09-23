VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmManipulate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mike's Picture Toolz"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10245
   Icon            =   "frmPicManipulate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog OpenPicture 
      Left            =   4920
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox ClearPicture 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   4800
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Save Picture!"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Load Picture"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rotate 45 Degrees"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Flip Vertical"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Flip Horizontal"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   5535
      Left            =   5160
      ScaleHeight     =   5475
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   5535
      Left            =   120
      Picture         =   "frmPicManipulate.frx":1CFA
      ScaleHeight     =   5475
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label frmPicManipulate 
      Caption         =   "By: Mike Canejo"
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "E-mail: Mike@dev-center.com for   Questions or Comments!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   12
      Top             =   5880
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Manipulated:"
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Loaded:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmManipulate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'By : Mike Canejo

'This project I put together will
'manipulate a picture by either flipping it
'horizontally, vertically or Rotates it 45 degrees
'
'AIM: TheLeadX or Mike3dd
'Email: Mike@dev-center.com

'     Enjoy!




'If you havn't Yet:

'  Goto: www.dev-center.com
'For the best VB Code DataBase
'       On the Web!







Option Explicit

Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020
Private Const Pi = 3.14159265359


Private Sub Command1_Click()
Dim px As Long, py As Long, retval As Long
    Picture2 = ClearPicture
    px& = Picture1.ScaleWidth: py& = Picture1.ScaleHeight
If Command1.Caption = "Flip Horizontal" Then
    Command1.Caption = "Un-Flip Horizontal"
    retval& = StretchBlt(Picture2.hDC, px&, 0, -px&, py&, Picture1.hDC, 0, 0, px&, py&, SRCCOPY): Exit Sub
  Else
    Command1.Caption = "Flip Horizontal"
    Picture2 = Picture1
End If
    Picture2.Refresh
    'flip horizontal
End Sub


Private Sub Command2_Click()
Dim px As Long, py As Long, retval As Long
    Picture2 = ClearPicture
    px& = Picture1.ScaleWidth: py& = Picture1.ScaleHeight
If Command2.Caption = "Flip Vertical" Then
    Command2.Caption = "Un-Flip Vertical"
    retval& = StretchBlt(Picture2.hDC, 0, py&, px&, -py&, Picture1.hDC, 0, 0, px&, py&, SRCCOPY): Exit Sub
  Else
    Command2.Caption = "Flip Vertical"
    Picture2 = Picture1
End If
    Picture2.Refresh
    'flip vertical
End Sub


Private Sub Command3_Click()
    Picture2 = ClearPicture
If Command3.Caption = "Rotate 45 Degrees" Then
    Command3.Caption = "Undo Rotate"
    RotatePicture Picture1, Picture2, 3.14 / 4: Exit Sub
 Else
    Command3.Caption = "Rotate 45 Degrees"
    Picture2 = Picture1
End If
    Picture2.Refresh
'rotate 45 degrees
End Sub


Private Sub Command4_Click()
    Dim x As Long
    x = MsgBox("Are you sure you would like to exit?", vbSystemModal + vbCritical + vbYesNo, "Exit:"): If x& = vbYes Then End
'Confirms exit incase u press exit by mistake
End Sub

Private Sub Command5_Click()
    Picture2 = ClearPicture
 'To "start over"
End Sub

Private Sub Command6_Click()
On Error GoTo Heaven
OpenPicture.DialogTitle = "Load Picture"
OpenPicture.Filter = "Picture(*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
OpenPicture.ShowOpen
If OpenPicture.FileName = "" Then Exit Sub
Picture1 = LoadPicture(OpenPicture.FileName)
Exit Sub
Heaven: MsgBox "There was an error loading the picture! " & vbCrLf & "Please make sure it's a valid picture file.", vbSystemModal + vbCritical, "Error:"
'Opens Picture
End Sub

Private Sub Command7_Click()
On Error GoTo Heaven
OpenPicture.DialogTitle = "Save Picture"
OpenPicture.Filter = "Picture(*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
OpenPicture.ShowSave
If OpenPicture.FileName = "" Then Exit Sub
SavePicture Picture2.Image, OpenPicture.FileName
Exit Sub
Heaven: MsgBox "There was an error saving the manipulated picture!", vbSystemModal + vbCritical, "Error:"
'Saves the manipulated picture
End Sub

Private Sub Form_Load()
    Picture1.ScaleMode = 3
    Picture2.ScaleMode = 3
End Sub

Private Function RotatePicture(pic1 As PictureBox, pic2 As PictureBox, ByVal theta!)
On Error Resume Next
    'Rotates the image in a picture box.
    'pic1 is the picture box with the bitmap
    '     to rotate
    'pic2 is the picture box to receive the
    '     rotated bitmap
    'theta is the angle of rotation
    
    Dim n As Integer, r As Integer
    Dim c1x As Integer, c1y As Integer
    Dim c2x As Integer, c2y As Integer
    Dim p1x As Integer, p1y As Integer
    Dim p2x As Integer, p2y As Integer
    Dim pic1hDC As Long, pic2hDC As Long, a As Single
    Dim c0 As Long, c1 As Long, c2 As Long, c3 As Long, xret As Long
    
    c1x = pic1.ScaleWidth \ 2
    c1y = pic1.ScaleHeight \ 2
    c2x = pic2.ScaleWidth \ 2
    c2y = pic2.ScaleHeight \ 2
    If c2x < c2y Then n = c2y Else n = c2x
    n = n - 1
    pic1hDC& = pic1.hDC
    pic2hDC& = pic2.hDC

    For p2x = 0 To n
        For p2y = 0 To n
            If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x)
            r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
            p1x = r * Cos(a + theta!)
            p1y = r * Sin(a + theta!)
            c0& = GetPixel(pic1hDC&, c1x + p1x, c1y + p1y)
            c1& = GetPixel(pic1hDC&, c1x - p1x, c1y - p1y)
            c2& = GetPixel(pic1hDC&, c1x + p1y, c1y - p1x)
            c3& = GetPixel(pic1hDC&, c1x - p1y, c1y + p1x)
            If c0& <> -1 Then xret& = SetPixel(pic2hDC&, c2x + p2x, c2y + p2y, c0&)
            If c1& <> -1 Then xret& = SetPixel(pic2hDC&, c2x - p2x, c2y - p2y, c1&)
            If c2& <> -1 Then xret& = SetPixel(pic2hDC&, c2x + p2y, c2y - p2x, c2&)
            If c3& <> -1 Then xret& = SetPixel(pic2hDC&, c2x - p2y, c2y + p2x, c3&)
        Next
    Next
End Function

