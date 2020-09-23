VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Sample Paint"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Save Picture"
      Height          =   405
      Left            =   210
      TabIndex        =   14
      Top             =   2880
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6150
      Left            =   1935
      ScaleHeight     =   6090
      ScaleWidth      =   7965
      TabIndex        =   8
      Top             =   90
      Width           =   8025
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Random Colors"
      Height          =   330
      Left            =   240
      TabIndex        =   7
      Top             =   2505
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   435
      Left            =   1050
      ScaleHeight     =   375
      ScaleWidth      =   540
      TabIndex        =   6
      Top             =   2040
      Width           =   600
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   270
      Top             =   4005
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   1
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Text            =   "200"
      Top             =   1545
      Width           =   720
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Text            =   "150"
      Top             =   1170
      Width           =   720
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Text            =   "10"
      Top             =   825
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   330
      Left            =   315
      TabIndex        =   2
      Top             =   2070
      Width           =   645
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "5"
      Top             =   450
      Width           =   720
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1095
      TabIndex        =   0
      Text            =   "3"
      Top             =   105
      Width           =   720
   End
   Begin VB.Label Label5 
      Caption         =   "Max Length"
      Height          =   240
      Left            =   105
      TabIndex        =   13
      Top             =   1605
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "Min Length"
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   1245
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "Max Branch"
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   870
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "Min Branch"
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   510
      Width           =   810
   End
   Begin VB.Label Label1 
      Caption         =   "Levels"
      Height          =   225
      Left            =   135
      TabIndex        =   9
      Top             =   150
      Width           =   705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PI As Double = 3.14159265358979

Dim col As Long, rcol As Byte

Private Sub SpinWeb(Levels As Integer, ByVal X As Integer, ByVal Y As Integer, lBranch As Integer, hBranch As Integer, lLow As Integer, lHigh As Integer, color As Long)
Randomize
If Levels <> 0 Then
    Dim i As Integer, Branches As Integer
    Branches = Int(Rnd() * (hBranch - lBranch) + lBranch)
    For i = 1 To Branches
        Dim l As Integer, a As Integer, px As Integer, py As Integer
        l = Int(Rnd() * (lHigh - lLow) + lLow)
        a = 360 / Branches * i
        px = X + ccos(a) * l
        py = Y + csin(a) * l
        Picture2.Line (X, Y)-(px, py), color
        If rcol = 1 Then
            SpinWeb Levels - 1, px, py, lBranch, hBranch, lLow, lHigh, RGB(Int(Rnd() * 256), Int(Rnd() * 256), Int(Rnd() * 256))
        Else
            SpinWeb Levels - 1, px, py, lBranch, hBranch, lLow, lHigh, col
        End If
    Next i
End If
End Sub

Private Function ccos(ByVal angle As Integer) As Double
ccos = Cos(angle * PI / 180)
End Function

Private Function csin(ByVal angle As Integer) As Double
csin = Sin(angle * PI / 180)
End Function

Private Sub Check1_Click()
rcol = Check1.Value
End Sub

Private Sub Command1_Click()
Picture2.Cls
End Sub

Private Sub Command2_Click()
On Error GoTo 1
cd.Filter = "JPEG|*.jpg"
cd.ShowSave
SavePicture Picture2.Image, cd.FileName
1:
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'DrawTree Text1.Text, x, y, Text2.Text, vbBlue

End Sub

Private Sub Form_Resize()
Picture2.Width = Me.ScaleWidth - 75 - Picture2.Left
Picture2.Height = Me.ScaleHeight - 75 - Picture2.Top
End Sub

Private Sub Picture1_Click()
On Error GoTo 1
cd.ShowColor
col = cd.color
Picture1.BackColor = col
1:
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    SpinWeb Text1.Text, X, Y, Text2.Text, Text3.Text, Text4.Text, Text5.Text, col
End If
End Sub
