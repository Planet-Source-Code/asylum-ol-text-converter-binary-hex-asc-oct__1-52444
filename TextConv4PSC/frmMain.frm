VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Text converter 1.0 (build 1008) -- By: Asylum"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import from file"
      Height          =   195
      Left            =   600
      TabIndex        =   22
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   5520
      Width           =   7815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   255
      Left            =   5400
      TabIndex        =   19
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert!"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   4920
      Width           =   5415
   End
   Begin VB.Frame fraOF 
      Caption         =   "Output format: (Encoding only)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3960
      TabIndex        =   11
      Top             =   3000
      Width           =   3855
      Begin VB.OptionButton optofHEX 
         Caption         =   "HEX"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2895
      End
      Begin VB.OptionButton optofASC 
         Caption         =   "ASC"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton optofOCT 
         Caption         =   "OCT"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton optofBIN 
         Caption         =   "BINARY"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2535
      End
   End
   Begin VB.Frame fraSTF 
      Caption         =   "Starting text format: (Decoding only)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   3855
      Begin MSComDlg.CommonDialog cdCD 
         Left            =   2280
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.OptionButton optstfBIN 
         Caption         =   "BINARY"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2535
      End
      Begin VB.OptionButton optstfOCT 
         Caption         =   "OCT"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton optstfASC 
         Caption         =   "ASC"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton optstfHEX 
         Caption         =   "HEX"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conversion type:"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   7815
      Begin VB.OptionButton optDec 
         Caption         =   "Decode"
         Height          =   195
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optEnc 
         Caption         =   "Encode"
         Height          =   195
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   7815
   End
   Begin VB.Label Label5 
      Caption         =   "Output:"
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   5280
      Width           =   7095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "When encoding, everything must START in ""Text"" format (e.g. Hello world!)"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   4680
      Width           =   7815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "When decoding, everything will be converted to ""Text"" format (e.g. Hello world!)"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   4440
      Width           =   7815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7800
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Input:"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   7815
   End
   Begin VB.Label Label1 
      Caption         =   "The return carrige (VbCrLf) will be replaced by a space so it doesn't show up in the encoded text..."
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MyWidth As Long = 7965
Const MyHeight As Long = 10065

Dim i As Integer, j As Integer '// my main for loop vars

Dim Convert As clsConvert '// the heart and soul of this proggy
Dim File As clsFile '// last minute addition so you can import large bodies of text and convert that


Private Sub cmdConvert_Click()
Dim Check As String

    If txtInput.Text = "" Then Exit Sub
    txtInput.Text = Replace(txtInput.Text, vbCrLf, " ")
    
With Convert
    If optEnc.Value = True Then '//Encode !
        If optofBIN.Value = True Then '// string to binary
            txtOutput.Text = .StrToBin(txtInput.Text)
        ElseIf optofHEX.Value = True Then '//string to hex
            txtOutput.Text = .StrToHex(txtInput.Text)
        ElseIf optofASC.Value = True Then '//string to asc
            txtOutput.Text = .StrToAsc(txtInput.Text)
        ElseIf optofOCT.Value = True Then '//string to oct
            txtOutput.Text = .StrToOct(txtInput.Text)
        End If
    ElseIf optDec.Value = True Then
        If optstfHEX.Value = True Then '//hex to string
            Check = .HexToStr(txtInput.Text)
            If Check = vbNullString Then
                .Error_
            Else
                txtOutput.Text = Check
            End If
        ElseIf optstfBIN.Value = True Then
            Check = .BinToStr(txtInput.Text)
            If Check = vbNullString Then
                .Error_
            Else
                txtOutput.Text = Check
            End If
        ElseIf optstfASC.Value = True Then
            Check = .AscToStr(txtInput.Text)
            If Check = vbNullString Then
                .Error_
            Else
                txtOutput.Text = Check
            End If
        End If
    End If
End With
End Sub

Private Sub cmdImport_Click()
cdCD.ShowOpen
If cdCD.FileName <> "" Then
    txtInput.Text = File.Import(cdCD.FileName)
    cdCD.FileName = ""
End If
End Sub

Private Sub Form_Load()
    Set Convert = New clsConvert '// Meow. :P
    Set File = New clsFile '// Bark. :P
    optEnc.Value = True
    Call optEnc_Click
    optofBIN.Value = True
    MsgBox "Please vote for this code and don't be lazy" & vbCrLf & "-Asylum", vbInformation + vbOKOnly
End Sub

Private Sub Form_Resize()
'/*
'This will force the form to stay the size I set at
'design time, I know that I can set the forms property to
'"non-resizable", but I want the icon at the top left to
'be shown :P
'*/
Select Case WindowState
    Case 0
    '/* WindowState 0 is our friend >_< */
    Case 1
        Exit Sub
    Case 2
        WindowState = 0
End Select
    Width = MyWidth
    Height = MyHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Convert = Nothing '// memory leaks = bad :(
    Set File = Nothing
End Sub

Private Sub optDec_Click()
    fraOF.Enabled = False
    fraSTF.Enabled = True
End Sub

Private Sub optEnc_Click()
    fraOF.Enabled = True
    fraSTF.Enabled = False
End Sub
