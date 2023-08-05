VERSION 5.00
Begin VB.Form Exercise4 
   BackColor       =   &H80000006&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculator"
   ClientHeight    =   4920
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmdcos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Cmdcos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "n!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Cmdcos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ln"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Cmdcos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "x^2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Cmdcos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "x^3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Cmdop 
      BackColor       =   &H00FFFFC0&
      Caption         =   "x^y"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Cmdcos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "sin"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Cmdcos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "tan"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Cmdcos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "cos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H8000000D&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdc 
      BackColor       =   &H80000017&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton Cmdop 
      BackColor       =   &H8000000E&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Cmdop 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Cmdop 
      BackColor       =   &H8000000E&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Cmdop 
      BackColor       =   &H8000000E&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmde 
      BackColor       =   &H8000000D&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H8000000D&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H8000000D&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H8000000D&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H8000000D&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H8000000D&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H8000000D&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H8000000D&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H8000000D&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H8000000D&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H8000000D&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3120
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   240
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Text            =   "0"
      Top             =   360
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "<<<Go back"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   4440
      Width           =   1815
   End
End
Attribute VB_Name = "Exercise4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim no1 As Double, ans As Double, no2 As Double, NO3 As Double, cnt As Integer, cv As Integer, cnt1 As Integer
Dim op As String, op1 As String, flag As Integer, a As Double, x As Integer, nos As Integer


Private Declare Function ShellExecute Lib _
   "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
   Private SW_SHOWNORMAL

Private Sub cmd_Click(Index As Integer)
    If txt.Text = "0" Then
    txt.Text = ""
    End If
    If x = 0 Then
    txt.Text = ""
    x = x + 1
    End If
    txt.Text = txt.Text + cmd(Index).Caption
End Sub

Private Sub cmdc_Click()
    txt.Text = "0"
    cnt = 0
    cnt1 = 0
    j = 0
    x = 0
End Sub

Private Sub cmdCos_Click(Index As Integer)
   op1 = Cmdcos(Index).Caption
   nos = txt.Text
    Select Case op1
        Case "cos"
                If flag = 1 Then
                ans = Cos((3.14 / 180) * nos)
                Else
                ans = Cos(nos)
                End If
        Case "sin"
                If flag = 1 Then
                ans = Sin((3.14 / 180) * nos)
                Else
                ans = Sin(nos)
                End If
        Case "tan"
                If flag = 1 Then
                ans = Tan((3.14 / 180) * nos)
                Else
                ans = Tan(nos)
                End If
        Case "x^2"
                ans = nos ^ 2
        Case "x^3"
                ans = nos ^ 3
         Case "ln"
                If nos = 0 Then
                i = MsgBox("Math error", vbExclamation + vbOKOnly, "ERROR")
                Else
                ans = Log(nos)
                End If
         Case "n!"
                ans = 1
                no2 = nos
                For i = 1 To no2
                ans = ans * i
                Next
         Case "sqrt"
                ans = Sqr(nos)
       End Select
    txt.Text = ans
End Sub

Private Sub cmde_Click()
cnt1 = 0
    Select Case op
        Case "+"
                ans = txt.Text + no1
                If cnt = 0 Then
                  no1 = txt.Text
                    cnt = cnt + 1
                End If
                txt.Text = ans
        Case "-"
                If cnt > 0 Then
                  txt.Text = no1
                  no1 = ans
                End If
                 ans = no1 - txt.Text
                 no1 = txt.Text
                 txt.Text = ans
                cnt = cnt + 1
        Case "*"
                ans = no1 * txt.Text
                If cnt = 0 Then
                  no1 = txt.Text
                    cnt = cnt + 1
                End If
        txt.Text = ans
        Case "/"
                If cnt > 0 Then
                  txt.Text = no1
                  no1 = ans
                End If
                If txt.Text = "0" Then
                i = MsgBox("Divide by zero error", vbExclamation + vbOKOnly, "ERROR")
                Else
                ans = no1 / txt.Text
                 no1 = txt.Text
                txt.Text = ans
                cnt = cnt + 1
                End If
        Case "x^y"
                ans = no1 ^ txt.Text
                txt.Text = ans
    End Select
End Sub
Private Sub cmdop_Click(Index As Integer)
    x = 0
    If cnt1 > 0 Then
        Select Case op
            Case "+"
                    ans = no1 + txt.Text
                    txt.Text = ans
                    no1 = txt.Text
                    op = Cmdop(Index).Caption
            Case "-"
                    ans = no1 - txt.Text
                    txt.Text = ans
                    no1 = txt.Text
                    op = Cmdop(Index).Caption
            Case "/"
                    ans = no1 / txt.Text
                    txt.Text = ans
                    no1 = txt.Text
                    op = Cmdop(Index).Caption
            Case "*"
                    ans = no1 * txt.Text
                    txt.Text = ans
                    no1 = txt.Text
                    op = Cmdop(Index).Caption
            Case "x^y"
                    ans = no1 ^ txt.Text
                    txt.Text = ans
                    no1 = txt.Text
                    op = Cmdop(Index).Caption
        End Select
    Else
    no1 = Val(txt.Text)
    op = Cmdop(Index).Caption
    txt.Text = 0
    cnt1 = cnt1 + 1
    End If
End Sub


Private Sub Form_Load()
cnt = 0
j = 0
i = 1
flag = 1
For i = 0 To 2
Cmdcos(i).Enabled = True
Next
End Sub

Private Sub Label1_Click()
Exercise4.Hide
MainForm.Show
End Sub
