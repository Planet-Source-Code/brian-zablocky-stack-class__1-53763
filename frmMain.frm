VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stack Class Tester"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   870
      Left            =   2655
      ScaleHeight     =   810
      ScaleWidth      =   900
      TabIndex        =   8
      Top             =   180
      Width           =   960
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         TabIndex        =   10
         Top             =   270
         Width           =   825
      End
      Begin VB.Label lblJunk 
         Alignment       =   2  'Center
         BackColor       =   &H00400040&
         Caption         =   "Count"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.TextBox txtPop 
      Height          =   285
      Left            =   4410
      Locked          =   -1  'True
      MaxLength       =   24
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2070
      Width           =   1455
   End
   Begin VB.TextBox txtRead 
      Height          =   285
      Left            =   2430
      Locked          =   -1  'True
      MaxLength       =   24
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1455
   End
   Begin VB.TextBox txtPush 
      Height          =   285
      Left            =   450
      MaxLength       =   24
      TabIndex        =   0
      Top             =   2070
      Width           =   1455
   End
   Begin VB.CommandButton cmdStack 
      Caption         =   "Read"
      Height          =   285
      Index           =   2
      Left            =   2430
      TabIndex        =   2
      Top             =   1620
      Width           =   1455
   End
   Begin VB.CommandButton cmdStack 
      Caption         =   "Pop"
      Height          =   285
      Index           =   1
      Left            =   4410
      TabIndex        =   3
      Top             =   2430
      Width           =   1455
   End
   Begin VB.CommandButton cmdStack 
      Caption         =   "Push"
      Height          =   285
      Index           =   0
      Left            =   450
      TabIndex        =   1
      Top             =   2430
      Width           =   1455
   End
   Begin VB.Frame frame 
      Caption         =   "Top Of Stack"
      Height          =   690
      Index           =   1
      Left            =   2175
      TabIndex        =   4
      Top             =   2160
      Width           =   1950
      Begin VB.TextBox txtStack 
         Height          =   285
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
         Width           =   1680
      End
   End
   Begin VB.Label lblJunk 
      Caption         =   $"frmMain.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   2
      Left            =   4140
      TabIndex        =   12
      Top             =   135
      Width           =   2040
   End
   Begin VB.Label lblJunk 
      Caption         =   $"frmMain.frx":00D1
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   1
      Left            =   135
      TabIndex        =   11
      Top             =   135
      Width           =   2040
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    

Private Sub cmdStack_Click(Index As Integer)
    
    Select Case Index
        Case 0
            If txtPush.Text = "" Then
                Beep
            Else
                Data.sPush txtPush.Text
            End If
        Case 1
            If Data.StackCount = 0 Then
                Beep
            Else
                txtPop.Text = Data.sPop
            End If
        Case 2
            If Data.StackCount = 0 Then
                Beep
            Else
                txtRead.Text = Data.sRead
            End If
    End Select
    
    lblCount.Caption = Trim(Str(Data.StackCount))
    txtStack.Text = Data.sRead
End Sub
