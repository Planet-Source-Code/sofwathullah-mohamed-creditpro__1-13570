VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOption 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   -90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3210
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   5662
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmOption.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "UpDown1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Picture1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   285
         Left            =   4335
         TabIndex        =   12
         Top             =   1125
         Width           =   360
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   4035
         TabIndex        =   10
         Top             =   2205
         Width           =   675
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2790
         TabIndex        =   9
         Top             =   2220
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse..."
         Height          =   345
         Left            =   3690
         TabIndex        =   7
         Top             =   1605
         Width           =   1020
      End
      Begin VB.PictureBox Picture1 
         Height          =   285
         Left            =   3675
         ScaleHeight     =   225
         ScaleWidth      =   525
         TabIndex        =   5
         Top             =   1125
         Width           =   585
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   4515
         TabIndex        =   1
         Top             =   675
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "UpDown1"
         BuddyDispid     =   196617
         OrigLeft        =   3615
         OrigTop         =   675
         OrigRight       =   3810
         OrigBottom      =   975
         Max             =   31
         Min             =   1
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "###"
         Height          =   270
         Left            =   3525
         TabIndex        =   11
         Top             =   2220
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Invoice Prefix and Suffix"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   2220
         Width           =   2115
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Main Form Back Image"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1650
         Width           =   1980
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Grid Seprator Color"
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   1162
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Default Invoice Settlement Time (days)"
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   675
         Width           =   3390
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   315
         Left            =   4020
         TabIndex        =   2
         Top             =   675
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UpDown1_Change()
    Label2.Caption = UpDown1.Value
End Sub
