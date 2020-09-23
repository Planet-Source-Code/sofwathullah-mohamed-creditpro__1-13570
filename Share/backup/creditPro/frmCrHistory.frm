VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCrHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer History..."
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "frmCrHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   0
      Picture         =   "frmCrHistory.frx":030A
      ScaleHeight     =   1440
      ScaleWidth      =   5370
      TabIndex        =   2
      Top             =   0
      Width           =   5370
      Begin MSDataListLib.DataCombo Combo1 
         Height          =   315
         Left            =   2025
         TabIndex        =   3
         Top             =   555
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer's Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   555
         Width           =   1770
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Customer To View Credit History"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   4
         Top             =   105
         Width           =   4680
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4005
      TabIndex        =   1
      Top             =   1560
      Width           =   1305
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2610
      TabIndex        =   0
      Top             =   1560
      Width           =   1305
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   15
      X2              =   5370
      Y1              =   1455
      Y2              =   1455
   End
End
Attribute VB_Name = "frmCrHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CusRS As ADODB.Recordset
Private Sub Command1_Click()
    '// Load the data and genreate report
    Dim sqlStatement As String
    If Not DEnv.rsCommand7.State = adStateClosed Then DEnv.rsCommand7.Close
    sqlStatement = "SELECT Name, CNum FROM mstCust"
    DEnv.rsCommand7.Open sqlStatement, db, adOpenStatic, adLockOptimistic
    DEnv.rsCommand7.Filter = "(Name='" & Combo1.Text & "')"
    DEnv.rsCommand7.Requery
    DoEvents
    crHistory.Show
    DEnv.rsCommand7.Close
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    '// fill teh combo with customers names
    Set CusRS = New ADODB.Recordset
    CusRS.Open "Select * FROM mstCust", db, adOpenStatic, adLockOptimistic
    '-
    Set Combo1.RowSource = CusRS
    Combo1.ListField = "Name"
    Combo1.DataField = "Name"
    Set Combo1.DataSource = CusRS
End Sub
