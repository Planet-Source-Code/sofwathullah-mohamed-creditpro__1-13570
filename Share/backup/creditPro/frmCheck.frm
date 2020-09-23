VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check Invoice..."
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   Icon            =   "frmCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   90
      Width           =   1875
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   450
      Left            =   2490
      TabIndex        =   3
      Top             =   540
      Width           =   1185
      Caption         =   "Check"
      PicturePosition =   196613
      Size            =   "2090;794"
      Picture         =   "frmCheck.frx":030A
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton3 
      Height          =   450
      Left            =   1230
      TabIndex        =   2
      Top             =   540
      Width           =   1185
      Caption         =   "Close"
      PicturePosition =   196613
      Size            =   "2090;794"
      Picture         =   "frmCheck.frx":135C
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Invoice Number:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   90
      Width           =   1440
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
    If Not DEnv.rsCommand9.State = adStateClosed Then DEnv.rsCommand9.Close
    DEnv.rsCommand9.Open "SELECT mstCust.Address, mstCust.Contact, mstCust.Name, mstCust.Tel, InvHeadder.InvNo, InvHeadder.DueDate, InvHeadder.Paid, InvHeadder.SalDate, InvHeadder.Settled, InvHeadder.Total, mstCust.CNum FROM mstCust, InvHeadder WHERE mstCust.CNum = InvHeadder.CNum", db, adOpenStatic, adLockOptimistic
    DEnv.rsCommand9.Requery
    DEnv.rsCommand9.Filter = "InvNo='" & Text1.Text & "'"
    DoEvents
    invDetails.Show
    DEnv.rsCommand9.Close
End Sub

Private Sub CommandButton3_Click()
    Unload Me
End Sub

