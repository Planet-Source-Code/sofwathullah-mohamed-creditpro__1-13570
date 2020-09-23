VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSettle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settle Invoice"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   -30
      Width           =   4695
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   3270
         TabIndex        =   7
         Top             =   1065
         Width           =   1290
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   3270
         TabIndex        =   4
         Top             =   390
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36509
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1050
         TabIndex        =   2
         Top             =   390
         Width           =   1050
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label6"
         Height          =   315
         Left            =   1050
         TabIndex        =   12
         Top             =   1065
         Width           =   1050
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Balance :"
         Height          =   210
         Left            =   90
         TabIndex        =   11
         Top             =   1065
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amount Paid For Settlement"
         Height          =   195
         Left            =   2145
         TabIndex        =   6
         Top             =   750
         Width           =   2400
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter Invoice Number To Sellte"
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
         Left            =   75
         TabIndex        =   5
         Top             =   75
         Width           =   3060
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date:"
         Height          =   195
         Left            =   2595
         TabIndex        =   3
         Top             =   390
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Invoice #:"
         Height          =   195
         Left            =   75
         TabIndex        =   1
         Top             =   390
         Width           =   900
      End
   End
   Begin MSForms.CommandButton CommandButton3 
      Height          =   450
      Left            =   870
      TabIndex        =   10
      Top             =   1545
      Width           =   1185
      Caption         =   "Close"
      PicturePosition =   196613
      Size            =   "2090;794"
      Picture         =   "frmSettle.frx":030A
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   450
      Left            =   2130
      TabIndex        =   9
      Top             =   1545
      Width           =   1185
      Caption         =   "Check"
      PicturePosition =   196613
      Size            =   "2090;794"
      Picture         =   "frmSettle.frx":135C
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   450
      Left            =   3390
      TabIndex        =   8
      Top             =   1545
      Width           =   1185
      Caption         =   "Settle"
      PicturePosition =   196613
      Size            =   "2090;794"
      Picture         =   "frmSettle.frx":23AE
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4665
      Y1              =   1410
      Y2              =   1410
   End
End
Attribute VB_Name = "frmSettle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public invHRS As ADODB.Recordset
Public CusNumber As Integer
Private Sub CommandButton1_Click()
    If Not Text2.Text = Empty Then
        If Not CDbl(Text2.Text) > CDbl(Label6.Caption) Then
            With invHRS
                !Paid = CDbl(!Paid) + CDbl(Text2.Text)
                If CDbl(Label6) - CDbl(Text2.Text) = 0 Then
                    !Settled = True
                End If
                .Update
                Unload Me
            End With
        Else
            MsgBox "Oops! Invalid Value", vbInformation
        End If
    End If
End Sub

Private Sub CommandButton2_Click()
    If Not DEnv.rsCommand9.State = adStateClosed Then DEnv.rsCommand9.Close
    DEnv.rsCommand9.Open "SELECT mstCust.Address, mstCust.Contact, mstCust.Name, mstCust.Tel, InvHeadder.InvNo, InvHeadder.DueDate, InvHeadder.Paid, InvHeadder.SalDate, InvHeadder.Settled, InvHeadder.Total, mstCust.CNum FROM mstCust, InvHeadder WHERE mstCust.CNum = InvHeadder.CNum", db, adOpenStatic, adLockOptimistic
    DEnv.rsCommand9.Requery
    DEnv.rsCommand9.Filter = "(CNum=" & CusNumber & ") AND InvNo='" & Text1.Text & "'"
    DoEvents
    invDetails.Show
    DEnv.rsCommand9.Close
End Sub

Private Sub CommandButton3_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Text1.SetFocus
End Sub
Private Sub Form_Load()
    Set invHRS = New ADODB.Recordset
    invHRS.Open "SELECT * From InvHeadder", db, adOpenStatic, adLockOptimistic
    CommandButton1.Enabled = False
    CommandButton2.Enabled = False
    Text2.Enabled = False
    Label6.Caption = Empty
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
          invHRS.Filter = "InvNo='" & (Text1.Text) & "'"
          If Not invHRS.EOF And invHRS.RecordCount > 0 Then
            CommandButton1.Enabled = True
            CommandButton2.Enabled = True
            Text2.Enabled = True
            Label6.Caption = Format((CDbl(invHRS!Total) - CDbl(invHRS!Paid)), "###,#0.00")
            DTPicker1.Value = invHRS!SalDate
            CusNumber = invHRS!CNum
          Else
            CommandButton1.Enabled = False
            CommandButton2.Enabled = False
            Text2.Enabled = False
            Label6.Caption = Empty
            MsgBox "Invoice Not Found!", vbExclamation
          End If
    End If
End Sub
