VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form CustomerReg 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Database"
   ClientHeight    =   4545
   ClientLeft      =   750
   ClientTop       =   1410
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustReg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7500
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
      Height          =   3300
      Left            =   0
      ScaleHeight     =   3300
      ScaleWidth      =   7515
      TabIndex        =   9
      Top             =   555
      Width           =   7515
      Begin MSComctlLib.ListView ListView1 
         Height          =   2790
         Left            =   135
         TabIndex        =   18
         Top             =   390
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   4921
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   3
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Customers"
            Object.Width           =   3635
         EndProperty
         Picture         =   "frmCustReg.frx":030A
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   3930
         TabIndex        =   4
         Top             =   405
         Width           =   3390
      End
      Begin VB.TextBox Text2 
         Height          =   840
         Left            =   3930
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   3390
      End
      Begin VB.TextBox Text3 
         Height          =   345
         Left            =   3930
         TabIndex        =   6
         Top             =   1770
         Width           =   1245
      End
      Begin VB.TextBox Text4 
         Height          =   345
         Left            =   3915
         TabIndex        =   7
         Top             =   2220
         Width           =   3390
      End
      Begin VB.TextBox Text5 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   345
         Left            =   3915
         TabIndex        =   8
         Top             =   2655
         Width           =   1245
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Credit Details"
         Height          =   345
         Left            =   5295
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2655
         Width           =   1590
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Current Customer List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   17
         Top             =   105
         Width           =   2250
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name :"
         Height          =   195
         Left            =   2670
         TabIndex        =   16
         Top             =   405
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Address :"
         Height          =   195
         Left            =   2670
         TabIndex        =   15
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tel :"
         Height          =   195
         Left            =   2670
         TabIndex        =   14
         Top             =   1770
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contact :"
         Height          =   195
         Left            =   2655
         TabIndex        =   13
         Top             =   2220
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "O/Balance :"
         Height          =   195
         Left            =   2655
         TabIndex        =   12
         Top             =   2655
         Width           =   1020
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      Picture         =   "frmCustReg.frx":1D34
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   -15
      Width           =   7500
   End
   Begin MSForms.CommandButton CommandButton4 
      Height          =   420
      Left            =   2415
      TabIndex        =   3
      Top             =   4020
      Width           =   1155
      Caption         =   "Delete"
      PicturePosition =   327683
      Size            =   "2037;741"
      Picture         =   "frmCustReg.frx":7148
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton3 
      Height          =   420
      Left            =   3710
      TabIndex        =   10
      Top             =   4020
      Width           =   1155
      Caption         =   "Update"
      PicturePosition =   327683
      Size            =   "2037;741"
      Picture         =   "frmCustReg.frx":7A22
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   420
      Left            =   5005
      TabIndex        =   2
      Top             =   4020
      Width           =   1155
      Caption         =   "Add "
      PicturePosition =   327683
      Size            =   "2037;741"
      Picture         =   "frmCustReg.frx":8A74
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   420
      Left            =   6300
      TabIndex        =   1
      Top             =   4020
      Width           =   1155
      Caption         =   "Done"
      PicturePosition =   327683
      Size            =   "2037;741"
      Picture         =   "frmCustReg.frx":9AC6
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7500
      Y1              =   3870
      Y2              =   3870
   End
End
Attribute VB_Name = "CustomerReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CusRS As ADODB.Recordset
Dim LCusRS As ADODB.Recordset
Dim invHRS As ADODB.Recordset
Public CusNumber As Integer
Private Sub Command1_Click()
    'SELECT InvHeadder.Total - InvHeadder.Paid AS OB, InvHeadder.SalDate, InvHeadder.InvNo, InvHeadder.Settled, InvHeadder.CNum, InvHeadder.DueDate, mstCust.Name FROM InvHeadder, mstCust WHERE InvHeadder.CNum = mstCust.CNum AND (InvHeadder.Settled = 0)
    Dim sqlStatement As String
    If Not DEnv.rsCommand4.State = adStateClosed Then DEnv.rsCommand4.Close
    sqlStatement = "SELECT InvHeadder.Total - InvHeadder.Paid AS OB, InvHeadder.SalDate, InvHeadder.InvNo, InvHeadder.Settled, InvHeadder.CNum, InvHeadder.DueDate, mstCust.Name FROM InvHeadder, mstCust WHERE InvHeadder.CNum = mstCust.CNum AND (InvHeadder.Settled = 0)"
    DEnv.rsCommand4.Open sqlStatement, db, adOpenStatic, adLockOptimistic
    DEnv.rsCommand4.Filter = "(CNum=" & CusNumber & ")"
    DEnv.rsCommand4.Requery
    DoEvents
    crDetails.Show
    DEnv.rsCommand4.Close
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub
Private Sub CommandButton2_Click()
    
    
    'If Not CusRS.State = adStateClosed Then CusRS.Close
    
    Set CusRS = New ADODB.Recordset
    CusRS.Open "SELECT * FROM mstCust", db, adOpenStatic, adLockOptimistic
    Set Text1.DataSource = CusRS
    Text1.DataField = "Name"
    Set Text2.DataSource = CusRS
    Text2.DataField = "Address"
    Set Text3.DataSource = CusRS
    Text3.DataField = "Tel"
    Set Text4.DataSource = CusRS
    Text4.DataField = "Contact"
    Text5.DataField = Empty
    Text5.Text = Empty
    
    Label6.Visible = False
    Text5.Visible = False
    Command1.Visible = False
    
    CusRS.AddNew
    
    CommandButton2.Enabled = False
    Text1.SetFocus
    
      
End Sub
Private Sub CommandButton3_Click()
    On Error GoTo errFucks
    If Text1.Text = Empty Then
        MsgBox "Oops! Invalid Data or Data Missing", vbCritical
        Exit Sub
    End If
    
    CusRS.Update
    CusRS.Close
    
    
    Set CusRS = New ADODB.Recordset
    CusRS.Open "SELECT * FROM mstCust", db, adOpenStatic, adLockOptimistic
    
    Set Text1.DataSource = CusRS
    Text1.DataField = "Name"
    Set Text2.DataSource = CusRS
    Text2.DataField = "Address"
    Set Text3.DataSource = CusRS
    Text3.DataField = "Tel"
    Set Text4.DataSource = CusRS
    Text4.DataField = "Contact"
    Set Text5.DataSource = CusRS
    'Text5.DataField = "OB"
    Text5.Enabled = True
    If CommandButton2.Enabled = False Then CommandButton2.Enabled = True
    Label6.Visible = True
    Text5.Visible = True
    Command1.Visible = True
    DoList
    Exit Sub
errFucks:
    MsgBox "oops! Unexpacted Error, contact vendor."
End Sub
Private Sub CommandButton4_Click()
    Dim sureTha As Integer
    If Not CusRS.EOF Or CusRS.BOF Then
        sureTha = MsgBox("Sure?", vbQuestion + vbYesNo)
        If sureTha = 6 Then
            If Not DEnv.rsCommand5.State = adStateClosed Then DEnv.rsCommand5.Close
            sqlStatement = "SELECT Name, CNum FROM mstCust"
            DEnv.rsCommand5.Open sqlStatement, db, adOpenStatic, adLockOptimistic
            DEnv.rsCommand5.Filter = "(Name='" & Text1.Text & "')"
            DEnv.rsCommand5.Requery
            MsgBox DEnv.rsCommand5.RecordCount
            If DEnv.rsCommand5.RecordCount > 0 Then
                MsgBox "Cannot Delete Record. The Customer Has Unsettled Payments", vbCritical
            Else
                CusRS.Delete '//delete on confirmation
            End If
            DEnv.rsCommand5.Close
        End If
    End If
    DoList
End Sub
Private Sub Form_Load()
    
    Set CusRS = New ADODB.Recordset
    CusRS.Open "SELECT * FROM mstCust", db, adOpenStatic, adLockOptimistic
    
    Set Text1.DataSource = CusRS
    Text1.DataField = "Name"
    Set Text2.DataSource = CusRS
    Text2.DataField = "Address"
    Set Text3.DataSource = CusRS
    Text3.DataField = "Tel"
    Set Text4.DataSource = CusRS
    Text4.DataField = "Contact"
    Set Text5.DataSource = CusRS
    'Text5.DataField = "OB"
    
    DoList
    FromListUpdateRecord
End Sub
Private Sub DoList()
    Set LCusRS = New ADODB.Recordset
    LCusRS.Open "SELECT * FROM mstCust", db, adOpenStatic, adLockOptimistic
    
    If Not LCusRS.BOF Then LCusRS.MoveFirst
    ListView1.ListItems.Clear
    Do While Not LCusRS.EOF
        ListView1.ListItems.Add , , LCusRS!Name '//Add names to list
        LCusRS.MoveNext
    Loop
    If Not LCusRS.EOF Then LCusRS.MoveFirst
    ListView1.Refresh
End Sub
Private Sub ListView1_Click()
    FromListUpdateRecord
End Sub
Private Sub FromListUpdateRecord()
    On Error GoTo ExiThis
    If Not CusRS.BOF Then CusRS.MoveFirst
    If Not ListView1.SelectedItem.Text = Empty Then
        CusRS.Find "Name='" & Trim(ListView1.SelectedItem.Text) & "'"
        If Not CusRS.EOF Then
            CusNumber = CusRS!CNum
            Set invHRS = New ADODB.Recordset
            invHRS.Open "SELECT SUM(Total - Paid) AS OB, CNum From InvHeadder GROUP BY CNum HAVING (CNum = " & CusNumber & ")", db, adOpenStatic, adLockOptimistic
            If Not invHRS.EOF Then
                Text5.Text = Format(invHRS!OB, "###,##0.00")
            Else
                Text5.Text = "0.00"
            End If
        End If
    End If
ExiThis:
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    FromListUpdateRecord
End Sub
Private Sub Picture1_Click()
    Text4.SetFocus
    DoEvents
End Sub
