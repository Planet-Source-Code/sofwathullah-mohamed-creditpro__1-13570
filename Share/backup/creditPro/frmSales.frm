VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales"
   ClientHeight    =   5985
   ClientLeft      =   750
   ClientTop       =   330
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7485
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3270
      ScaleHeight     =   480
      ScaleWidth      =   4200
      TabIndex        =   13
      Top             =   5490
      Width           =   4200
      Begin MSForms.CommandButton Command2 
         Height          =   435
         Left            =   1245
         TabIndex        =   15
         Top             =   15
         Width           =   1365
         Caption         =   " Cancel"
         PicturePosition =   327683
         Size            =   "2408;767"
         Picture         =   "frmSales.frx":030A
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton Command1 
         Height          =   435
         Left            =   2775
         TabIndex        =   14
         Top             =   15
         Width           =   1365
         Caption         =   " Done"
         PicturePosition =   327683
         Size            =   "2408;767"
         Picture         =   "frmSales.frx":135C
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.PictureBox Picture2 
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
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4815
      ScaleWidth      =   7575
      TabIndex        =   1
      Top             =   570
      Width           =   7575
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1125
         TabIndex        =   18
         Top             =   225
         Width           =   1785
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   2550
         Left            =   225
         TabIndex        =   17
         Top             =   1245
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   4498
         _Version        =   393216
         BackColor       =   16777215
         FixedCols       =   0
         BackColorBkg    =   16777215
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSDataListLib.DataCombo Combo1 
         Height          =   315
         Left            =   1965
         TabIndex        =   16
         Top             =   720
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "DataCombo1"
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1770
         TabIndex        =   12
         Top             =   4365
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Format          =   53608449
         CurrentDate     =   36476
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   3915
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   5580
         TabIndex        =   9
         Top             =   4365
         Width           =   1710
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   5505
         TabIndex        =   4
         Top             =   225
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         _Version        =   393216
         Format          =   53608449
         CurrentDate     =   36475
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Payment Due On:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   4365
         Width           =   1530
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amount Paid:"
         Height          =   195
         Left            =   4230
         TabIndex        =   8
         Top             =   4365
         Width           =   1155
      End
      Begin VB.Line Line2 
         X1              =   5610
         X2              =   7260
         Y1              =   4125
         Y2              =   4125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "xxxx.xxx MRf"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5745
         TabIndex        =   7
         Top             =   3915
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total For This Invoice:"
         Height          =   195
         Left            =   3735
         TabIndex        =   6
         Top             =   3915
         Width           =   1920
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer's Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sales Date:"
         Height          =   315
         Left            =   4245
         TabIndex        =   3
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Invoice #:"
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   225
         Width           =   975
      End
   End
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
      Height          =   630
      Left            =   -105
      Picture         =   "frmSales.frx":23AE
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   506
      TabIndex        =   0
      Top             =   -60
      Width           =   7590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7560
      Y1              =   5400
      Y2              =   5400
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CusRS As ADODB.Recordset
Dim invHRS As ADODB.Recordset
Dim invDRS As ADODB.Recordset
Dim SetHRS As ADODB.Recordset
Dim tRS As ADODB.Recordset
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    '// stop user from entering to combo, only let user choose
    '// by ignoring the pressed key
    Select Case KeyCode
        Case vbKeyReturn '// but if enter is pressed go to grid
            Grid.SetFocus
            Exit Sub
        Case vbKeyUp, vbKeyDown '// scroll through the list of customers
            Exit Sub
        Case vbKeyTab   '// tab is the standerd key so switch to grid
            Grid.Col = 0: Grid.Row = 1
            Grid_EnterCell
            Exit Sub
    End Select
    KeyCode = 0
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    '// set focus on the combo, before that initilize grid too(interface bug fix)
    Grid.Col = 0: Grid.Row = 1
    Grid_EnterCell
    DoEvents
    Combo1.SetFocus
    Text3.SetFocus
End Sub

Private Sub Form_Load()
    '// initilize form and do the setup
    Grid.Cols = 4
    Grid.Rows = 100
    Grid.Row = 0
    Grid.Col = 0: Grid.Text = "Qty"
    Grid.Col = 1: Grid.Text = "Description"
    Grid.Col = 2: Grid.Text = "Rate"
    Grid.Col = 3: Grid.Text = "Total"
    Grid.ColWidth(1) = 3900
    Grid.ColWidth(3) = 1000
    Text2.Text = Empty
    Text2.Visible = False
    Grid.Rows = 2
    Label8.Caption = "0.00"
    DTPicker1.Value = Now
    DTPicker2.Value = Now + 14 '// by default now we give 14 days for credit
    
    '// fill teh combo with customers names
    Set CusRS = New ADODB.Recordset
    CusRS.Open "Select * FROM mstCust", db, adOpenStatic, adLockOptimistic
    '-
    Set Combo1.RowSource = CusRS
    Combo1.ListField = "Name"
    Combo1.DataField = "Name"
    Set Combo1.DataSource = CusRS
    '-
    '// generate the next invoice number (last inv # + 1 is the trick)
    Dim newInvNo As Integer
    Set SetHRS = New ADODB.Recordset
    SetHRS.Open "SELECT InvNo FROM Settings", db, adOpenStatic, adLockOptimistic
    If Not SetHRS.EOF Then
        newInvNo = SetHRS!InvNo
    Else
        newInvNo = 1
    End If
    SetHRS.Close
    
    'text3.text = newInvNo '// display the new inv #
End Sub
Private Sub Grid_EnterCell()
    '// when click on cell
    Select Case Grid.Col
        Case 0, 1, 2
            With Text2
                .Move Grid.CellLeft + Grid.Left, _
                Grid.CellTop + Grid.Top, Grid.CellWidth - 25, Grid.CellHeight - 25
                .Text = Grid.Text
                If Len(.Text) > 0 Then
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End If
                .Visible = True
                .ZOrder 0
                If Grid.Row Mod 2 = 0 Then
                    Text2.BackColor = RGB(174, 245, 214) '// lets make the grid color diff, every other grid
                Else
                    Text2.BackColor = RGB(255, 255, 255)
                End If
                .SetFocus
            End With
    End Select
End Sub
Private Sub Grid_GotFocus()
    'Grid_EnterCell
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Grid_EnterCell
    End If
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim Qty As Double, Rate As Double, Total As Double
    Dim lr As Integer, lTotal As Double

    Select Case KeyCode
            Case vbKeyEscape
                '// when esc is pressed cancel and get out
                With Text2
                    .Text = Empty
                    .Visible = False
                End With
                Grid.SetFocus
            Case vbKeyLeft
                '// move left
                If Grid.Col = 0 Or Grid.Col = 1 Or Grid.Col = 2 And Text2.SelLength > 0 Then
                    With Text2
                        If Not .Text = Empty Then
                            Grid.Text = .Text
                        End If
                        .Visible = False
                        .Text = Empty
                    End With
                    If Grid.Col = 2 Then
                        Grid.Col = 1
                    ElseIf Grid.Col = 1 Then
                        Grid.Col = 0
                    Else
                        Grid.Col = 2
                    End If
                    Grid_EnterCell
                End If
            Case vbKeyRight
                '// move right
                If Grid.Col = 0 Or Grid.Col = 1 Or Grid.Col = 2 And Text2.SelLength > 0 Then
                    With Text2
                        If Not .Text = Empty Then
                            Grid.Text = .Text
                        End If
                        .Visible = False
                        .Text = Empty
                    End With
                    If Grid.Col = 2 Then
                        Grid.Col = 0
                    ElseIf Grid.Col = 1 Then
                        Grid.Col = 2
                    Else
                        Grid.Col = 1
                    End If
                    Grid_EnterCell
                End If
            Case vbKeyDown
                '// move down until last row, if last move to first
                With Text2
                    If Not .Text = Empty Then
                        Grid.Text = .Text
                    End If
                    .Visible = False
                    .Text = Empty
                End With
                If Not Grid.Row = Grid.Rows - 1 Then
                    Grid.Row = Grid.Row + 1
                    Grid_EnterCell
                Else
                    Grid.Row = 1
                    Grid_EnterCell
                End If
            Case vbKeyUp
                '// move up until first row -1, if first then move last
                With Text2
                    If Not .Text = Empty Then
                        Grid.Text = .Text
                    End If
                    .Visible = False
                    .Text = Empty
                End With
                If Not Grid.Row = 1 Then
                    Grid.Row = Grid.Row - 1
                    Grid_EnterCell
                Else
                    Grid.Row = Grid.Rows - 1
                    Grid_EnterCell
                End If
            Case vbKeyReturn
                '// when enter is pressed, move to next col
                With Text2
                    If Not .Text = Empty Then
                        Grid.Text = .Text
                    End If
                    .Visible = False
                    .Text = Empty
                End With
                Select Case Grid.Col
                    Case 0
                        Grid.Col = 1
                        Grid_EnterCell
                    Case 1
                        Grid.Col = 2
                        Grid_EnterCell
                    Case 2
                        'hmmm! this is tricky , but cool (naa! not at all)
                        
                        Grid.Col = 0: Qty = Val(Grid.Text)
                        Grid.Col = 2: Rate = Val(Grid.Text)
                        Total = Qty * Rate: Grid.Col = 3: Grid.Text = Format(Total, "###,###,##0.00")
                        DoTotals
                        If Not Grid.Row = Grid.Rows - 1 Then
                            If Not Grid.Text = Empty And CDbl(Grid.Text) > 0 Then
                                Grid.Row = Grid.Row + 1
                            End If
                            Grid.Col = 0
                            Grid_EnterCell
                        Else
                            '// we need to add a new row ey, baby
                            If Not Grid.Text = Empty And CDbl(Grid.Text) > 0 Then
                                Grid.Rows = Grid.Rows + 1
                                Grid.Row = Grid.Row + 1
                                Fancy
                            End If
                            Grid.Col = 0
                            Grid_EnterCell
                        End If
                End Select
            Case vbKeyHome
                If Not Grid.Col = 0 And Text2.SelLength > 0 Then
                    With Text2
                        If Not .Text = Empty Then
                            Grid.Text = .Text
                        End If
                        .Visible = False
                        .Text = Empty
                    End With
                    Grid.Col = 0
                    Grid_EnterCell
                End If
            Case vbKeyEnd
                If Not Grid.Col = 2 And Text2.SelLength > 0 Then
                    With Text2
                        If Not .Text = Empty Then
                            Grid.Text = .Text
                        End If
                        .Visible = False
                        .Text = Empty
                    End With
                    Grid.Col = 2
                    Grid_EnterCell
                End If
    End Select
End Sub
Public Sub Fancy()
    '// since this is the last row as we know
    '// so lets add one more(van mor)
    Dim CurrentCell As Integer
    With Grid
        If .Row Mod 2 = 0 Then
            '// trying to make this row diff col
            CurrentCell = .Col
            Dim r As Integer
            For r = 0 To 3
                .Col = r
                .CellBackColor = RGB(174, 245, 214)
            Next
            .Col = CurrentCell
        End If
    End With
End Sub
Private Sub DoTotals()
    '// get the total from all
    Dim CurrentCell As Integer
    Dim CurrentRow As Integer
    
    CurrentCell = Grid.Col
    CurrentRow = Grid.Row
    
    lTotal = 0
    Grid.Col = 3
    For r = 1 To Grid.Rows - 1
        Grid.Row = r
        If Not Grid.Text = Empty Then
            lTotal = lTotal + CDbl(Grid.Text)
        End If
    Next
    Label8.Caption = Format(lTotal, "###,###,##0.00")
    
    DoEvents
    
    Grid.Col = CurrentCell
    Grid.Row = CurrentRow
End Sub
Private Function FindCustNo(cName As String)
    '// find the cust no for a given cust name
    Set tRS = New ADODB.Recordset
    tRS.Open "Select * FROM mstCust WHERE Name='" & cName & "'", db, adOpenStatic, adLockOptimistic
    If tRS.RecordCount > 0 Then
        FindCustNo = tRS!CNum
    Else
        FindCustNo = Empty
    End If
    tRS.Close
End Function
Private Sub WriteHadder()
    '// write hadder data to db
    Set invHRS = New ADODB.Recordset
    invHRS.Open "SELECT * From InvHeadder", db, adOpenStatic, adLockOptimistic
    With invHRS
        .AddNew
        !CNum = FindCustNo(Combo1.Text)
        !InvNo = Text3.Text
        !SalDate = DTPicker1.Value
        !DueDate = DTPicker2.Value
        !Total = CDbl(Label8.Caption)
        !Paid = IIf(Val(Text1.Text) > 0, CDbl(Text1.Text), 0)
        !Settled = IIf(CDbl(Label8.Caption) = CDbl(Text1.Text), True, False)
        .Update
    End With
    Set SetHRS = New ADODB.Recordset
    SetHRS.Open "SELECT * FROM Settings", db, adOpenStatic, adLockOptimistic
    If SetHRS.EOF Or SetHRS.BOF Then SetHRS.AddNew
    SetHRS!InvNo = Val(Text3.Text) + 1
    SetHRS.Update
    SetHRS.Close
End Sub
Private Sub WriteDetails()
    '// update inv details from grid
    Dim r As Integer
    Set invDRS = New ADODB.Recordset
    invDRS.Open "SELECT * From InvDetails", db, adOpenStatic, adLockOptimistic
    
    For r = 1 To Grid.Rows - 1
        Grid.Row = r
        With invDRS
            .AddNew
            !InvNo = Text3.Text
            Grid.Col = 0: !Qty = IIf(Not Grid.Text = Empty, Val(Grid.Text), 1)
            Grid.Col = 1: !Desc = IIf(Grid.Text = Empty, "Misc.", Grid.Text)
            Grid.Col = 2
            If Not Trim(Grid.Text) = Empty Then
                !Rate = CDbl(Grid.Text)
            Else
                !Rate = 0
            End If
            Grid.Col = 3
            If Val(Grid.Text) > 0 Then
                !Total = CDbl(Grid.Text)
                .Update                 '// update only if total is > 0
            Else
                .Cancel
            End If
        End With
    Next
    
End Sub
Private Sub Command1_Click()
    '// check if vaild then alaka zoom! write the data to db
    Dim okayS As Boolean
    okayS = CheckValidInv
    If okayS = False Then
        MsgBox "Oops! Invoice Number Already Taken!", vbInformation
        Text3.SetFocus
        Text3.Text = Empty
        Exit Sub
    End If
    If Val(Text1.Text) < 1 Then Text1.Text = "0"
    If Val(Label8.Caption) > 0 And Val(Text3.Text) > 0 Then
        WriteHadder
        WriteDetails
        Unload Me
        Exit Sub
    End If
    MsgBox "Oops! Data Missing or Invalid", vbCritical
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim okayS As Boolean
        okayS = CheckValidInv
        If okayS = False Then
            MsgBox "Oops! Invoice Number Already Taken!", vbInformation
            Text3.SetFocus
            Text3.Text = Empty
        Else
            Grid.SetFocus
        End If
    End If
End Sub
Function CheckValidInv() As Boolean
        Set tRS = New ADODB.Recordset
        tRS.Open "SELECT * FROM InvHeadder WHERE InvNo ='" & Text3.Text & "'", db, adOpenStatic, adLockOptimistic
        If tRS.RecordCount > 0 Then
            CheckValidInv = False
        Else
            CheckValidInv = True
        End If
End Function
