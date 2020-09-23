VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form eCredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "eCredit Sales"
   ClientHeight    =   5085
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6870
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6870
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   465
      Top             =   2880
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   4800
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Text            =   "© 1999 enix information systems"
            TextSave        =   "© 1999 enix information systems"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
      EndProperty
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   15
      Picture         =   "frmmain.frx":030A
      ScaleHeight     =   570
      ScaleWidth      =   6930
      TabIndex        =   1
      Top             =   0
      Width           =   6930
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   585
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   7435
      Arrange         =   2
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HotTracking     =   -1  'True
      PictureAlignment=   3
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6990
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4D61
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":563D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6319
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6BF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":74D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7DAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8A89
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9365
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img 
      Left            =   7110
      Top             =   1035
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9C41
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A165
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A689
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":ABAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B0D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B5F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":BB19
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":C03D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&Customer"
      Begin VB.Menu mnuCustomerReg 
         Caption         =   "&Customer Registraion"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuCustomerHis 
         Caption         =   "&Customer History"
      End
   End
   Begin VB.Menu mnuTran 
      Caption         =   "&Transactions"
      Begin VB.Menu mnuSales 
         Caption         =   "&Sales"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSettleInvoice 
         Caption         =   "Settle &Invoice"
      End
   End
   Begin VB.Menu mnuRept 
      Caption         =   "&Reports"
      Begin VB.Menu mnuAllOut 
         Caption         =   "&All Outstandings"
      End
      Begin VB.Menu mnuCustOut 
         Caption         =   "&Customer Outstanding"
      End
      Begin VB.Menu mnuDue 
         Caption         =   "&Print All Due Payments"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuCheckInvoice 
         Caption         =   "Check &Invoice"
      End
   End
   Begin VB.Menu mnuSys 
      Caption         =   "S&ystem"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "eCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xx As Double, yy As Double
Private Sub Form_Load()
    Dim itmX As ListItem
    Set itmX = ListView1.ListItems.Add(1, "Sal", "Sales", 1, 1)
    Set itmX = ListView1.ListItems.Add(2, "His", "Customer History", 4, 4)
    Set itmX = ListView1.ListItems.Add(3, "Reg", "Customer Registration", 5, 5)
    Set itmX = ListView1.ListItems.Add(4, "All", "Print All Outstandings", 3, 3)
    Set itmX = ListView1.ListItems.Add(5, "Cus", "Print A Customer's Outstandings", 2, 2)
    Set itmX = ListView1.ListItems.Add(6, "Due", "Print All Due Payments", 6, 6)
    Set itmX = ListView1.ListItems.Add(7, "Pay", "Settle Invoice", 7, 7)
    Set itmX = ListView1.ListItems.Add(8, "Chk", "Check Invoice", 8, 8)
    '--
    AddBitMapsToMenu
    'ListView1.Arrange
    OpenDB
    
End Sub
Private Sub ListView1_Click()
    LoadModules
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        LoadModules
    End If
End Sub
Private Sub LoadModules()
    On Error GoTo ExitThis
    If ListView1.HitTest(xx, yy).Key = "Reg" Then
        CustomerReg.Show
    End If
    If ListView1.HitTest(xx, yy).Key = "Sal" Then
        frmSales.Show
    End If
    If ListView1.HitTest(xx, yy).Key = "All" Then
        If Not DEnv.rsCommand2.State = adStateClosed Then DEnv.rsCommand2.Close
        DEnv.rsCommand2.Open "SELECT mstCust.Tel, mstCust.Name, mstCust.Contact, mstCust.Address, SUM(InvHeadder.Total - InvHeadder.Paid) AS netTotal, mstCust.CNum, InvHeadder.Settled FROM mstCust, InvHeadder WHERE mstCust.CNum = InvHeadder.CNum AND Settled=0 GROUP BY mstCust.Tel, mstCust.Name, mstCust.Contact, mstCust.Address, mstCust.CNum, InvHeadder.Settled ORDER BY mstCust.Name", db, adOpenStatic, adLockOptimistic
        DEnv.rsCommand2.Requery
        DoEvents
        allOut.Show
        DEnv.rsCommand2.Close
    End If
    If ListView1.HitTest(xx, yy).Key = "Cus" Then
        frmCrCheck.Show
    End If
    If ListView1.HitTest(xx, yy).Key = "His" Then
        frmCrHistory.Show
    End If
    If ListView1.HitTest(xx, yy).Key = "Pay" Then
        frmSettle.Show
    End If
    If ListView1.HitTest(xx, yy).Key = "Due" Then
        If Not DEnv.rsCommand11.State = adStateClosed Then DEnv.rsCommand11.Close
        DEnv.rsCommand11.Open "SELECT mstCust.Tel, mstCust.Name, mstCust.Contact, mstCust.Address, SUM(InvHeadder.Total - InvHeadder.Paid) AS netTotal, mstCust.CNum, InvHeadder.DueDate, InvHeadder.SalDate, InvHeadder.Settled FROM mstCust, InvHeadder WHERE mstCust.CNum = InvHeadder.CNum AND InvHeadder.Settled = 0 GROUP BY mstCust.Tel, mstCust.Name, mstCust.Contact, mstCust.Address, mstCust.CNum, InvHeadder.DueDate, InvHeadder.SalDate, InvHeadder.Settled ORDER BY mstCust.Name", db, adOpenStatic, adLockOptimistic
        DEnv.rsCommand11.Filter = "DueDate < #" & Format(Now, "Medium Date") & "#"
        DEnv.rsCommand11.Requery
        DoEvents
        allDue.Show
        DEnv.rsCommand11.Close
    End If
    If ListView1.HitTest(xx, yy).Key = "Chk" Then
        frmCheck.Show
    End If
ExitThis:
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    xx = x
    yy = y
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuAllOut_Click()
    If Not DEnv.rsCommand2.State = adStateClosed Then DEnv.rsCommand2.Close
    DEnv.rsCommand2.Open "SELECT mstCust.Tel, mstCust.Name, mstCust.Contact, mstCust.Address, SUM(InvHeadder.Total - InvHeadder.Paid) AS netTotal, mstCust.CNum, InvHeadder.Settled FROM mstCust, InvHeadder WHERE mstCust.CNum = InvHeadder.CNum AND Settled=0 GROUP BY mstCust.Tel, mstCust.Name, mstCust.Contact, mstCust.Address, mstCust.CNum, InvHeadder.Settled ORDER BY mstCust.Name", db, adOpenStatic, adLockOptimistic
    DEnv.rsCommand2.Requery
    DoEvents
    allOut.Show
    DEnv.rsCommand2.Close
End Sub

Private Sub mnuCheckInvoice_Click()
    frmCheck.Show
End Sub

Private Sub mnuCustomerHis_Click()
    frmCrHistory.Show
End Sub

Private Sub mnuCustomerReg_Click()
    CustomerReg.Show
End Sub

Private Sub mnuCustOut_Click()
    frmCrCheck.Show
End Sub

Private Sub mnuDue_Click()
    If Not DEnv.rsCommand11.State = adStateClosed Then DEnv.rsCommand11.Close
    DEnv.rsCommand11.Open "SELECT mstCust.Tel, mstCust.Name, mstCust.Contact, mstCust.Address, SUM(InvHeadder.Total - InvHeadder.Paid) AS netTotal, mstCust.CNum, InvHeadder.DueDate, InvHeadder.SalDate, InvHeadder.Settled FROM mstCust, InvHeadder WHERE mstCust.CNum = InvHeadder.CNum AND InvHeadder.Settled = 0 GROUP BY mstCust.Tel, mstCust.Name, mstCust.Contact, mstCust.Address, mstCust.CNum, InvHeadder.DueDate, InvHeadder.SalDate, InvHeadder.Settled ORDER BY mstCust.Name", db, adOpenStatic, adLockOptimistic
    DEnv.rsCommand11.Filter = "DueDate < #" & Format(Now, "Medium Date") & "#"
    DEnv.rsCommand11.Requery
    DoEvents
    allDue.Show
    DEnv.rsCommand11.Close
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuSales_Click()
    frmSales.Show
End Sub

Private Sub mnuSettleInvoice_Click()
    frmSettle.Show
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Timer1_Timer()
    StatusBar1.Panels(2).Text = Now
End Sub
Public Sub AddBitMapsToMenu()
  
  'add bitmaps to the menus
  
  Dim i%
  Dim hMenu, hSubMenu, menuID, x
  hMenu = GetMenu(hwnd)
  hSubMenu = GetSubMenu(hMenu, 0)
  
  menuID = GetMenuItemID(hSubMenu, 0)
  x = SetMenuItemBitmaps(hMenu, menuID, 0, img.ListImages(3).Picture, 0&)
  
  menuID = GetMenuItemID(hSubMenu, 1)
  x = SetMenuItemBitmaps(hMenu, menuID, 1, img.ListImages(2).Picture, 0&)
  '---
  hMenu = GetMenu(hwnd)
  hSubMenu = GetSubMenu(hMenu, 1)
  
  menuID = GetMenuItemID(hSubMenu, 0)
  x = SetMenuItemBitmaps(hMenu, menuID, 0, img.ListImages(1).Picture, 0&)
  
  menuID = GetMenuItemID(hSubMenu, 1)
  x = SetMenuItemBitmaps(hMenu, menuID, 1, img.ListImages(7).Picture, 0&)
  '----
  hMenu = GetMenu(hwnd)
  hSubMenu = GetSubMenu(hMenu, 2)
  
  menuID = GetMenuItemID(hSubMenu, 0)
  x = SetMenuItemBitmaps(hMenu, menuID, 0, img.ListImages(4).Picture, 0&)
  
  menuID = GetMenuItemID(hSubMenu, 1)
  x = SetMenuItemBitmaps(hMenu, menuID, 1, img.ListImages(5).Picture, 0&)
  
  menuID = GetMenuItemID(hSubMenu, 2)
  x = SetMenuItemBitmaps(hMenu, menuID, 2, img.ListImages(6).Picture, 0&)
  
  menuID = GetMenuItemID(hSubMenu, 3)
  x = SetMenuItemBitmaps(hMenu, menuID, 3, img.ListImages(8).Picture, 0&)
  Debug.Print x
  '--
  
End Sub
