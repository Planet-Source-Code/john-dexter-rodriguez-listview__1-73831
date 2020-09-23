VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcustom 
   Caption         =   "PICK A CATEGORY"
   ClientHeight    =   8805
   ClientLeft      =   1260
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   14985
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   8454143
      ImageWidth      =   200
      ImageHeight     =   200
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcustom.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcustom.frx":94C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcustom.frx":FD36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcustom.frx":163DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcustom.frx":1C2E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcustom.frx":23BF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcustom.frx":2A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcustom.frx":32F0F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lvproduct 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   14843
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483644
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmcustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub LoadForm(ByRef srcForm As Form)
    srcForm.Show
   ' srcForm.WindowState = vbMaximized
    srcForm.SetFocus
    
End Sub

Private Sub Form_Load()
With Lvproduct
        Set .SmallIcons = ImageList1
        Set .Icons = ImageList1
        
        'For Hardware and Software
        .ListItems.Add , "frmHardwareSoftware", "View Hardware and Software", 1, 1
        
        'View all Custom Notebooks
        .ListItems.Add , "frmNotebook", "View all Custom Notebooks", 2, 2
        
        'View all Custom Desktop
        .ListItems.Add , "frmDesktop", "View All Custom Desktop", 3, 3
        
        'View all Custom Workstation
        .ListItems.Add , "frmWorkstation", "View all Custom Workstation", 4, 4
       
        
        'View all Intel Rack Servers
        .ListItems.Add , "frmintelrackserver", "View all Intel Rack Servers", 5, 5
        
        'View all Storage Servers
        .ListItems.Add , "frmstorageserver", "View all Storage Servers", 6, 6
        
        'View Surveillance Solutions
         .ListItems.Add , "frmSurveillance", "View Surveillance Solutions", 7, 7
        
        'View all Pedestial Server
        .ListItems.Add , "frmpedestialServer", "View all Pedestial Server", 8, 8
        
    End With
End Sub




Private Sub Form_Resize()
    On Error Resume Next
    Lvproduct.Width = ScaleWidth
    Lvproduct.Height = ScaleHeight
  
End Sub



Private Sub Lvproduct_DblClick()

Select Case Lvproduct.SelectedItem.Key
    
        'For Hardware and Software
        Case "frmHardwareSoftware"
        MsgBox "You selected " & Lvproduct.SelectedItem.Key, vbInformation, "MESSAGE"
        'View all Customer Notebooks
       Case "frmNotebook":
        MsgBox "You selected " & Lvproduct.SelectedItem.Key, vbInformation, "MESSAGE"
        'View all Custom Desktop
        Case "frmDesktop"
         MsgBox "You selected " & Lvproduct.SelectedItem.Key, vbInformation, "MESSAGE"
        'View all Custom Workstation
        Case "frmWorkstation"
         MsgBox "You selected " & Lvproduct.SelectedItem.Key, vbInformation, "MESSAGE"
        'View all Intel Rack Server
        Case "frmintelrackserver"
         MsgBox "You selected " & Lvproduct.SelectedItem.Key, vbInformation, "MESSAGE"
        'View all Storage Server
        Case "frmstorageserver"
         MsgBox "You selected " & Lvproduct.SelectedItem.Key, vbInformation, "MESSAGE"
        'View Surveillance Solutions
        Case "frmSurveillance"
         MsgBox "You selected " & Lvproduct.SelectedItem.Key, vbInformation, "MESSAGE"
        'View all Pedestial Server
        Case "frmpedestialServer"
         MsgBox "You selected " & Lvproduct.SelectedItem.Key, vbInformation, "MESSAGE"
  End Select
End Sub
