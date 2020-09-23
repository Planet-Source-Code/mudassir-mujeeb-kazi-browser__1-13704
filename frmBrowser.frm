VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowser 
   ClientHeight    =   5130
   ClientLeft      =   3060
   ClientTop       =   3000
   ClientWidth     =   8850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin VB.CommandButton ADD 
         Caption         =   "&ADD"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   15
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   1680
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton Cont 
         Caption         =   "&Continue"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   19
         Top             =   1080
         Width           =   3375
      End
      Begin VB.ComboBox SITE 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   360
         Left            =   240
         TabIndex        =   18
         Text            =   "MAIL"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   285
         Left            =   2400
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Authentication code Required"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   3615
      End
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   8880
      ExtentX         =   15663
      ExtentY         =   6376
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   7080
      Top             =   1500
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   8850
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   8850
      Begin VB.ComboBox INF 
         BackColor       =   &H80000002&
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmBrowser.frx":0000
         Left            =   7200
         List            =   "frmBrowser.frx":0002
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox ENT 
         BackColor       =   &H80000002&
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmBrowser.frx":0004
         Left            =   5400
         List            =   "frmBrowser.frx":0006
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox NEWS 
         BackColor       =   &H80000002&
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmBrowser.frx":0008
         Left            =   3600
         List            =   "frmBrowser.frx":000A
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox PHONE 
         BackColor       =   &H80000002&
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmBrowser.frx":000C
         Left            =   1800
         List            =   "frmBrowser.frx":000E
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox MAIL 
         BackColor       =   &H80000002&
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmBrowser.frx":0010
         Left            =   0
         List            =   "frmBrowser.frx":0012
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cboAddress 
         BackColor       =   &H80000002&
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   45
         TabIndex        =   2
         Top             =   780
         Width           =   3795
      End
      Begin VB.Label Label5 
         Caption         =   "INFORMATION"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   13
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "ENTERTAINMENT"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "NEWS PAPER"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "PHONE"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "MAIL "
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblAddress 
         Caption         =   "&Address:"
         Height          =   255
         Left            =   45
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   540
         Width           =   3075
      End
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   2670
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBrowser.frx":0014
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBrowser.frx":06A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBrowser.frx":0D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBrowser.frx":13CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBrowser.frx":1A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBrowser.frx":20EE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TMP As String
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean
Dim CONSTR As String
Dim db As ADODB.Connection
Dim RS1 As ADODB.Recordset
Dim RS2 As ADODB.Recordset
Dim RS3 As ADODB.Recordset
Dim RS4 As ADODB.Recordset
Dim RS5 As ADODB.Recordset
Private Sub ADD_Click()
 ADD.Enabled = False
 Frame1.Visible = True
End Sub
Private Sub Cont_Click()
 ADD.Enabled = True
 If Text2.Text = "NetTech" Then
  MAIL.Clear
  PHONE.Clear
  NEWS.Clear
  ENT.Clear
  INF.Clear
 If Text1.Text = "" Then
  MsgBox " Site Address Can't Be Blank ", vbCritical
 Else
 If SITE.Text = "MAIL" Then
  RS1.AddNew
  RS1.Fields("SITE") = Text1.Text
  RS1.Update
 End If
 
 If SITE.Text = "PHONE" Then
  RS2.AddNew
  RS2.Fields("SITE") = Text1.Text
  RS2.Update
 End If
 
 If SITE.Text = "NEWS PAPER" Then
  RS3.AddNew
  RS3.Fields("SITE") = Text1.Text
  RS3.Update
 End If
 
 If SITE.Text = "ENTERTAINMENT" Then
  RS4.AddNew
  RS4.Fields("SITE") = Text1.Text
  RS4.Update
 End If
 
 If SITE.Text = "INFORMATION" Then
  RS5.AddNew
  RS5.Fields("SITE") = Text1.Text
  RS5.Update
 End If
 End If
  Form_Load
  MsgBox " Please Wait .....", vbInformation, "Adding Site"
 Else
  MsgBox " Authentication Code Error", vbCritical
End If
  Frame1.Visible = False
  Text1.Text = ""
  Text2.Text = ""
End Sub
Private Sub ENT_Click()
  TMP = "WWW." & ENT.Text & ".COM"
  cboAddress.Text = TMP
  cboAddress_Click
End Sub
Private Sub Form_Load()
  On Error Resume Next
  Set db = New ADODB.Connection
  CONSTR = "Provider=" & "Microsoft.Jet.OLEDB." & "4.0;Data Source=" & App.Path & "\brow.mdb"
  db.Open CONSTR
  Set RS1 = New ADODB.Recordset
  Set RS2 = New ADODB.Recordset
  Set RS3 = New ADODB.Recordset
  Set RS4 = New ADODB.Recordset
  Set RS5 = New ADODB.Recordset
  RS1.Open "MAIL", db, adOpenStatic, adLockOptimistic
  RS2.Open "PHONE", db, adOpenStatic, adLockOptimistic
  RS3.Open "NEWS", db, adOpenStatic, adLockOptimistic
  RS4.Open "ENT", db, adOpenStatic, adLockOptimistic
  RS5.Open "INF", db, adOpenStatic, adLockOptimistic
   Me.Show
   tbToolBar.Refresh
   Form_Resize

    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15

    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If
   
   'MAIL SITES
    RS1.MoveFirst
    TOT1 = RS1.RecordCount
    For i = 1 To TOT1
       MAIL.AddItem RS1.Fields("SITE")
       RS1.MoveNext
   Next
   
   'PHONE SITES
    RS2.MoveFirst
    TOT2 = RS2.RecordCount
    For i = 1 To TOT2
       PHONE.AddItem RS2.Fields("SITE")
       RS2.MoveNext
   Next
    
   'NEWS PAPER SITES
    RS3.MoveFirst
    TOT3 = RS3.RecordCount
    For i = 1 To TOT3
       NEWS.AddItem RS3.Fields("SITE")
       RS3.MoveNext
   Next
   
   'ENTERTAINMENT SITES
    RS4.MoveFirst
    TOT4 = RS4.RecordCount
    For i = 1 To TOT4
       ENT.AddItem RS4.Fields("SITE")
       RS4.MoveNext
   Next
   
   'INFORMATION SITES
    RS5.MoveFirst
    TOT5 = RS5.RecordCount
    For i = 1 To TOT5
       INF.AddItem RS5.Fields("SITE")
       RS5.MoveNext
   Next
    
    'ADDING SITES
    SITE.Clear
    SITE.AddItem "MAIL"
    SITE.AddItem "PHONE"
    SITE.AddItem "NEWS PAPER"
    SITE.AddItem "ENTERTAINMENT"
    SITE.AddItem "INFORMATION"
    
End Sub
Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
   End Sub
Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub
Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
 End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub Form_Resize()
    cboAddress.Width = Me.ScaleWidth - 100
    brwWebBrowser.Width = Me.ScaleWidth - 100
    brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
 RS1.Close
 RS2.Close
 RS3.Close
 RS4.Close
 RS5.Close
 db.Close
End Sub

Private Sub INF_Click()
  TMP = "WWW." & INF.Text & ".COM"
  cboAddress.Text = TMP
  cboAddress_Click
End Sub

Private Sub MAIL_Click()
  TMP = "WWW." & MAIL.Text & ".COM"
  cboAddress.Text = TMP
  cboAddress_Click
End Sub
Private Sub NEWS_Click()
  TMP = "WWW." & NEWS.Text & ".COM"
  cboAddress.Text = TMP
  cboAddress_Click
End Sub

Private Sub PHONE_Click()
  TMP = "WWW." & PHONE.Text & ".COM"
  cboAddress.Text = TMP
  cboAddress_Click
End Sub
Private Sub SITE_Click()
 Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Char = Chr(KeyAscii)
 KeyAscii = Asc(UCase(Char))
 If KeyAscii = 13 Then
  Text2.SetFocus
 End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
 Cont.Enabled = True
 If KeyAscii = 13 Then
  Cont_Click
 End If
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working..."
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
      timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            brwWebBrowser.GoHome
        Case "Search"
            brwWebBrowser.GoSearch
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
    End Select

End Sub

