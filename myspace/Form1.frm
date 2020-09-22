VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHotTrack 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Timer tmrHotTrack 
      Left            =   5640
      Top             =   4080
   End
   Begin VB.CheckBox ckHotTrack 
      Caption         =   "Hot Track IE"
      Height          =   255
      Left            =   9480
      TabIndex        =   10
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CheckBox ckContextMenu 
      Caption         =   "Allow IE context menu"
      Height          =   255
      Left            =   9480
      TabIndex        =   9
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "write to IE"
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Text            =   "<H1>write whatever you want to IE</H1>"
      Top             =   720
      Width           =   2295
   End
   Begin VB.ListBox lstImages 
      Height          =   2010
      Left            =   6120
      TabIndex        =   5
      Top             =   6600
      Width           =   3255
   End
   Begin VB.ListBox lstForms 
      Height          =   2010
      Left            =   2880
      TabIndex        =   3
      Top             =   6600
      Width           =   3255
   End
   Begin VB.ListBox lstLinks 
      Height          =   2010
      Left            =   0
      TabIndex        =   1
      Top             =   6600
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   6255
      Left            =   0
      ScaleHeight     =   6195
      ScaleWidth      =   9315
      TabIndex        =   0
      Top             =   0
      Width           =   9375
   End
   Begin myspace_hack.IEcontroller1 IE1 
      Left            =   9360
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label lblImages 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label lblForms 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label lblLinks 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long


 
 
Private Sub Form_Load()
  Show
  DoEvents
  IE1.show_statusbar = True
  IE1.createIE 5000, 5000, "www.ip-mask.com/pm/pm.html", False, Picture1
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
  tmrHotTrack.Enabled = False
End Sub
 
Private Sub ckContextMenu_Click()
  IE1.allow_context_menu = ckContextMenu.Value
End Sub

Private Sub ckHotTrack_Click()
  If IE1.IEdoc Is Nothing Then
    ckHotTrack.Value = vbUnchecked
    Exit Sub
  End If
  
  tmrHotTrack.Interval = 50
  tmrHotTrack.Enabled = ckHotTrack.Value
  txtHotTrack.Enabled = ckHotTrack.Value
End Sub

Private Sub Command1_Click()
  IE1.WriteHtmlToDoc IE1.IEdoc, Text1
End Sub

Private Sub Form_Click()
  IE1.show_addressbar = True
End Sub
 

Private Sub IE1_Error(sProcName As String, iErrNum As Long, sErrDescr As String)
  Debug.Print sProcName & "|" & sErrDescr
End Sub

Private Sub IE1_IEDocReady(odoc As MSHTML.HTMLDocument)
  'clear listboxes and labels
  lstLinks.Clear: lstForms.Clear: lstImages.Clear
  lblLinks = "": lblForms = "": lblImages = ""
  'return to listboxes the anchors forms and images
  IE1.return_links odoc
  IE1.return_forms odoc
  IE1.return_images odoc
End Sub

Private Sub IE1_IEdocState(state As String)
  Caption = "IE ready state is " & state
End Sub


Private Sub IE1_returnedForms(oform As MSHTML.HTMLFormElement, cnt As Integer)
  lstForms.AddItem Chr(32) & oform.Name & Chr(32) & " submits data to " & oform.Action
  lblForms.Caption = lstForms.ListCount & " forms in " & IE1.str_current_IEurl
End Sub

Private Sub IE1_returnedImages(oimage As MSHTML.HTMLImg, cnt As Integer)
  lstImages.AddItem "this image is " & (oimage.fileSize \ 1024) & " kilobytes"
  lblImages.Caption = lstImages.ListCount & " images in " & IE1.str_current_IEurl
End Sub

Private Sub IE1_returnedLinks(olink As MSHTML.HTMLAnchorElement, cnt As Integer)
  lstLinks.AddItem olink.href
  lblLinks.Caption = lstLinks.ListCount & " links in " & IE1.str_current_IEurl
End Sub
 
Private Sub tmrHotTrack_Timer()
Dim doc As HTMLDocument
Dim pt As POINTAPI
     
  Set doc = IE1.IEdoc
  GetCursorPos pt
  ScreenToClient IE1.hwnd, pt
  txtHotTrack = doc.elementFromPoint(pt.x, pt.y).outerHTML
End Sub
