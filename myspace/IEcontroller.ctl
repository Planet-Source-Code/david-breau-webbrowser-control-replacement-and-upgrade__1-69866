VERSION 5.00
Begin VB.UserControl IEcontroller1 
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   690
   ScaleWidth      =   600
End
Attribute VB_Name = "IEcontroller1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const GWL_STYLE = (-16)
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const WM_CLOSE = &H10
Private Const WS_CAPTION = &HC00000
Private Const WS_THICKFRAME = &H40000

Enum enGoType
 go_back = 0
 go_forward = 1
 go_home = 2
 go_search = 3
End Enum

Public ref_anch As HTMLAnchorElement
Public ref_table As HTMLTable
Public ref_image As HTMLImg
Public ref_form As HTMLFormElement
Public ref_tablerow As HTMLTableRow
Public bCancelOperation As Boolean
Public IEdoc As HTMLDocument
Private WithEvents privIEdoc As HTMLDocument
Attribute privIEdoc.VB_VarHelpID = -1
Public HiddenDoc As HTMLDocument
Public str_current_IEurl As String
Public str_current_hidden_IEurl As String
Private WithEvents oIE As InternetExplorer
Attribute oIE.VB_VarHelpID = -1
Private WithEvents odoc As HTMLDocument
Attribute odoc.VB_VarHelpID = -1


Event IEDocReady(odoc As HTMLDocument)
Event IEdocState(state As String)
Event HiddenDocReady(odoc As HTMLDocument, surl As String)
Event HiddenDocState(state As String)
Event HiddenDocTimeout(lelapsed As Long)
Event IEDownloadStart()
Event IEDownloadDone()
Event returnedLinks(olink As HTMLAnchorElement, cnt As Integer)
Event returnedImages(oimage As HTMLImg, cnt As Integer)
Event returnedForms(oform As HTMLFormElement, cnt As Integer)
Event returnedTables(otable As HTMLTable, cnt As Integer)
Event returnedTableRows(oTr As HTMLTableRow, cnt As Integer)
Event returnedTableDowns(oTd As HTMLTableCell, cnt As Integer)
Event contextMenu()
Event mousedown(ibutton As Integer)
Event closing()
Event IEcreated()
Event processingDone(sFunctionName As String)
Event Error(sProcName As String, iErrNum As Long, sErrDescr As String)
Event NewWindow()

'Default Property Values:
Const m_def_start_url = "about:blank"
Const m_def_show_menubar = 0
Const m_def_allow_context_menu = 0
Const m_def_show_toolbar = 0
Const m_def_show_statusbar = 0
Const m_def_show_addressbar = 0

'Property Variables:
Dim m_start_url As String
Dim m_show_menubar As Boolean
Dim m_allow_context_menu As Boolean
Dim m_show_toolbar As Boolean
Dim m_show_statusbar As Boolean
Dim m_show_addressbar As Boolean

Private Type variables
  mbshow_titlebar As Boolean
  mb_iedoc_loaded As Boolean
End Type
Dim v As variables
 
Sub Go(go_type As enGoType)
On Error GoTo err_handler:

If oIE Is Nothing Then Exit Sub

If go_type = go_back Then
  oIE.GoBack
ElseIf go_type = go_forward Then
  oIE.GoForward
ElseIf go_type = go_home Then
  oIE.GoHome
ElseIf go_type = go_search Then
  oIE.GoSearch
End If

Exit Sub
err_handler:
  If Err.Number <> 0 Then RaiseEvent Error("IEcontroller1.Go", Err.Number, Err.Description)
End Sub

 
Sub return_links(objhtml As Object)
On Error GoTo err_handler:
Dim i As Integer, icnt As Integer, anch As HTMLAnchorElement
 
With objhtml
 icnt = .getElementsByTagName("a").length - 1
 If icnt <= 0 Then Exit Sub
 Dim oa As HTMLAnchorElement
 
 For i = 0 To icnt
  DoEvents
  If bCancelOperation Then GoTo cleanup:
  Set oa = .getElementsByTagName("a").Item(i)
  DoEvents
  RaiseEvent returnedLinks(oa, icnt + 1)
 Next
End With
 
cleanup:
 Set oa = Nothing
 RaiseEvent processingDone("return_links")

Exit Sub
err_handler:
  If Err.Number <> 0 Then RaiseEvent Error("IEcontroller1.return_links", Err.Number, Err.Description)
End Sub
Sub return_tables(objhtml As Object)
On Error GoTo err_handler:
Dim i As Integer, icnt As Integer
 
With objhtml
 icnt = .getElementsByTagName("table").length - 1
 If icnt <= 0 Then Exit Sub
 Dim oTbl As HTMLTable
 
 For i = 0 To icnt
  DoEvents
  If bCancelOperation Then GoTo cleanup:
  Set oTbl = .getElementsByTagName("table").Item(i)
  DoEvents
  RaiseEvent returnedTables(oTbl, icnt + 1)
 Next i
End With

cleanup:
 Set oTbl = Nothing
 RaiseEvent processingDone("return_tables")
 
Exit Sub
err_handler:
  If Err.Number <> 0 Then RaiseEvent Error("IEcontroller1.return_tables", Err.Number, Err.Description)
End Sub
Sub return_forms(objhtml As Object)
On Error GoTo err_handler:
Dim i As Integer, icnt As Integer

With objhtml
 icnt = .getElementsByTagName("form").length - 1
 If icnt <= 0 Then Exit Sub
 Dim oFrm As HTMLFormElement
 
 For i = 0 To icnt
  DoEvents
  If bCancelOperation Then GoTo cleanup:
  Set oFrm = .getElementsByTagName("form").Item(i)
  DoEvents
  RaiseEvent returnedForms(oFrm, icnt + 1)
 Next i
End With
 
cleanup:
 Set oFrm = Nothing
 RaiseEvent processingDone("return_forms")
 
Exit Sub
err_handler:
  If Err.Number <> 0 Then RaiseEvent Error("IEcontroller1.return_forms", Err.Number, Err.Description)
End Sub
Sub return_images(objhtml As Object)
On Error GoTo err_handler:
Dim i As Integer, icnt As Integer
 
With objhtml
 icnt = .getElementsByTagName("img").length - 1
 If icnt <= 0 Then Exit Sub
 Dim oImg As HTMLTable
 
 For i = 0 To icnt
  DoEvents
  If bCancelOperation Then GoTo cleanup:
  Set oImg = .getElementsByTagName("img").Item(i)
  DoEvents
  RaiseEvent returnedImages(oImg, icnt + 1)
 Next i
End With

cleanup:
 Set oImg = Nothing
 RaiseEvent processingDone("return_images")
 
Exit Sub
err_handler:
  If Err.Number <> 0 Then RaiseEvent Error("IEcontroller1.return_images", Err.Number, Err.Description)
End Sub
Sub return_tablerows(objhtml As Object)
On Error GoTo err_handler:
Dim i As Integer, icnt As Integer

With objhtml
 icnt = .getElementsByTagName("tr").length - 1
 If icnt <= 0 Then Exit Sub
 Dim oTr As HTMLTableRow
 
 For i = 0 To icnt
  DoEvents
  If bCancelOperation Then GoTo cleanup:
  Set oTr = .getElementsByTagName("tr").Item(i)
  DoEvents
  RaiseEvent returnedTableRows(oTr, icnt + 1)
 Next i
End With

cleanup:
 Set oTr = Nothing
 RaiseEvent processingDone("return_tablerows")

Exit Sub
err_handler:
  If Err.Number <> 0 Then RaiseEvent Error("IEcontroller1.return_tablerows", Err.Number, Err.Description)
End Sub
Sub return_tabledowns(objhtml As Object)
On Error GoTo err_handler:
Dim i As Integer, icnt As Integer

With objhtml
 icnt = .getElementsByTagName("td").length - 1
 If icnt <= 0 Then Exit Sub
 Dim oTd As HTMLTableCell
 
 For i = 0 To icnt
  DoEvents
  If bCancelOperation Then GoTo cleanup:
  Set oTd = .getElementsByTagName("td").Item(i)
  DoEvents
  RaiseEvent returnedTableDowns(oTd, icnt + 1)
 Next i
End With

cleanup:
 Set oTd = Nothing
 RaiseEvent processingDone("return_tabledowns")
 
Exit Sub
err_handler:
  If Err.Number <> 0 Then RaiseEvent Error("IEcontroller1.return_tabledowns", Err.Number, Err.Description)
End Sub


Sub CreateHiddenDocument(surl As String, Optional ltimeoutval As Long = 5000)
On Error GoTo err_handler:
Dim lelapsed As Long
Dim old_state As String, curr_state As String
Dim newdoc As New HTMLDocument, odoc As HTMLDocument
 
bCancelOperation = False
Set odoc = newdoc.createDocumentFromUrl(surl, vbNullString)
 
With odoc
 While odoc.readyState <> "complete"
  If bCancelOperation Then Exit Sub
  DoEvents
  
  curr_state = .readyState
  If old_state <> curr_state Then
    old_state = curr_state
    RaiseEvent HiddenDocState(curr_state)
  End If
  Sleep 50
  lelapsed = (lelapsed + 60)
  
  If lelapsed >= ltimeoutval Then
    RaiseEvent HiddenDocTimeout(lelapsed)
    Exit Sub
  End If
  DoEvents
 Wend
 
 Set HiddenDoc = odoc
 str_current_hidden_IEurl = odoc.location.href
 RaiseEvent HiddenDocState("complete")
 RaiseEvent HiddenDocReady(odoc, odoc.location.href)
End With
 
cleanup:
 Set odoc = Nothing
 Set newdoc = Nothing
 RaiseEvent processingDone("CreateHiddenDocument")
 
Exit Sub
err_handler:
  If Err.Number <> 0 Then RaiseEvent Error("IEcontroller1.CreateHiddenDocument", Err.Number, Err.Description)
End Sub
 
Sub WriteHtmlToDoc(odoc As Object, shtml As String)
On Error GoTo err_handler:
 
odoc.Clear
odoc.open
odoc.write shtml
odoc.Close

cleanup:
  RaiseEvent processingDone("WriteHtmlToDoc")

Exit Sub
err_handler:
  With Err
    If .Number <> 0 Then RaiseEvent Error("IEcontroller1.WriteHtmlToDoc", Err.Number, Err.Description)
  End With
End Sub

Sub createIE(iwidth As Integer, iheight As Integer, _
             Optional surl As String, Optional bShowTitlebar As Boolean = True, _
             Optional lngIeParent As Object)
On Error GoTo err_handler:

bCancelOperation = False
'destroy any previous instance
Call terminate_previous
Set oIE = New InternetExplorer
 
With oIE
 If Not m_show_addressbar Then .AddressBar = False
 If Not m_show_statusbar Then .StatusBar = False
 If Not m_show_toolbar Then .ToolBar = False
 If Not m_show_menubar Then .MenuBar = False
 v.mbshow_titlebar = bShowTitlebar
 
 If Len(Trim$(surl)) > 0 Then
  .navigate surl
 Else
  If Len(Trim$(m_start_url)) > 0 Then .navigate m_start_url
 End If
 
 .Parent.Width = (iwidth / Screen.TwipsPerPixelX)
 .Parent.Height = (iheight / Screen.TwipsPerPixelY)
 
 'are we removing IE titlebar
 If Not v.mbshow_titlebar Then
   'remove titlebar
   Call toggle_titlebar(WS_CAPTION, False)
   'prevent resizing
   Call toggle_titlebar(WS_THICKFRAME, False)
 End If
 
 'set a new parent for ie?
 If Not lngIeParent Is Nothing Then
   'we need the parents width and height and convert to pixels
   Dim lwid As Long, lhei As Long
   lwid = (lngIeParent.Width / Screen.TwipsPerPixelX)
   lhei = (lngIeParent.Height / Screen.TwipsPerPixelY)
   
   SetParent oIE.Parent.hwnd, lngIeParent.hwnd
   MoveWindow oIE.Parent.hwnd, 0, 0, lwid, lhei, True
 End If
 
 .Visible = True
End With

cleanup:
 RaiseEvent IEcreated
 RaiseEvent processingDone("createIE")
 
Exit Sub
err_handler:
  If Err.Number <> 0 Then RaiseEvent Error("IEcontroller1.createIE", Err.Number, Err.Description)
End Sub

Private Sub toggle_titlebar(ByVal Bit As Long, ByVal Value As Boolean)
Dim lStyle As Long
   
   ' Retrieve current style bits.
   lStyle = GetWindowLong(oIE.Parent.hwnd, GWL_STYLE)
   
   ' Set requested bit On or Off and Redraw.
   If Value Then
      lStyle = lStyle Or Bit
   Else
      lStyle = lStyle And Not Bit
   End If
   
   Call SetWindowLong(oIE.Parent.hwnd, GWL_STYLE, lStyle)
   Call pRedraw
End Sub

Private Sub pRedraw()
   ' Redraw window with new style.
   Const swpFlags As Long = _
      SWP_FRAMECHANGED Or SWP_NOMOVE Or _
      SWP_NOZORDER Or SWP_NOSIZE
   Call SetWindowPos(oIE.Parent.hwnd, 0, 0, 0, 0, 0, swpFlags)
End Sub

Private Function oDoc_oncontextmenu() As Boolean
On Error Resume Next
oDoc_oncontextmenu = m_allow_context_menu
RaiseEvent contextMenu
End Function

Private Sub odoc_onmousedown()
On Error Resume Next
RaiseEvent mousedown(odoc.parentWindow.event.button)
End Sub
 
Private Sub oIE_DocumentComplete(ByVal pDisp As Object, URL As Variant)
 If v.mb_iedoc_loaded = False Then
  v.mb_iedoc_loaded = True
  Set IEdoc = oIE.document
  Set privIEdoc = oIE.document
  str_current_IEurl = oIE.LocationURL
  RaiseEvent IEDocReady(oIE.document)
 End If
End Sub

Private Sub oIE_DownloadBegin()
On Error Resume Next
  RaiseEvent IEDownloadStart
End Sub
Private Sub oIE_DownloadComplete()
On Error Resume Next
v.mb_iedoc_loaded = False
RaiseEvent IEDownloadDone
End Sub

Private Sub oIE_NewWindow2(ppDisp As Object, Cancel As Boolean)
  RaiseEvent NewWindow
End Sub

Private Sub oIE_OnQuit()
On Error Resume Next

Call terminate_previous
Set oIE = Nothing
Set odoc = Nothing
RaiseEvent closing
End Sub
 

Private Function privIEdoc_oncontextmenu() As Boolean
   privIEdoc_oncontextmenu = m_allow_context_menu
End Function

Private Sub UserControl_Paint()
UserControl.Line (0, 0)-(Width - 20, Height - 20), vbBlack, B
End Sub
Private Sub UserControl_Resize()
UserControl.Size 32 * Screen.TwipsPerPixelX, 32 * Screen.TwipsPerPixelY
End Sub
 
Public Property Get hwnd()
   If oIE Is Nothing Then Exit Property
   hwnd = oIE.Parent.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get allow_context_menu() As Boolean
    allow_context_menu = m_allow_context_menu
End Property
Public Property Let allow_context_menu(ByVal New_allow_context_menu As Boolean)
    m_allow_context_menu = New_allow_context_menu
    PropertyChanged "allow_context_menu"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get show_toolbar() As Boolean
Attribute show_toolbar.VB_Description = "shows the IE toolbar"
    show_toolbar = m_show_toolbar
End Property
Public Property Let show_toolbar(ByVal New_show_toolbar As Boolean)
    m_show_toolbar = New_show_toolbar
    PropertyChanged "show_toolbar"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get show_statusbar() As Boolean
Attribute show_statusbar.VB_Description = "shows the IE status bar"
    show_statusbar = m_show_statusbar
End Property
Public Property Let show_statusbar(ByVal New_show_statusbar As Boolean)
    m_show_statusbar = New_show_statusbar
    PropertyChanged "show_statusbar"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get show_addressbar() As Boolean
Attribute show_addressbar.VB_Description = "shows the IE addressbar"
    show_addressbar = m_show_addressbar
End Property
Public Property Let show_addressbar(ByVal New_show_addressbar As Boolean)
    m_show_addressbar = New_show_addressbar
    PropertyChanged "show_addressbar"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get show_menubar() As Boolean
    show_menubar = m_show_menubar
End Property
Public Property Let show_menubar(ByVal New_show_menubar As Boolean)
    m_show_menubar = New_show_menubar
    PropertyChanged "show_menubar"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,about:blank
Public Property Get start_url() As String
    start_url = m_start_url
End Property
Public Property Let start_url(ByVal New_start_url As String)
    m_start_url = New_start_url
    PropertyChanged "start_url"
End Property
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Error Resume Next
    m_allow_context_menu = m_def_allow_context_menu
    m_show_toolbar = m_def_show_toolbar
    m_show_statusbar = m_def_show_statusbar
    m_show_addressbar = m_def_show_addressbar
    m_show_menubar = m_def_show_menubar
    m_start_url = m_def_start_url
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    m_allow_context_menu = PropBag.ReadProperty("allow_context_menu", m_def_allow_context_menu)
    m_show_toolbar = PropBag.ReadProperty("show_toolbar", m_def_show_toolbar)
    m_show_statusbar = PropBag.ReadProperty("show_statusbar", m_def_show_statusbar)
    m_show_addressbar = PropBag.ReadProperty("show_addressbar", m_def_show_addressbar)
    m_show_menubar = PropBag.ReadProperty("show_menubar", m_def_show_menubar)
    m_start_url = PropBag.ReadProperty("start_url", m_def_start_url)
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Call PropBag.WriteProperty("allow_context_menu", m_allow_context_menu, m_def_allow_context_menu)
    Call PropBag.WriteProperty("show_toolbar", m_show_toolbar, m_def_show_toolbar)
    Call PropBag.WriteProperty("show_statusbar", m_show_statusbar, m_def_show_statusbar)
    Call PropBag.WriteProperty("show_addressbar", m_show_addressbar, m_def_show_addressbar)
    Call PropBag.WriteProperty("show_menubar", m_show_menubar, m_def_show_menubar)
    Call PropBag.WriteProperty("start_url", m_start_url, m_def_start_url)
End Sub

Sub navigate(surl As String)
On Error Resume Next
If oIE Is Nothing Then Exit Sub
oIE.navigate surl
End Sub

Private Sub terminate_previous()
On Error GoTo err_handler:
 
DoEvents
If Not (oIE Is Nothing) Then
  oIE.stop: oIE.Quit
  SendMessage oIE.Parent.hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
  Set oIE = Nothing: Set IEdoc = Nothing: Set privIEdoc = Nothing
End If

Exit Sub
err_handler:
  If Err.Number <> 0 Then RaiseEvent Error("IEcontroller1.terminate_previous", Err.Number, Err.Description)
End Sub
Private Sub UserControl_Terminate()
On Error GoTo err_handler:

Call terminate_previous
Set odoc = Nothing
Set oIE = Nothing
Set IEdoc = Nothing
Set privIEdoc = Nothing
Set HiddenDoc = Nothing
Exit Sub

Exit Sub
err_handler:
  If Err.Number <> 0 Then
    If Err.Number = -2147417848 Then
      Resume Next
    Else
       RaiseEvent Error("IEcontroller1.UserControl_Terminate", Err.Number, Err.Description)
    End If
  End If
End Sub
