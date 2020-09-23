VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CS Advanced NetStats Monitor"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8550
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check15 
      BackColor       =   &H00000000&
      Caption         =   "Enable Netstats Monitor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame FileInfo 
      BackColor       =   &H00000000&
      Caption         =   "File Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2580
      Width           =   8295
      Begin VB.PictureBox Pic32 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
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
         Height          =   480
         Left            =   120
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   200
         Width           =   480
      End
      Begin VB.PictureBox PicFolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Height          =   255
         Left            =   6360
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   220
         Width           =   255
      End
      Begin VB.Label Label105 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   720
         TabIndex        =   9
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   4560
         TabIndex        =   7
         Top             =   480
         Width           =   690
      End
      Begin VB.Label LblSize 
         BackStyle       =   0  'Transparent
         Caption         =   "Size :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   6
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label lblOpenFolder 
         BackStyle       =   0  'Transparent
         Caption         =   "Open Parent Folder"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6720
         TabIndex        =   5
         Top             =   220
         Width           =   1485
      End
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   0
   End
   Begin VB.PictureBox Pic16 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2640
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicQuestion 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2280
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList Iml16 
      Left            =   3000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView5 
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   390
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "Iml16"
      ForeColor       =   16777215
      BackColor       =   4210752
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Process ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Local IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Local Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Remote Host"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Remote Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Protocal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "File Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label100 
      BackStyle       =   0  'Transparent
      Caption         =   "NetStats Monitor:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   150
      Width           =   1575
   End
   Begin VB.Menu menuColums 
      Caption         =   "Columns"
      Begin VB.Menu menushow 
         Caption         =   "Show Columns"
         Begin VB.Menu menufilename 
            Caption         =   "File Name"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuprocessid 
            Caption         =   "Process ID"
            Checked         =   -1  'True
         End
         Begin VB.Menu menulocalip 
            Caption         =   "Local IP"
            Checked         =   -1  'True
         End
         Begin VB.Menu menulocalport 
            Caption         =   "Local Port"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuremotehost 
            Caption         =   "Remote Host"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuremoteport 
            Caption         =   "Remote Port"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuprotocal 
            Caption         =   "Protocal"
            Checked         =   -1  'True
         End
         Begin VB.Menu menustate 
            Caption         =   "State"
            Checked         =   -1  'True
         End
         Begin VB.Menu menupath 
            Caption         =   "File Path"
            Checked         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LVM_FIRST                 As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH        As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE            As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER  As Long = -2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private TVHost As Long
Private TVPath As String
Private TVTAG As Long
Private TVPI As Long

Public Sub RefreshList()
  Dim i
  Dim Item As ListItem

ListView5.ListItems.Clear

    RefreshStack
    DoEvents

    
    LoadNTProcess
    DoEvents

For i = 0 To GetEntryCount

If Connection(i).State <> 2 Then

    If Connection(i).ProcessName = "" Then
    Set Item = ListView5.ListItems.Add(, , "Unknown")
    Else
    Set Item = ListView5.ListItems.Add(, , Connection(i).ProcessName)
    End If
    
    Item.SubItems(1) = Connection(i).ProcessID
    Item.SubItems(2) = GetIPAddress(Connection(i).LocalHost)
    Item.SubItems(3) = GetPort(Connection(i).LocalPort)
    
    Item.SubItems(4) = GetIPAddress(Connection(i).RemoteHost)
    
    'if there is no remote IP then of course there is no connected port
    If Item.SubItems(4) = "0.0.0.0" Then
    Item.SubItems(5) = "0"
    Else
    Item.SubItems(5) = GetPort(Connection(i).RemotePort)
    End If
    Item.SubItems(6) = "TCP"
    Item.SubItems(7) = c_state(Connection(i).State)
    
    If Connection(i).FileName = "" Then
    Item.SubItems(8) = "Path Unknown"
    Else
    Item.SubItems(8) = Connection(i).FileName
    End If
    
    Item.Tag = i
DoEvents
End If
Next i

GetAllIcons
DoEvents

ShowIcons
DoEvents

Call AutoSizeListViewNetstats(menufilename.Checked, menuprocessid.Checked, menulocalip.Checked, menulocalport.Checked, menuremotehost.Checked, menuremoteport.Checked, menuprotocal.Checked, menustate.Checked, menupath.Checked)
DoEvents

End Sub

Private Sub Check15_Click()
Timer3.Enabled = Check15.Value
End Sub

Private Sub FileInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOpenFolder.FontBold = False
End Sub

Private Sub LblSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOpenFolder.FontBold = False
End Sub

Private Sub ListView5_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim FP As FILE_PARAMS
Dim CurFile As Long

TVHost = Connection(ListView5.ListItems(ListView5.SelectedItem.index).Tag).RemoteHost
TVPath = Connection(ListView5.ListItems(ListView5.SelectedItem.index).Tag).FileName
TVTAG = ListView5.ListItems(ListView5.SelectedItem.index).Tag
TVPI = Connection(ListView5.ListItems(ListView5.SelectedItem.index).Tag).ProcessID

GetLargeIcon (TVPath)

   With FP
      .sFileNameExt = TVPath
   End With
   
CurFile = GetFileInformation(FP)

DoEvents
End Sub

Private Sub menufilename_Click()

menufilename.Checked = Not menufilename.Checked
Call AutoSizeListViewNetstats(menufilename.Checked, menuprocessid.Checked, menulocalip.Checked, menulocalport.Checked, menuremotehost.Checked, menuremoteport.Checked, menuprotocal.Checked, menustate.Checked, menupath.Checked)

End Sub

Private Sub menulocalip_Click()
menulocalip.Checked = Not menulocalip.Checked
Call AutoSizeListViewNetstats(menufilename.Checked, menuprocessid.Checked, menulocalip.Checked, menulocalport.Checked, menuremotehost.Checked, menuremoteport.Checked, menuprotocal.Checked, menustate.Checked, menupath.Checked)
End Sub

Private Sub menulocalport_Click()
menulocalport.Checked = Not menulocalport.Checked
Call AutoSizeListViewNetstats(menufilename.Checked, menuprocessid.Checked, menulocalip.Checked, menulocalport.Checked, menuremotehost.Checked, menuremoteport.Checked, menuprotocal.Checked, menustate.Checked, menupath.Checked)
End Sub

Private Sub menupath_Click()
menupath.Checked = Not menupath.Checked
Call AutoSizeListViewNetstats(menufilename.Checked, menuprocessid.Checked, menulocalip.Checked, menulocalport.Checked, menuremotehost.Checked, menuremoteport.Checked, menuprotocal.Checked, menustate.Checked, menupath.Checked)
End Sub

Private Sub menuprocessid_Click()
menuprocessid.Checked = Not menuprocessid.Checked
Call AutoSizeListViewNetstats(menufilename.Checked, menuprocessid.Checked, menulocalip.Checked, menulocalport.Checked, menuremotehost.Checked, menuremoteport.Checked, menuprotocal.Checked, menustate.Checked, menupath.Checked)
End Sub

Private Sub menuprotocal_Click()
menuprotocal.Checked = Not menuprotocal.Checked
Call AutoSizeListViewNetstats(menufilename.Checked, menuprocessid.Checked, menulocalip.Checked, menulocalport.Checked, menuremotehost.Checked, menuremoteport.Checked, menuprotocal.Checked, menustate.Checked, menupath.Checked)
End Sub

Private Sub menuremotehost_Click()
menuremotehost.Checked = Not menuremotehost.Checked
Call AutoSizeListViewNetstats(menufilename.Checked, menuprocessid.Checked, menulocalip.Checked, menulocalport.Checked, menuremotehost.Checked, menuremoteport.Checked, menuprotocal.Checked, menustate.Checked, menupath.Checked)
End Sub

Private Sub menuremoteport_Click()
menuremoteport.Checked = Not menuremoteport.Checked
Call AutoSizeListViewNetstats(menufilename.Checked, menuprocessid.Checked, menulocalip.Checked, menulocalport.Checked, menuremotehost.Checked, menuremoteport.Checked, menuprotocal.Checked, menustate.Checked, menupath.Checked)
End Sub

Private Sub menustate_Click()
menustate.Checked = Not menustate.Checked
Call AutoSizeListViewNetstats(menufilename.Checked, menuprocessid.Checked, menulocalip.Checked, menulocalport.Checked, menuremotehost.Checked, menuremoteport.Checked, menuprotocal.Checked, menustate.Checked, menupath.Checked)
End Sub

Private Sub PicFolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOpenFolder.FontBold = False
End Sub

Private Sub Timer3_Timer()
If GetRefreshTCP = True Then RefreshList
DoEvents
End Sub
Private Function TrimNull(startstr As String) As String

   Dim pos As Integer
   
   pos = InStr(startstr, Chr$(0))
   
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
  
   TrimNull = startstr
  
End Function
Private Sub GetAllIcons()
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim FileName As String

    ListView5.SmallIcons = Nothing
    Iml16.ListImages.Clear
    
'On Local Error Resume Next
For Each Item In ListView5.ListItems

  FileName = Connection(Item.Tag).FileName

  GetIcon FileName, Item.index
   
Next

End Sub
Private Function GetLargeIcon(FileName As String) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
On Error Resume Next
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long


If FileName = "" Then
'Set imgObj = Iml16.ListImages.Add(index, , PicQuestion.Image)
  
  With Pic32
    Set .Picture = PicQuestion.Image
    .AutoRedraw = True
    .Refresh
  End With
  
Exit Function
End If


'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then

  'Large Icon
  With Pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hLIcon, ShInfo.iIcon, Pic32.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  
    Else

End If

End Function

Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the lvw
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With ListView5
  '.ListItems.Clear
  .SmallIcons = Iml16   'Small
  For Each Item In .ListItems
    Item.SmallIcon = Item.index
  Next
End With

End Sub
Private Function GetIcon(FileName As String, index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
On Error Resume Next
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long


If Connection(ListView5.ListItems(index).Tag).FileName = "Path Unknown" Then
Set imgObj = Iml16.ListImages.Add(index, , PicQuestion.Image)
Exit Function
End If


'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
'hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
'         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then

  'Large Icon
  'With Pic32
  '  Set .Picture = LoadPicture("")
  '  .AutoRedraw = True
  '  r = ImageList_Draw(hLIcon, ShInfo.iIcon, Pic32.hDC, 0, 0, ILD_TRANSPARENT)
  '  .Refresh
  'End With
  
    Else
  'Small Icon
  With Pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, Pic16.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  
  Set imgObj = Iml16.ListImages.Add(index, , Pic16.Image)
End If

End Function

Private Function GetFileInformation(FP As FILE_PARAMS) As Long

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim nSize As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
   Dim itmx As ListItem
   Dim LV As Control
       

  'FP.sFileRoot (assigned to sRoot) contains
  'the path to search.
  '
  'FP.sFileNameExt (assigned to sPath) contains
  'the full path and filespec.
   sPath = FP.sFileNameExt
   
   FileInfo.Caption = "File Information (" & GetFileNameFromPath(sPath) & ")"

  'obtain handle to the first filespec match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then

      
        'remove trailing nulls
         sTmp = TrimNull(WFD.cFileName)
         
        'Even though this routine uses filespecs,
        '*.* is still valid and will cause the search
        'to return folders as well as files, so a
        'check against folders is still required.
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) _
            = FILE_ATTRIBUTE_DIRECTORY Then
      
            
           'retrieve the size and assign to nSize to
           'be returned at the end of this function call
            nSize = nSize + (WFD.nFileSizeHigh * (MAXDWORD + 1)) + WFD.nFileSizeLow
            
           'add to the list if the flag indicates
                             
              'got the data, so add it to the listview
               'Set itmx = lv.ListItems.Add(, , LCase$(sTmp))
               
               'itmx.SubItems(1) = GetFileVersion(sRoot & sTmp)
               'itmx.SubItems(3) = GetFileSizeStr(WFD.nFileSizeHigh + WFD.nFileSizeLow)
               'itmx.SubItems(2) = GetFileDescription(sRoot & sTmp)
               'itmx.SubItems(4) = LCase$(sRoot)
               
               lblOpenFolder.Enabled = True
               
                If GetFileDescription(sPath) = "" Then Lblinfo.Caption = "Description : (No Description) " Else Lblinfo.Caption = "Description : " & GetFileDescription(sPath)
                If GetFileCompany(sPath) = "" Then Label105.Caption = "Company : (No Company) " Else Label105.Caption = "Company : " & GetFileCompany(sPath)
                If GetFileVersion(sPath) = "" Then LblVersion.Caption = "Version : (No Version) " Else LblVersion.Caption = "Version : " & GetFileVersion(sPath)
                LblSize.Caption = "Size : " & GetFileSizeStr(WFD.nFileSizeHigh + WFD.nFileSizeLow)

         Else
         
         lblOpenFolder.Enabled = False
         Lblinfo.Caption = "Description : (Unknown)"
         LblVersion.Caption = "Version : (Unknown)"
         LblSize.Caption = "Size : (Unknown)"
         Label105.Caption = "Company : (Unknown)"
         End If
         
      
     'close the handle
      hFile = FindClose(hFile)
   
            Else
         
         lblOpenFolder.Enabled = False
         Lblinfo.Caption = "Description : (Unknown)"
         LblVersion.Caption = "Version : (Unknown)"
         LblSize.Caption = "Size : (Unknown)"
         Label105.Caption = "Company : (Unknown)"
         
   End If
   
   GetFileInformation = nSize
   
End Function


Private Function GetFileSizeStr(fsize As Long) As String

    GetFileSizeStr = GiveByteValues(Format$((fsize), "###,###,###"))  '& " kb"
  
End Function
Private Function GetFileNameFromPath(ByVal sFullPath As String) As String

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
         
         If sFullPath = "" Then
         GetFileNameFromPath = "Unknown"
         Exit Function
         End If
         
   hFile = FindFirstFile(sFullPath, WFD)
   
   If hFile <> INVALID_HANDLE_VALUE Then
   
     'the filename portion is in cFileName
      GetFileNameFromPath = TrimNull(WFD.cFileName)
      Call FindClose(hFile)
      
   End If
   
End Function
Private Sub lblOpenFolder_Click()
StartNewBrowser (GetFilePath(TVPath, True))
End Sub

Private Sub lblOpenFolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOpenFolder.FontBold = True
End Sub

Private Sub LblVersion_Change()
LblSize.Left = LblVersion.Left + (LblVersion.Width + 300)
End Sub
Private Function StartNewBrowser(sURL As String) As Boolean
       
  'start a new instance of the user's browser
  'at the page passed as sURL
   Dim success As Long
   Dim hProcess As Long
   Dim sBrowser As String
   Dim start As STARTUPINFO
   Dim proc As PROCESS_INFORMATION
   
   sBrowser = GetBrowserName(success)
   
  'did sBrowser get correctly filled?
   If success >= ERROR_FILE_SUCCESS Then
      
     'prepare STARTUPINFO members
      With start
         .cb = Len(start)
         .dwFlags = STARTF_USESHOWWINDOW
         .wShowWindow = SW_SHOWNORMAL
      End With
      
     'start a new instance of the default
     'browser at the specified URL. The
     'lpCommandLine member (second parameter)
     'requires a leading space or the call
     'will fail to open the specified page.
      success = CreateProcess(sBrowser, _
                              " " & sURL, _
                              0&, 0&, 0&, _
                              NORMAL_PRIORITY_CLASS, _
                              0&, 0&, start, proc)
                                  
     'if the process handle is valid, return success
      StartNewBrowser = proc.hProcess <> 0
     
     'don't need the process
     'handle anymore, so close it
      Call CloseHandle(proc.hProcess)

     'and close the handle to the thread created
      Call CloseHandle(proc.hThread)

   End If

End Function


Private Function GetBrowserName(dwFlagReturned As Long) As String

  'find the full path and name of the user's
  'associated browser
   Dim hFile As Long
   Dim sResult As String
   Dim sTempFolder As String
        
  'get the user's temp folder
   sTempFolder = GetTempDir()
   
  'create a dummy html file in the temp dir
   hFile = FreeFile
      Open sTempFolder & "dummy.html" For Output As #hFile
   Close #hFile

  'get the file path & name associated with the file
   sResult = Space$(MAX_PATH)
   dwFlagReturned = FindExecutable("dummy.html", sTempFolder, sResult)
  
  'clean up
   Kill sTempFolder & "dummy.html"
   
  'return result
   GetBrowserName = TrimNull(sResult)
   
End Function
Public Function GetTempDir() As String

  'retrieve the user's system temp folder
   Dim tmp As String
   
   tmp = Space$(256)
   Call GetTempPath(Len(tmp), tmp)
   
   GetTempDir = TrimNull(tmp)
    
End Function
'--end block--'

Public Function BasePath(ByVal fname As String, Optional delim As String = "\", Optional keeplast As Boolean = True) As String
    Dim outstr As String
    Dim llen As Long
    llen = InStrRev(fname, delim)


    If (Not keeplast) Then
        llen = llen - 1
    End If


    If (llen > 0) Then
        BasePath = Mid(fname, 1, llen)
    Else
        BasePath = fname
    End If
End Function
Public Sub AutoSizeListViewNetstats(HideColumn1 As Boolean, HideColumn2 As Boolean, HideColumn3 As Boolean, HideColumn4 As Boolean, HideColumn5 As Boolean, HideColumn6 As Boolean, HideColumn7 As Boolean, HideColumn8 As Boolean, HideColumn9 As Boolean)
    
    If HideColumn1 = True Then
    Call SendMessage(ListView5.hwnd, LVM_SETCOLUMNWIDTH, 0, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Else
    ListView5.ColumnHeaders(1).Width = 0
    End If
    
    If HideColumn2 = True Then
    Call SendMessage(ListView5.hwnd, LVM_SETCOLUMNWIDTH, 1, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Else
    ListView5.ColumnHeaders(2).Width = 0
    End If
    
    If HideColumn3 = True Then
    Call SendMessage(ListView5.hwnd, LVM_SETCOLUMNWIDTH, 2, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Else
    ListView5.ColumnHeaders(3).Width = 0
    End If
    
    If HideColumn4 = True Then
    Call SendMessage(ListView5.hwnd, LVM_SETCOLUMNWIDTH, 3, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Else
    ListView5.ColumnHeaders(4).Width = 0
    End If
    
    If HideColumn5 = True Then
    Call SendMessage(ListView5.hwnd, LVM_SETCOLUMNWIDTH, 4, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Else
    ListView5.ColumnHeaders(5).Width = 0
    End If
    
    If HideColumn6 = True Then
    Call SendMessage(ListView5.hwnd, LVM_SETCOLUMNWIDTH, 5, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Else
    ListView5.ColumnHeaders(6).Width = 0
    End If
    
    If HideColumn7 = True Then
    Call SendMessage(ListView5.hwnd, LVM_SETCOLUMNWIDTH, 6, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Else
    ListView5.ColumnHeaders(7).Width = 0
    End If
    
    If HideColumn8 = True Then
    Call SendMessage(ListView5.hwnd, LVM_SETCOLUMNWIDTH, 7, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Else
    ListView5.ColumnHeaders(8).Width = 0
    End If
    
    If HideColumn9 = True Then
    Call SendMessage(ListView5.hwnd, LVM_SETCOLUMNWIDTH, 8, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Else
    ListView5.ColumnHeaders(9).Width = 0
    End If
  
  
  ListView5.Refresh
End Sub
