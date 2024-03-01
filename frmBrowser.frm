'#Region "Form"
	#if defined(__FB_MAIN__) AndAlso Not defined(__MAIN_FILE__)
		#define __MAIN_FILE__
		#ifdef __FB_WIN32__
			#cmdline "FileBrowser.rc"
		#endif
		Const _MAIN_FILE_ = __FILE__
	#endif
	#include once "mff/Form.bi"
	#include once "mff/ComboBoxEx.bi"
	#include once "mff/TreeView.bi"
	#include once "mff/ListView.bi"
	#include once "mff/Splitter.bi"
	#include once "mff/ImageList.bi"
	#include once "mff/Menus.bi"
	#include once "mff/StatusBar.bi"
	
	#include once "..\MDINotepad\FileAct.bi"
	#include once "win\shlobj.bi"
	
	Using My.Sys.Forms
	
	Type frmBrowserType Extends Form
		mImageList As ULong
		mFileInfo As SHFILEINFO
		mRootNode As TreeNode Ptr
		mSelectPath As WString Ptr
		mClose As Boolean 
		
		Declare Function RootInit() As PTreeNode
		Declare Sub RootList()
		Declare Sub FileList(pathroot As WString, ByRef Item As TreeNode)
		
		Declare Sub Form_Create(ByRef Sender As Control)
		Declare Sub Form_Show(ByRef Sender As Form)
		Declare Sub TreeView1_NodeClick(ByRef Sender As TreeView, ByRef Item As TreeNode)
		Declare Sub TreeView1_Click(ByRef Sender As Control)
		Declare Sub TreeView1_NodeDblClick(ByRef Sender As TreeView, ByRef Item As TreeNode)
		Declare Sub TreeView1_NodeActivate(ByRef Sender As TreeView, ByRef Item As TreeNode)
		Declare Sub TreeView1_SelChanged(ByRef Sender As TreeView, ByRef Item As TreeNode)
		Declare Sub TreeView1_SelChanging(ByRef Sender As TreeView, ByRef Item As TreeNode, ByRef Cancel As Boolean)
		Declare Sub MenuView_Click(ByRef Sender As MenuItem)
		Declare Sub Form_Resize(ByRef Sender As Control, NewWidth As Integer, NewHeight As Integer)
		Declare Sub ListView1_Resize(ByRef Sender As Control, NewWidth As Integer, NewHeight As Integer)
		Declare Sub ListView1_SelectedItemChanged(ByRef Sender As ListView, ByVal ItemIndex As Integer)
		Declare Sub Form_Close(ByRef Sender As Form, ByRef Action As Integer)
		Declare Sub ListView1_Click(ByRef Sender As Control)
		Declare Sub ListView1_ItemClick(ByRef Sender As ListView, ByVal ItemIndex As Integer)
		Declare Sub MenuFile_Click(ByRef Sender As MenuItem)
		Declare Constructor
		
		Dim As ComboBoxEx ComboBoxEx1
		Dim As TreeView TreeView1
		Dim As ListView ListView1
		Dim As PopupMenu PopupMenu1
		Dim As MenuItem MenuItem1, MenuItem2, MenuItem3, MenuItem4, MenuItem5, MenuItem6, MenuOpen, MenuBrowser, MenuItem9, MenuNotepad
		Dim As StatusBar StatusBar1
		Dim As StatusPanel StatusPanel1
		Dim As Splitter Splitter1
	End Type
	
	Constructor frmBrowserType
		#if _MAIN_FILE_ = __FILE__
			With App
				.CurLanguagePath = ExePath & "/Languages/"
				.CurLanguage = "english"
			End With
		#endif
		' frmBrowser
		With This
			.Name = "frmBrowser"
			.Text = "File Browser"
			.Designer = @This
			.OnCreate = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control), @Form_Create)
			.OnShow = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Form), @Form_Show)
			.Caption = "File Browser"
			.OnResize = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control, NewWidth As Integer, NewHeight As Integer), @Form_Resize)
			.OnClose = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Form, ByRef Action As Integer), @Form_Close)
			.StartPosition = FormStartPosition.CenterScreen
			.SetBounds 0, 0, 900, 700
		End With
		' ComboBoxEx1
		With ComboBoxEx1
			.Name = "ComboBoxEx1"
			.Text = ""
			.TabIndex = 0
			.Style = ComboBoxEditStyle.cbDropDown
			.Align = DockStyle.alTop
			.ExtraMargins.Top = 5
			.ExtraMargins.Right = 5
			.ExtraMargins.Left = 5
			.SetBounds 5, 5, 874, 22
			.Designer = @This
			.Parent = @This
		End With
		' TreeView1
		With TreeView1
			.Name = "TreeView1"
			.Text = "TreeView1"
			.TabIndex = 1
			.Align = DockStyle.alLeft
			.ExtraMargins.Top = 5
			.ExtraMargins.Right = 0
			.ExtraMargins.Left = 5
			.ExtraMargins.Bottom = 5
			.EditLabels = True
			.SetBounds 5, 32, 200, 602
			.Designer = @This
			.OnNodeClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As TreeView, ByRef Item As TreeNode), @TreeView1_NodeClick)
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control), @TreeView1_Click)
			.OnNodeDblClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As TreeView, ByRef Item As TreeNode), @TreeView1_NodeDblClick)
			.OnNodeActivate = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As TreeView, ByRef Item As TreeNode), @TreeView1_NodeActivate)
			.OnSelChanged = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As TreeView, ByRef Item As TreeNode), @TreeView1_SelChanged)
			.OnSelChanging = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As TreeView, ByRef Item As TreeNode, ByRef Cancel As Boolean), @TreeView1_SelChanging)
			.Parent = @This
		End With
		' Splitter1
		With Splitter1
			.Name = "Splitter1"
			.Text = "Splitter1"
			.SetBounds 195, 27, 5, 612
			.Designer = @This
			.Parent = @This
		End With
		' ListView1
		With ListView1
			.Name = "ListView1"
			.Text = "ListView1"
			.TabIndex = 2
			.Align = DockStyle.alClient
			.MultiSelect = True
			.ExtraMargins.Top = 5
			.ExtraMargins.Right = 5
			.ExtraMargins.Left = 0
			.ExtraMargins.Bottom = 5
			.GridLines = False
			.ContextMenu = @PopupMenu1
			.Anchor.Top = AnchorStyle.asAnchor
			.Anchor.Right = AnchorStyle.asAnchor
			.Anchor.Left = AnchorStyle.asAnchor
			.Anchor.Bottom = AnchorStyle.asAnchor
			.SetBounds 200, 32, 679, 602
			.Designer = @This
			.OnResize = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control, NewWidth As Integer, NewHeight As Integer), @ListView1_Resize)
			.OnSelectedItemChanged = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As ListView, ByVal ItemIndex As Integer), @ListView1_SelectedItemChanged)
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As Control), @ListView1_Click)
			.OnItemClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As ListView, ByVal ItemIndex As Integer), @ListView1_ItemClick)
			.Parent = @This
		End With
		' StatusBar1
		With StatusBar1
			.Name = "StatusBar1"
			.Text = "StatusBar1"
			.Align = DockStyle.alBottom
			.SetBounds 0, 419, 704, 22
			.Designer = @This
			.Parent = @This
		End With
		' StatusPanel1
		With StatusPanel1
			.Name = "StatusPanel1"
			.Designer = @This
			.Parent = @StatusBar1
		End With
		' PopupMenu1
		With PopupMenu1
			.Name = "PopupMenu1"
			.SetBounds 0, 0, 16, 16
			.Designer = @This
			.Parent = @This
		End With
		' MenuOpen
		With MenuOpen
			.Name = "MenuOpen"
			.Designer = @This
			.Caption = "Open"
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As MenuItem), @MenuFile_Click)
			.Parent = @PopupMenu1
		End With
		' MenuNotepad
		With MenuNotepad
			.Name = "MenuNotepad"
			.Designer = @This
			.Caption = "Notepad"
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As MenuItem), @MenuFile_Click)
			.Parent = @PopupMenu1
		End With
		' MenuBrowser
		With MenuBrowser
			.Name = "MenuBrowser"
			.Designer = @This
			.Caption = "Browser"
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As MenuItem), @MenuFile_Click)
			.Parent = @PopupMenu1
		End With
		' MenuItem9
		With MenuItem9
			.Name = "MenuItem9"
			.Designer = @This
			.Caption = "-"
			.Parent = @PopupMenu1
		End With
		' MenuItem1
		With MenuItem1
			.Name = "MenuItem1"
			.Designer = @This
			.Caption = "Icon"
			.Tag = @"0"
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As MenuItem), @MenuView_Click)
			.Parent = @PopupMenu1
		End With
		' MenuItem2
		With MenuItem2
			.Name = "MenuItem2"
			.Designer = @This
			.Caption = "Detials"
			.Tag = @"1"
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As MenuItem), @MenuView_Click)
			.Parent = @PopupMenu1
		End With
		' MenuItem3
		With MenuItem3
			.Name = "MenuItem3"
			.Designer = @This
			.Caption = "Small Icon"
			.Tag = @"2"
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As MenuItem), @MenuView_Click)
			.Parent = @PopupMenu1
		End With
		' MenuItem4
		With MenuItem4
			.Name = "MenuItem4"
			.Designer = @This
			.Caption = "List"
			.Tag = @"3"
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As MenuItem), @MenuView_Click)
			.Parent = @PopupMenu1
		End With
		' MenuItem5
		With MenuItem5
			.Name = "MenuItem5"
			.Designer = @This
			.Caption = "Title"
			.Tag = @"4"
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As MenuItem), @MenuView_Click)
			.Parent = @PopupMenu1
		End With
		' MenuItem6
		With MenuItem6
			.Name = "MenuItem6"
			.Designer = @This
			.Caption = "Max"
			.Tag = @"5"
			.OnClick = Cast(Sub(ByRef Designer As My.Sys.Object, ByRef Sender As MenuItem), @MenuView_Click)
			.Parent = @PopupMenu1
		End With
	End Constructor
	
	Dim Shared frmBrowser As frmBrowserType
	Debug.Clear
	
	#if _MAIN_FILE_ = __FILE__
		App.DarkMode = True
		frmBrowser.MainForm = True
		frmBrowser.Show
		App.Run
	#endif
'#End Region

Private Sub frmBrowserType.Form_Show(ByRef Sender As Form)
	'Debug.Print "TreeView1_Form_Show"
End Sub

Private Sub frmBrowserType.TreeView1_NodeDblClick(ByRef Sender As TreeView, ByRef Item As TreeNode)
	'Debug.Print "TreeView1_NodeDblClick"
End Sub

Private Sub frmBrowserType.TreeView1_Click(ByRef Sender As Control)
	'Debug.Print "TreeView1_Click"
End Sub

Private Sub frmBrowserType.TreeView1_NodeClick(ByRef Sender As TreeView, ByRef Item As TreeNode)
	'Debug.Print "TreeView1_NodeClick"
End Sub

Private Sub frmBrowserType.TreeView1_NodeActivate(ByRef Sender As TreeView, ByRef Item As TreeNode)
	'Debug.Print "TreeView1_NodeActivate"
End Sub

Private Sub frmBrowserType.TreeView1_SelChanging(ByRef Sender As TreeView, ByRef Item As TreeNode, ByRef Cancel As Boolean)
	'Debug.Print "TreeView1_SelChanging"
End Sub

Private Sub frmBrowserType.ListView1_Resize(ByRef Sender As Control, NewWidth As Integer, NewHeight As Integer)
	'Debug.Print "ListView1_Resize"
End Sub

Private Sub frmBrowserType.ListView1_Click(ByRef Sender As Control)
	'Debug.Print "ListView1_Click"
End Sub

Private Sub frmBrowserType.ListView1_SelectedItemChanged(ByRef Sender As ListView, ByVal ItemIndex As Integer)
	'Debug.Print "ListView1_SelectedItemChanged" & ItemIndex
End Sub

Private Sub frmBrowserType.Form_Resize(ByRef Sender As Control, NewWidth As Integer, NewHeight As Integer)
	Debug.Print "Form_Resize"
	With ListView1
		.Columns.Column(0)->Width = .Width - 20 - .Columns.Column(1)->Width - .Columns.Column(2)->Width - .Columns.Column(3)->Width - .Columns.Column(4)->Width
	End With
	StatusPanel1.Width = StatusBar1.Width - 6
End Sub

Private Sub frmBrowserType.Form_Create(ByRef Sender As Control)
	Debug.Print "Form_Create"
	Debug.Clear
	mImageList = SHGetFileInfo("", 0, @mFileInfo, SizeOf(mFileInfo), SHGFI_SYSICONINDEX Or SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_LARGEICON Or SHGFI_PIDL Or SHGFI_DISPLAYNAME Or SHGFI_TYPENAME Or SHGFI_ATTRIBUTES)
	SendMessage(TreeView1.Handle, TVM_SETIMAGELIST, TVSIL_NORMAL, Cast(LPARAM, mImageList))
	SendMessage(ListView1.Handle, LVM_SETIMAGELIST, LVSIL_NORMAL, Cast(LPARAM, mImageList))
	SendMessage(ListView1.Handle, LVM_SETIMAGELIST, LVSIL_SMALL, Cast(LPARAM, mImageList))
	
	mRootNode = RootInit()
	RootList()
	
	ListView1.Columns.Add("Name", , 150)
	ListView1.Columns.Add("Size", , 100, cfRight)
	ListView1.Columns.Add("Write", , 120)
	ListView1.Columns.Add("Creation", , 120)
	ListView1.Columns.Add("Access", , 120)
End Sub

Private Sub frmBrowserType.TreeView1_SelChanged(ByRef Sender As TreeView, ByRef Item As TreeNode)
	Debug.Print "TreeView1_SelChanged"
	If mClose Then Exit Sub
	
	WLet(mSelectPath, Item.Name)
	ComboBoxEx1.Text = *mSelectPath
	Item.Nodes.Clear
	ListView1.ListItems.Clear
	
	If *mSelectPath = "" Then
		RootList()
		
	Else
		FileList(*mSelectPath, Item)
	End If
End Sub

Private Function frmBrowserType.RootInit() As PTreeNode
	Debug.Print "RootInit"
	If mClose Then Return NULL
	
	Dim lPIDL As ITEMIDLIST Ptr
	Dim Path As WString Ptr = CAllocate(MAX_PATH)
	
	SHGetSpecialFolderLocation(NULL, CSIDL_DRIVES, @lPIDL)
	SHGetPathFromIDList(lPIDL, Path)
	SHGetFileInfo(Cast(LPCTSTR, lPIDL), 0, @mFileInfo, Len(mFileInfo), SHGFI_PIDL Or SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)
	
	Return TreeView1.Nodes.Add(mFileInfo.szDisplayName, *Path, *Path, mFileInfo.iIcon, mFileInfo.iIcon)
End Function

Private Sub frmBrowserType.RootList()
	Debug.Print "RootList"
	If mClose Then Exit Sub
	
	Dim lPIDL As ITEMIDLIST Ptr
	Dim Path As WString Ptr = CAllocate(MAX_PATH)
	Dim mydi(4) As Long = { _
	CSIDL_DESKTOP, _
	CSIDL_PERSONAL, _
	CSIDL_MYMUSIC, _
	CSIDL_MYPICTURES, _
	CSIDL_MYVIDEO _
	}
	Dim i As Integer
	
	mRootNode->Nodes.Clear
	
	For i = 0 To 4
		SHGetSpecialFolderLocation(NULL, mydi(i), @lPIDL)
		SHGetPathFromIDList(lPIDL, Path)
		SHGetFileInfo(Cast(LPCTSTR, lPIDL), 0, @mFileInfo, Len(mFileInfo), SHGFI_PIDL Or SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)
		mRootNode->Nodes.Add(mFileInfo.szDisplayName, *Path, *Path, mFileInfo.iIcon, mFileInfo.iIcon)
	Next
	
	Dim d As String * MAX_PATH
	Dim n As WString Ptr = CAllocate(MAX_PATH)
	Dim lpVolumeNameBuffer As WString Ptr = CAllocate(MAX_PATH)
	Dim lpVolumeSerialNumber As DWORD
	Dim lpFileSystemNameBuffer As WString Ptr = CAllocate(MAX_PATH)
	
	GetLogicalDriveStrings(MAX_PATH, Cast(WString Ptr, @d))
	For i = 65 To 90
		If InStr(d, WChr(i)) Then
			*n = WChr(i) & ":"
			GetVolumeInformation(n, lpVolumeNameBuffer, MAX_PATH, @lpVolumeSerialNumber, 0, 0, lpFileSystemNameBuffer, MAX_PATH)
			SHGetFileInfo(n, FILE_ATTRIBUTE_NORMAL, @mFileInfo, SizeOf(mFileInfo), SHGFI_USEFILEATTRIBUTES Or SHGFI_SMALLICON Or SHGFI_SYSICONINDEX)
			mRootNode->Nodes.Add(*lpVolumeNameBuffer & " (" & *n & ")", *n, *n, mFileInfo.iIcon, mFileInfo.iIcon)
		End If
	Next
	
	mRootNode->Expand()
End Sub

Private Sub frmBrowserType.FileList(pathroot As WString, ByRef Item As TreeNode)
	Debug.Print "FileList"
	If mClose Then Exit Sub
	
	Dim wfd As WIN32_FIND_DATA
	Dim hFind As Any Ptr
	Dim hNext As WINBOOL
	Dim fullname As WString Ptr
	Dim i As Integer
	hFind = FindFirstFile(pathroot & "\*.*", @wfd)
	If hFind <> INVALID_HANDLE_VALUE Then
		Do
			If wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
				If wfd.cFileName <> "." And wfd.cFileName <> ".." Then
					WLet(fullname, pathroot & "\" & wfd.cFileName)
					SHGetFileInfo(fullname, wfd.dwFileAttributes, @mFileInfo, SizeOf(mFileInfo), SHGFI_USEFILEATTRIBUTES Or SHGFI_SMALLICON Or SHGFI_SYSICONINDEX Or SHGFI_ICON Or SHGFI_LARGEICON)
					Item.Nodes.Add(wfd.cFileName, pathroot & "\" & wfd.cFileName, pathroot & "\" & wfd.cFileName, mFileInfo.iIcon, mFileInfo.iIcon)
					ListView1.ListItems.Add(wfd.cFileName, mFileInfo.iIcon)
					i = ListView1.ListItems.Count - 1
					ListView1.ListItems.Item(i)->Text(2) = WFD2TimeStr(wfd.ftCreationTime)
					ListView1.ListItems.Item(i)->Text(3) = WFD2TimeStr(wfd.ftLastAccessTime)
					ListView1.ListItems.Item(i)->Text(4) = WFD2TimeStr(wfd.ftLastWriteTime)
				End If
			Else
				WLet(fullname, pathroot & "\" & wfd.cFileName)
				SHGetFileInfo(fullname, FILE_ATTRIBUTE_NORMAL, @mFileInfo, SizeOf(mFileInfo), SHGFI_USEFILEATTRIBUTES Or SHGFI_SMALLICON Or SHGFI_SYSICONINDEX Or SHGFI_ICON Or SHGFI_LARGEICON)
				ListView1.ListItems.Add(wfd.cFileName, mFileInfo.iIcon)
				i = ListView1.ListItems.Count - 1
				ListView1.ListItems.Item(i)->Text(1) = Format(WFD2Size(@wfd), "#,#")
				ListView1.ListItems.Item(i)->Text(2) = WFD2TimeStr(wfd.ftLastWriteTime)
				ListView1.ListItems.Item(i)->Text(3) = WFD2TimeStr(wfd.ftCreationTime)
				ListView1.ListItems.Item(i)->Text(4) = WFD2TimeStr(wfd.ftLastAccessTime)
			End If
			hNext = FindNextFile(hFind , @wfd)
		Loop While (hNext)
		FindClose(hFind)
	End If
	FindClose(hFind)
	If fullname Then Deallocate(fullname)
	If Item.Nodes.Count < 0 Then Exit Sub
	Item.Expand
End Sub

Private Sub frmBrowserType.MenuFile_Click(ByRef Sender As MenuItem)
	Dim st As String
	Dim file As String
	Dim i As Integer
	
	Dim j As Integer = ListView1.ListItems.Count - 1
	For i = 0 To j
		If ListView1.ListItems.Item(i)->Selected Then
			file = *mSelectPath & "\" & ListView1.ListItems.Item(i)->Text(0)
			Exit For
		End If
	Next
	
	Select Case Sender.Name
	Case "MenuOpen"
		st = "Open = " & ShellExecute (Handle, "open", file, "", "", 1)
	Case "MenuNotepad"
		st = "Notepad = " & Exec ("c:\windows\notepad.exe" , file)
	Case "MenuBrowser"
		st = "Browser = " & Exec ("c:\windows\explorer.exe" , "/select," & file)
	End Select
	StatusPanel1.Caption = st
End Sub

Private Sub frmBrowserType.MenuView_Click(ByRef Sender As MenuItem)
	Debug.Print "MenuView_Click"
	MenuItem1.Checked = False
	MenuItem1.Checked = False
	MenuItem2.Checked = False
	MenuItem3.Checked = False
	MenuItem4.Checked = False
	MenuItem5.Checked = False
	MenuItem6.Checked = False
	Sender.Checked = True
	ListView1.View = Cast(ViewStyle, CLng(*Cast(WString Ptr, Sender.Tag)))
End Sub

Private Sub frmBrowserType.Form_Close(ByRef Sender As Form, ByRef Action As Integer)
	Debug.Print "Form_Close"
	mClose = True
	App.DoEvents()
	
	ListView1.ListItems.Clear
	TreeView1.Nodes.Clear
End Sub

Private Sub frmBrowserType.ListView1_ItemClick(ByRef Sender As ListView, ByVal ItemIndex As Integer)
	Debug.Print "ListView1_ItemClick" & ItemIndex
	If ItemIndex < 0 Then
		StatusPanel1.Caption = *mSelectPath & ItemIndex
	Else
		StatusPanel1.Caption = *mSelectPath & "\" & ListView1.ListItems.Item(ItemIndex)->Text(0)
	End If
End Sub

