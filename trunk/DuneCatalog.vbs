' =============================================================================================
' Name		: DuneCatalog
' Version	: 1.0
' Date		: 2012-10-31
' 
' This MediaMonkey Script creates Index Files for a Dune Streamer to use it as a Music Jukebox.
'
' Author	: Allart
'
' INSTALL	: See DuneCatalog.txt
' =============================================================================================
'
'
' Change next values to reflect your Windows/Network/Dune Setup
'
' Location of the music index on the Dune Player
DuneIndexFolder = "J:\_Music\"

' Location of the local (Dune) Music files. It is written as the internal storage path.
DuneMusicFolderName = "storage_name://DuneHDD/"
' Drive Letter of the Dune in Windows
DuneDriveLetter = "J"

' Location of the network Music files, as seen by the Dune. It must be accessible (anonymous)
NetworkMusicFolderName = "smb://bat/music/"
' Drive Letter of the network music path in Windows
NetworkDriveLetter = "U"

' Some default checkboxes
DefaultSortAlbums = TRUE
' Overwrite checkbox is not implemented (yet)
REM DefaultOverwriteFiles = TRUE
OpenAdvancedOptionsByDefault = TRUE

' Changes until here. Keep the rest unchanged, unless you know what you are doing.
' =============================================================================================

lowform = 110
highform = 300
newheight = lowform
if OpenAdvancedOptionsByDefault Then newheight = highform

sScriptName = "DuneCatalog"

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
sScriptPath = fso.GetParentFolderName(Script.ScriptPath)
DCScriptFilesFolder = sScriptPath & "\DuneCatalogFiles\"

Dim list, trackcount
Set list = SDB.CurrentSongList

Sub OnStartUp() ' create form and controls
	'
	If list Is Nothing Then
		SDB.MessageBox "Nothing Selected", mtWarning, Array(mbOk)
		Exit Sub
	Else
		trackcount = list.Count
	End If
	If (trackcount = 0) Then
		SDB.MessageBox "Nothing Selected", mtWarning, Array(mbOk)
		Exit Sub
	End If

	Dim albumlist : Set albumlist = list.Albums
	Dim albumcount : albumcount = list.Albums.Count

	Dim Form1 : Set Form1= SDB.UI.NewForm
	Form1.Common.SetRect 10, 100, 370, newheight
	Form1.BorderStyle = 3
	Form1.FormPosition = 4
	Form1.Caption = "Dune Catalog Creator"
	SDB.Objects("Form1") = Form1
	
	Dim label1 : Set label1 = SDB.UI.Newlabel(Form1)
	label1.Common.SetRect 10, 10, 110, 20
	label1.Caption = "Selected Albums:"
	
	Dim label2 : Set label2 = SDB.UI.Newlabel(Form1)
	label2.Common.SetRect 190, 10, 140, 20
	label2.Caption = "Total Number of Files:"
	
	Dim label3 : Set label3 = SDB.UI.Newlabel(Form1)
	label3.Common.SetRect 100, 10, 35, 20
	label3.Caption = albumcount
	
	Dim label4 : Set label4 = SDB.UI.Newlabel(Form1)
	label4.Common.SetRect 310, 10, 35, 20
	label4.Caption = trackcount
	SDB.Objects("LblTrackCount") = label4
	
	Dim MusicDrive : Set MusicDrive = SDB.UI.NewEdit(Form1)
	MusicDrive.Common.ControlName = "MusicDrive"
	MusicDrive.Common.SetRect 10, 90, 20, 20
	MusicDrive.Text = DuneDriveLetter
	MusicDrive.Common.Hint = "Dune Drive Root (as seen in Windows)"
	Set SDB.Objects("SourceMusicDrive") = MusicDrive
	
	Dim DuneMusicFolder : Set DuneMusicFolder = SDB.UI.NewEdit(Form1)
	DuneMusicFolder.Common.ControlName = "DuneMusicFolder"
	DuneMusicFolder.Common.SetRect 35, 90, 315, 20
	DuneMusicFolder.Text = DuneMusicFolderName
	DuneMusicFolder.Common.Hint = "Local (Dune) Music Folder"
	Set SDB.Objects("MusicFolder") = DuneMusicFolder  
	
	Dim NetMusicDrive : Set NetMusicDrive = SDB.UI.NewEdit(Form1)
	NetMusicDrive.Common.ControlName = "NetMusicDrive"
	NetMusicDrive.Common.SetRect 10, 120, 20, 20
	NetMusicDrive.Text = NetworkDriveLetter
	NetMusicDrive.Common.Hint = "Network Drive Root (as seen in Windows)"
	Set SDB.Objects("SourceNetMusicDrive") = NetMusicDrive  
	
	Dim NetworkMusicFolder : Set NetworkMusicFolder = SDB.UI.NewEdit(Form1)
	NetworkMusicFolder.Common.ControlName = "NetworkMusicFolder"
	NetworkMusicFolder.Common.SetRect 35, 120, 315, 20
	NetworkMusicFolder.Text = NetworkMusicFolderName
	NetworkMusicFolder.Common.Hint = "Network Music Folder"
	Set SDB.Objects("NetMusicFolder") = NetworkMusicFolder
	
	Dim IndexFolder : Set IndexFolder = SDB.UI.NewEdit(Form1)
	IndexFolder.Common.ControlName = "IndexFolder"
	IndexFolder.Common.SetRect 35, 150, 315, 20
	IndexFolder.Text = DuneIndexFolder
	IndexFolder.Common.Hint = "Music Index Folder on Dune"
	Set SDB.Objects("IndexFolder") = IndexFolder  
	
	Set ButtonCancel = SDB.UI.NewButton(Form1)
	ButtonCancel.Common.SetRect 130, 45, 100, 28
	ButtonCancel.Caption = "Cancel"
	Script.RegisterEvent ButtonCancel, "OnClick", "ButtonCancelClick"
	ButtonCancel.Cancel = True
	ButtonCancel.Common.Hint = "End-Stop-Close-Cancel-Exit"
	
	Set ButtonGo = SDB.UI.NewButton(Form1)
	ButtonGo.Common.SetRect 250, 45, 100, 28
	ButtonGo.Caption = "Go"
	ButtonGo.Common.Hint = "Start/Run/GO!"
	Script.RegisterEvent ButtonGo.Common, "OnClick", "ButtonGoClick"
	
	Set cbxAlbumSort = SDB.UI.NewCheckBox(Form1)
	cbxAlbumSort.Caption = "Sort Album Selection"
	cbxAlbumSort.Common.SetRect 35, 190, 315, 20
	cbxAlbumSort.Checked = DefaultSortAlbums
	cbxAlbumSort.Common.Hint = "Sort selection by Album, then by Track number"
	SDB.Objects("SortAlbum") = cbxAlbumSort
	
	REM Set cbxOverwrite = SDB.UI.NewCheckBox(Form1)
	REM cbxOverwrite.Caption = "Overwrite Existing Files (not implemented yet)"
	REM cbxOverwrite.Common.SetRect 35, 210, 315, 20
	REM cbxOverwrite.Checked = DefaultOverwriteFiles
	REM cbxOverwrite.Common.Hint = "Overwrite existing files"
	REM SDB.Objects("OverwriteFiles") = cbxOverwrite
	REM cbxOverwrite.Common.Enabled = FALSE
	
	Dim lblInfo : Set lblInfo = SDB.UI.Newlabel(Form1)
	lblInfo.Common.SetRect 140, 250, 315, 20
	lblInfo.Caption = "Keep mouse on any item for some more info"
	
	Set ButtonOptions = SDB.UI.NewButton(Form1)
	ButtonOptions.Common.SetRect 10, 54, 70, 20
	ButtonOptions.Caption = "Options vvv"
	ButtonOptions.Common.Hint = "Open Advanced Options below"
	Script.RegisterEvent ButtonOptions.Common, "OnClick", "ButtonOptionsClick"
	
	Set ButtonOpen = SDB.UI.NewButton(Form1)
	ButtonOpen.Common.SetRect 10, 245, 120, 20
	ButtonOpen.Caption = "Open Script in Editor"
	ButtonOpen.Common.Hint = "Opens Script in Editor"
	Script.RegisterEvent ButtonOpen.Common, "OnClick", "ButtonOpenClick"
	
	Form1.Common.Visible = True
End Sub

Sub ButtonOptionsClick (Form1)
	Set frm1 = SDB.Objects("Form1")
	HH = frm1.Common.Height
	If HH = lowform Then	
		newheight = highform
	ElseIf HH = highform Then
		newheight = lowform
	End If
	frm1.Common.SetRect frm1.Common.Left, frm1.Common.Top, 370, newheight
End Sub

Sub ButtonOpenClick (Form1)
	cmd = "notepad++ "& sScriptPath & "\" & sScriptName & ".vbs"
	dim objShell : Set objShell = CreateObject ("WScript.Shell")
	On Error Resume Next
	objShell.Run(cmd)
	if Err.Number <> 0 Then
		cmd = "notepad "& sScriptPath & "\" & sScriptName & ".vbs"
		objShell.Run(cmd)
	End If
	Set objShell = Nothing
	Call ButtonCancelClick
End Sub

Sub ButtonGoClick (Form1)
	Dim arrAlbum()
	Dim m3u, m3uvar
	m3u = ""
	m3uvar = ""
	msg = ""
	
	Set musicfolder = SDB.Objects("MusicFolder")
	tf = SwapSlashes(musicfolder.Text)
	Set netmusicfolder = SDB.Objects("NetMusicFolder")
	ntf = SwapSlashes(netmusicfolder.Text)
	
	Set cbxSort = SDB.Objects("SortAlbum") 
	
	LoadAlbumArray arrAlbum
	maxFiles = UBound(arrAlbum,2)
	if (maxFiles > 0) And (cbxSort.Checked) Then SortAlbumArray arrAlbum
	
	Dim Progress : Set Progress = SDB.Progress
	Progress.MaxValue = maxFiles
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")  
	
	For index = 0 to maxFiles ' Loop through All Songs
		Progress.Text = "Processing " & index+1 & "/" & maxFiles+1 & ": " & arrAlbum(3, index)
	
		if UCase(Left(arrAlbum(5, index), 1)) = UCase(DuneDriveLetter) Then loc = tf
		if UCase(Left(arrAlbum(5, index), 1)) = UCase(NetworkDriveLetter) Then loc = ntf
		' create .list.m3u
		if index = 0 Then
			newalbum = TRUE
			REM if isVarAlbum(arrAlbum(7, index)) Then newvaralbum = TRUE
		End If
		newline = loc & SwapSlashes(SkipDrive(arrAlbum(5, index)))
		if HasSpecialCharacter(newline) Then newline = CharSwap(newline) ' ascii/ansi/utf-8 conversion
		If newalbum Then
			m3u = newline & chr(13) & chr(10)
			newalbum = FALSE
		Else
			m3u = m3u & newline & chr(13) & chr(10)
		End If
		If newvaralbum Then
			m3uvar = newline & chr(13) & chr(10)
			newvaralbum = FALSE
		Else
			If isVarAlbum(arrAlbum(7, index)) Then m3uvar = m3uvar & newline & chr(13) & chr(10)
		End If
		If index = maxFiles Then ' End of list
			If isVarAlbum(arrAlbum(7, index)) Then
				CreateCatalogFolders arrAlbum, index, m3uvar, TRUE
				m3uvar = ""
			Else
				CreateCatalogFolders arrAlbum, index, m3u, TRUE
				m3u = ""
			End If
				REM CreateCatalogFolders arrAlbum, index, m3u, TRUE
				REM m3u = ""
			REM Else
				REM CreateCatalogFolders arrAlbum, index, m3u, TRUE
				REM m3u = ""
			REM End If
		Else
			If arrAlbum(6, index) < arrAlbum(6, index+1) Then ' End of Current ArtistAlbum
				If isVarAlbum(arrAlbum(7, index)) Then
					CreateCatalogFolders arrAlbum, index, m3uvar, TRUE
					m3uvar = ""
					newvaralbum = TRUE
				Else
					CreateCatalogFolders arrAlbum, index, m3u, TRUE
					m3u = ""
					newalbum = TRUE
					REM CreateCatalogFolders arrAlbum, index, m3u, TRUE
				REM Else
					REM CreateCatalogFolders arrAlbum, index, m3u, TRUE
					REM m3u = ""
					REM newalbum = TRUE
				End If
			Else
				If isVarAlbum(arrAlbum(7, index)) Then
					REM CreateCatalogFolders arrAlbum, index, m3u, FALSE
					m3u = ""
				End If
			End If
		End If
		Progress.Increase
	Next
	SDB.MessageBox "Files Created, bye!", mtInformation, Array(mbOK)
	SDB.Objects("Form1") = Nothing
End Sub

Function HasSpecialCharacter(iString)
	For i=1 To Len(iString)
		If (Asc(Mid(iString,i,1)) < 128) Then
			a = FALSE
		Else
			a = TRUE
			Exit For
		End If
	Next
	HasSpecialCharacter = a
End Function

Function CharSwap(iString)
	a = Replace(FolderName, "?", "_")
	iString = Replace(iString, "¡", "Â¡")
	iString = Replace(iString, "¢", "Â¢")
	iString = Replace(iString, "£", "Â£")
	iString = Replace(iString, "¤", "Â¤")
	iString = Replace(iString, "¥", "Â¥")
	iString = Replace(iString, "¦", "Â¦")
	iString = Replace(iString, "§", "Â§")
	iString = Replace(iString, "¨", "Â¨")
	iString = Replace(iString, "©", "Â©")
	iString = Replace(iString, "ª", "Âª")
	iString = Replace(iString, "«", "Â«")
	iString = Replace(iString, "¬", "Â¬")
	iString = Replace(iString, "­", "Â­")
	iString = Replace(iString, "®", "Â®")
	iString = Replace(iString, "¯", "Â¯")
	iString = Replace(iString, "°", "Â°")
	iString = Replace(iString, "±", "Â±")
	iString = Replace(iString, "²", "Â²")
	iString = Replace(iString, "³", "Â³")
	iString = Replace(iString, "´", "Â´")
	iString = Replace(iString, "µ", "Âµ")
	iString = Replace(iString, "¶", "Â¶")
	iString = Replace(iString, "·", "Â·")
	iString = Replace(iString, "¸", "Â¸")
	iString = Replace(iString, "¹", "Â¹")
	iString = Replace(iString, "º", "Âº")
	iString = Replace(iString, "»", "Â»")
	iString = Replace(iString, "¼", "Â¼")
	iString = Replace(iString, "½", "Â½")
	iString = Replace(iString, "¾", "Â¾")
	iString = Replace(iString, "¿", "Â¿")
	iString = Replace(iString, "À", "Ã0")
	iString = Replace(iString, "Á", "Ã0")
	iString = Replace(iString, "Â", "Ã0")
	iString = Replace(iString, "Ã", "Ã0")
	iString = Replace(iString, "Ä", "Ã0")
	iString = Replace(iString, "Å", "Ã0")
	iString = Replace(iString, "Æ", "Ã0")
	iString = Replace(iString, "Ç", "Ã0")
	iString = Replace(iString, "È", "Ã0")
	iString = Replace(iString, "É", "Ã0")
	iString = Replace(iString, "Ê", "Ã0")
	iString = Replace(iString, "Ë", "Ã0")
	iString = Replace(iString, "Ì", "Ã0")
	iString = Replace(iString, "Í", "Ã0")
	iString = Replace(iString, "Î", "Ã0")
	iString = Replace(iString, "Ï", "Ã0")
	iString = Replace(iString, "Ð", "Ã0")
	iString = Replace(iString, "Ñ", "Ã0")
	iString = Replace(iString, "Ò", "Ã0")
	iString = Replace(iString, "Ó", "Ã0")
	iString = Replace(iString, "Ô", "Ã0")
	iString = Replace(iString, "Õ", "Ã0")
	iString = Replace(iString, "Ö", "Ã0")
	iString = Replace(iString, "×", "Ã0")
	iString = Replace(iString, "Ø", "Ã0")
	iString = Replace(iString, "Ù", "Ã0")
	iString = Replace(iString, "Ú", "Ã0")
	iString = Replace(iString, "Û", "Ã0")
	iString = Replace(iString, "Ü", "Ã0")
	iString = Replace(iString, "Ý", "Ã0")
	iString = Replace(iString, "Þ", "Ã0")
	iString = Replace(iString, "ß", "Ã0")
	iString = Replace(iString, "à", "Ã0")
	iString = Replace(iString, "á", "Ã¡")
	iString = Replace(iString, "â", "Ã¢")
	iString = Replace(iString, "ã", "Ã£")
	iString = Replace(iString, "ä", "Ã¤")
	iString = Replace(iString, "å", "Ã¥")
	iString = Replace(iString, "æ", "Ã¦")
	iString = Replace(iString, "ç", "Ã§")
	iString = Replace(iString, "è", "Ã¨")
	iString = Replace(iString, "é", "Ã©")
	iString = Replace(iString, "ê", "Ãª")
	iString = Replace(iString, "ë", "Ã«")
	iString = Replace(iString, "ì", "Ã¬")
	iString = Replace(iString, "í", "Ã­")
	iString = Replace(iString, "î", "Ã®")
	iString = Replace(iString, "ï", "Ã¯")
	iString = Replace(iString, "ð", "Ã°")
	iString = Replace(iString, "ñ", "Ã±")
	iString = Replace(iString, "ò", "Ã²")
	iString = Replace(iString, "ó", "Ã³")
	iString = Replace(iString, "ô", "Ã´")
	iString = Replace(iString, "õ", "Ãµ")
	iString = Replace(iString, "ö", "Ã¶")
	iString = Replace(iString, "÷", "Ã·")
	iString = Replace(iString, "ø", "Ã¸")
	iString = Replace(iString, "ù", "Ã¹")
	iString = Replace(iString, "ú", "Ãº")
	iString = Replace(iString, "û", "Ã»")
	iString = Replace(iString, "ü", "Ã¼")
	iString = Replace(iString, "ý", "Ã½")
	iString = Replace(iString, "þ", "Ã¾")
	iString = Replace(iString, "ÿ", "Ã¿")
	CharSwap = iString
End Function

Function ABCFolder(SubFolder)
	ABCFolder = UCase(Left(SubFolder,1))
End Function

Function DuneYearFolder(Y)
	if Not ((vartype(Y) = 2) Or (vartype(Y) = 8) Or (vartype(Y) = 3)) Then
		SDB.MessageBox "Error in Year", mtError, Array(mbOk)
		Exit Function
	End If
	If vartype(Cint(Y)) <> 2 Then
		DuneYearFolder = "unknown"
	Else
		StartY = 10 * (Y \ 10)
		EndY = StartY + 9
		DuneYearFolder = StartY & "-" & EndY
	End If
End Function

Function isNumeric(xyz)
	isNumeric = ((vartype(xyz) = 2) Or (vartype(xyz) = 8) Or (vartype(xyz) = 3))
End Function

Function DuneABCFolder(SubFolder)
	SFLetter = Left(UCase(SubFolder),1)
	SFNumber = Asc(SFLetter)
	DuneABCFolder = "28_-"
	If ((SFNumber > 64) AND (SFNumber < 91)) Then
		intSF = SFNumber - 64
		if (intSF < 10) Then intSF = "0" & intSF' Prepend Zero
		DuneABCFolder = intSF & "_" & SFLetter
	End If
	if ((SFNumber > 48) AND (SFNumber < 58)) Then 	DuneABCFolder = "27_#"
	REM SDB.MessageBox DuneABCFolder, mtInformation, Array(mbOk)
End Function

Sub CreateCatalogFolders(Arr, i, m3ufile, isAlbum)
	' Creates Folder Structure and Copies Files into it.
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	
	' Create Artist Folder
	If isVarAlbum(Arr(7, i)) Then
		ArtistFolder = "Artists\" & DuneABCFolder("Various") & "\Various\"
	Else
		ArtistFolder = "Artists\" & DuneABCFolder(Arr(2, i)) & "\" _
			& Arr(2, i) & "\"
	End If
	ArtistFolder = AllBackSlashes(RemoveSpecialCharacters(ArtistFolder))
	ArtistAlbumFolder = ArtistFolder & Arr(3, i) & " (" & Arr(4, i) & ")\"
	ArtistAlbumFolder = AllBackSlashes(RemoveSpecialCharacters(ArtistAlbumFolder))
	ArtistFolder = DuneIndexFolder & ArtistFolder
	ArtistAlbumFolder = DuneIndexFolder & ArtistAlbumFolder
	REM SDB.MessageBox ArtistAlbumFolder, mtInformation, Array(mbOk)
	' this sequence seems a bit strange.
	' ArtistFolder is needed later. When some steps are combined, an out of memory error
	' occurred. This way it doesn't occur. My guess is that the last call:
	' "folder = dunefolder & folder" does also some kind of typecasting ? Anyway, it works.
	
	' Create Album Folder
	If isVarAlbum(Arr(7, i)) Then
		AlbumFolder = "Albums\" & DuneABCFolder(Arr(3, i)) & "\" _
			& Arr(3, i) & " - Various (" & Arr(4, i) & ")\"
	Else
		AlbumFolder = "Albums\" & DuneABCFolder(Arr(3, i)) & "\" _
			& Arr(3, i) & " - " & Arr(2, i) & " (" & Arr(4, i) & ")\"
	End If
	AlbumFolder = AllBackSlashes(RemoveSpecialCharacters(AlbumFolder))
	AlbumFolder = DuneIndexFolder & AlbumFolder
	
	' Create Year Folder
	If (Arr(4, i) = "") Then
		YearSubFolder = "Unknown\Empty"
	ElseIf isNumeric(Arr(4, i)) Then
		If CInt(Arr(4, i)) < 1950 Then
			YearSubFolder = "0000-1949\" & Arr(4, i)
		Else
			YearSubFolder = DuneYearFolder(Arr(4, i)) & "\" & Arr(4, i)
		End If
	Else
		YearSubFolder = "Unknown\" &  Arr(4, i)
	End If
		
	If isVarAlbum(Arr(7, i)) Then
		YearFolder = "Years\" & YearSubFolder & "\" & Arr(3, i) & " - Various\"
	Else
		YearFolder = "Years\" & YearSubFolder & "\" & Arr(3, i) & " - " & Arr(2, i) & "\"
	End If
	YearFolder = AllBackSlashes(RemoveSpecialCharacters(YearFolder))
	YearFolder = DuneIndexFolder & YearFolder
	
	REM if isAlbum Then 
		FirstFolder = AlbumFolder
	REM Else
		REM FirstFolder = ArtistAlbumFolder
	REM End If
	
	' Create Files
	Call GeneratePath(fso, FirstFolder) ' Create First Path
	m3ufilename = EndSlash(FirstFolder) & ".list.m3u"
	Set m3ufso = fso.CreateTextFile(m3ufilename ,True, False) ' False creates ascii file, which Dune likes/needs
	m3ufso.Write(m3ufile)
	m3ufso.Close ' Create m3u file
	WriteCoverArt Arr, i, FirstFolder & ".icon.png" ' Cover art
	
	set opic=loadpicture(FirstFolder & ".icon.png")
	'height and width properties return in himetric (0.01mm)
	'numeric factors are just to convert them to pixel
	h2=round(opic.height/2540*96)
	w2=round(opic.width/2540*96)
	set opic=nothing
	
	ScaleFactor = Round(350/Max(h2, w2),3)
	WriteDuneFolder EndSlash(FirstFolder) & "dune_folder.txt", ScaleFactor
	
	REM Call fso.CopyFile(DCScriptFilesFolder & "dune_folder.txt", FirstFolder, True) ' dune_folder.txt
	
	REM If isAlbum Then
		Call GeneratePath(fso, YearFolder)
		CopyFiles FirstFolder, YearFolder

		Call GeneratePath(fso, ArtistAlbumFolder)
		CopyFiles FirstFolder, ArtistAlbumFolder
		CopyFolderFiles ArtistFolder, ArtistAlbumFolder
		REM If not isVarAlbum(Arr(7, i)) Then
			REM Call GeneratePath(fso, ArtistAlbumFolder)
			REM CopyFiles FirstFolder, ArtistAlbumFolder
			REM CopyFolderFiles ArtistFolder ArtistAlbumFolder
		REM End If
	REM Else
		REM CopyFolderFiles ArtistFolder ArtistAlbumFolder
	REM End If
End Sub

Function RemoveSpecialCharacters(FolderName)
	' remove special characters forbidden for file- and folder names
	' " 	* 	/ 	: 	< 	> 	? 	\ 	|
	a = Replace(FolderName, "?", "_")
	a = Replace(a, "*", "_")
	a = Replace(a, ":", "_")
	a = Replace(a, "<", "[")
	a = Replace(a, ">", "]")
	a = Replace(a, "|", "_")
	
	a = Replace(a, """", "_")
	RemoveSpecialCharacters = a
End Function

Function EndSlash(pPath)
	If Right(pPath,1) = "\" Then
		EndSlash = pPath
	Else
		EndSlash = pPath&"\"
	End If
End Function

Function GeneratePath(fso,pFolderPath)
  GeneratePath = False
	REM SDB.MessageBox pFolderPath, mtInformation, Array(mbOk))
  If Not fso.FolderExists(pFolderPath) Then
    If GeneratePath(fso,fso.GetParentFolderName(pFolderPath)) Then
      GeneratePath = True
      fso.CreateFolder(pFolderPath)
    End If
  Else
    GeneratePath = True
  End If
End Function

Sub WriteCoverArt(aArr, i, aPath)
	' test for existing cover art, preference for png
	' jpg is also possible, Dune will show cover anyway :)
	' NOTE: filename must be .icon.png, contents may be jpg. bmp is not checked
	Set fso = CreateObject("Scripting.FileSystemObject")
	REM SDB.MessageBox aArr(5, i), mtInformation, Array(mbOk)
	Dim File : Set File = fso.GetFile(aArr(5, i))
	FilePath = File.ParentFolder
	AlbumArtFile = EndSlash(FilePath) & "cover.png"
	If fso.FileExists(AlbumArtFile) Then
		fso.CopyFile AlbumArtFile, aPath
		Exit Sub
	End If
	AlbumArtFile = EndSlash(FilePath) & "front.png"
	If fso.FileExists(AlbumArtFile) Then
		fso.CopyFile AlbumArtFile, aPath
		Exit Sub
	End If
	AlbumArtFile = EndSlash(FilePath) & "folder.png"
	If fso.FileExists(AlbumArtFile) Then
		fso.CopyFile AlbumArtFile, aPath
		Exit Sub
	End If
	AlbumArtFile = EndSlash(FilePath) & "cover.jpg"
	If fso.FileExists(AlbumArtFile) Then
		fso.CopyFile AlbumArtFile, aPath
		Exit Sub
	End If
	AlbumArtFile = EndSlash(FilePath) & "front.jpg"
	If fso.FileExists(AlbumArtFile) Then
		fso.CopyFile AlbumArtFile, aPath
		Exit Sub
	End If
	AlbumArtFile = EndSlash(FilePath) & "folder.jpg"
	If fso.FileExists(AlbumArtFile) Then
		fso.CopyFile AlbumArtFile, aPath
		Exit Sub
	End If
	fso.CopyFile DCScriptFilesFolder & "\cover.png", aPath
End Sub

Sub CopyFiles(src, tgt)
	' Copy files from source to target
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	' copy m3u
	Call fso.CopyFile(src & ".list.m3u", tgt,True)
	' copy png
	Call fso.CopyFile(src & ".icon.png", tgt,True)
	' copy dune_folder.txt
	Call fso.CopyFile(src & "dune_folder.txt", tgt,True)
	REM Call fso.CopyFile(DCScriptFilesFolder & "dune_folder.txt", tgt, True)
End Sub

Sub CopyFolderFiles(tgt, artistF)
	' Copy files from source to target
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	
	' copy and rename dune_folder.txt
	
	If not fso.FileExists(tgt & ".icon.png") Then
		If fso.FileExists(artistF & ".icon.png") Then
			Call fso.CopyFile(artistF & ".icon.png", tgt, True)
			set opic=loadpicture(artistF & ".icon.png")
			'height and width properties return in himetric (0.01mm)
			'numeric factors are just to convert them to pixel
			h2=round(opic.height/2540*96)
			w2=round(opic.width/2540*96)
			set opic=nothing
			
			ScaleFactor = Round(350/Max(h2, w2),3)
			WriteDuneSubFolder EndSlash(tgt) & "dune_folder.txt", ScaleFactor
		Else 
			' copy and rename png
			Call fso.CopyFile(DCScriptFilesFolder & ".icon.png", tgt, True)
			Call fso.CopyFile(DCScriptFilesFolder & "SFdune_folder.txt", tgt & "dune_folder.txt", True)
		End If
	End If
End Sub

Sub LoadAlbumArray(aTracks) ' Loads all files into an Array
	Set Tracks = SDB.CurrentSongList
	Dim NumTracks : NumTracks = Tracks.Count
	For i = 0 to NumTracks-1
		Dim trk : Set trk = Tracks.Item(i)
		AddTrack aTracks, trk
	Next
End Sub

Function SkipDrive(aPath)
	' remove driveletter and :\ of aPath
	' remove :\
	A = Replace(aPath, ":\", "")
	' remove driveletter
	SkipDrive = Right(a,Len(A)-1)
End Function

Function SwapSlashes(aString)
	' Swap All backslashes to forwardslashes
	SwapSlashes = Replace(aString, "\", "/")
End Function

Function AllBackSlashes(aString)
	' Swap All backslashes to forwardslashes
	AllBackSlashes = Replace(aString, "/", "\")
End Function

Sub SortAlbumArray(aTracks)
	' Sort Main Array by Album and TrackNumber
	
	maxFiles = UBound(aTracks,2)
	
	' Sort by Album Name
	LowBound = LBound(aTracks, 2)
	HighBound = UBound(aTracks, 2)
	QuickSortCol aTracks,LowBound,HighBound,3
	
	' indentify albums and sort by tracknumber
	' albums with the same name are not identified (add artist/various/empty)
	albumindex = 1 ' First album index
	newalbum = TRUE
	For i=0 to Ubound(aTracks, 2) - 1
		If newalbum Then
			lowerB = i
			newalbum = FALSE
		End If
		aTracks(6, i) = albumindex
		
		If (aTracks(3, i) <> aTracks(3, i+1)) Then
			newalbum = TRUE
			albumindex = albumindex + 1
			If i > lowerB Then
				QuickSortCol aTracks,lowerB,i,0
			End If
			If (i = maxFiles - 1) Then aTracks(6, i+1) = albumindex
		Else
			If (i = maxFiles - 1) Then
				aTracks(6, i+1) = albumindex
				QuickSortCol aTracks,lowerB,i+1,0
			End If
		End If
	Next
End Sub

Sub AddTrack(myArray, Track)
	' An "Intuitive" Array with fixed #columns and variable #rows CANNOT be REDIM'ed AND having its data PRESERVED.
	' SO: the data "matrix" is transposed ...
	On Error Resume Next
	' [1] Retrieve the index number of the last element in the array
	idxLast = UBound(myArray,2)
	If not Err = 0 Then
			idxLast = -1
			' This array is not empty.
			Err.Clear
	End If
	
	' [2] Resize the array, preserving the current content
	ReDim Preserve myArray(7, idxLast + 1)
	' [3] Add the new element to the array
	discno = ""
	if Track.DiscNumberStr <> 0 Then discno = Track.DiscNumberStr & "."
	REM SDB.MessageBox discno, mtInformation, Array(mbOk)
	REM SDB.MessageBox len(Track.TrackOrderStr), mtInformation, Array(mbOk)
	preZero = ""
	For i = 0 to 3 - len(Track.TrackOrderStr)
		preZero = preZero & "0"
	Next
	REM SDB.MessageBox preZero, mtInformation, Array(mbOk)
	
	myArray(0, idxLast + 1) = discno & preZero & Track.TrackOrderStr
	REM myArray(0, idxLast + 1) = discno & Track.TrackOrderStr
	myArray(1, idxLast + 1) = Track.Title
	
	REM SDB.MessageBox myArray(0, idxLast + 1), mtInformation, Array(mbOk)
	' Artist
	If Track.ArtistName = "" Then
		If isVarAlbum(Track.AlbumArtistName) Then
			myArray(2, idxLast + 1) = "Various"
		Else	
			myArray(2, idxLast + 1) = "Unknown"
		End If
	Else
		myArray(2, idxLast + 1) = Track.ArtistName
	End If
	
	' See if AlbumName exists. If not, name it unknown
	If Track.AlbumName = "" Then
		myArray(3, idxLast + 1) = "Unknown"
	Else
		myArray(3, idxLast + 1) = Track.AlbumName
	End If
	
	' See if Year exists
	If Track.Year = "" Then
		myArray(4, idxLast + 1) = "0000"
	Else
		myArray(4, idxLast + 1) = Track.Year
	End If
	
	myArray(5, idxLast + 1) = Track.Path
	myArray(7, idxLast + 1) = Track.AlbumArtistName
End Sub

Function isVarAlbum(AlbumArtist)
	isVarAlbum = FALSE
	If UCase(Left(AlbumArtist, 7)) = "VARIOUS" Then
		isVarAlbum = TRUE
	Else
		If Left(AlbumArtist, 2) = "VA" Then	
			isVarAlbum = TRUE' must be uppercase by itself
		Else
			If UCase(Left(AlbumArtist, 4)) = "V.A." Then	isVarAlbum = TRUE'
		End If
	End If
End Function

Sub Btn2Click
  SDB.Objects("Form1") = Nothing ' Remove the last reference to our form which also causes it to disappear
  Script.Reload("c:\Users\allart\AppData\Roaming\MediaMonkey\Scripts\DuneCatalog.vbs")
End Sub

Sub ButtonCancelClick
  SDB.Objects("Form1") = Nothing ' Remove the last reference to our form which also causes it to disappear
End Sub

Sub QuickSortCol(vec,loBound,hiBound,SortField)
	'
  '==--------------------------------------------------------==
  '== Sort a 2 dimensional array on SortField                ==
  '==                                                        ==
  '== This procedure is adapted from the algorithm given in: ==
  '==    ~ Data Abstractions & Structures using C++ by ~     ==
  '==    ~ Mark Headington and David Riley, pg. 586    ~     ==
  '== Quicksort is the fastest array sorting routine for     ==
  '== unordered arrays.  Its big O is  n log n               ==
  '==                                                        ==
  '== Parameters:                                            ==
  '== vec       - array to be sorted                         ==
  '== SortField - The field to sort on (2nd dimension value) ==
  '== loBound and hiBound are simply the upper and lower     ==
  '==   bounds of the array's 1st dimension.  It's probably  ==
  '==   easiest to use the LBound and UBound functions to    ==
  '==   set these.                                           ==
  '==--------------------------------------------------------==

  Dim pivot(),loSwap,hiSwap,temp,counter
  Redim pivot (Ubound(vec,1))

  '== Two items to sort
  if hiBound - loBound = 1 then
    if vec(SortField,loBound) > vec(SortField,hiBound) then Call SwapCols(vec,hiBound,loBound)
  End If

  '== Three or more items to sort
  
  For counter = 0 to Ubound(vec,1)
    pivot(counter) = vec(counter,int((loBound + hiBound) / 2))
    vec(counter,int((loBound + hiBound) / 2)) = vec(counter,loBound)
    vec(counter,loBound) = pivot(counter)
  Next

  loSwap = loBound + 1
  hiSwap = hiBound
  
  do
    '== Find the right loSwap
    while loSwap < hiSwap and vec(SortField,loSwap) <= pivot(SortField)
      loSwap = loSwap + 1
    wend
    '== Find the right hiSwap
    while vec(SortField,hiSwap) > pivot(SortField)
      hiSwap = hiSwap - 1
    wend
    '== Swap values if loSwap is less then hiSwap
    if loSwap < hiSwap then Call SwapCols(vec,loSwap,hiSwap)


  loop while loSwap < hiSwap
  
  For counter = 0 to Ubound(vec,1)
    vec(counter,loBound) = vec(counter,hiSwap)
    vec(counter,hiSwap) = pivot(counter)
  Next
    
  '== Recursively call function .. the beauty of Quicksort
    '== 2 or more items in first section
    if loBound < (hiSwap - 1) then Call QuickSortCol(vec,loBound,hiSwap-1,SortField)
    '== 2 or more items in second section
    if hiSwap + 1 < hibound then Call QuickSortCol(vec,hiSwap+1,hiBound,SortField)

End Sub  'QuickSortCol

Sub SwapCols(ary,col1,col2)
  '== This proc swaps two rows of an array 
  Dim x,tempvar
  For x = 0 to Ubound(ary,1)
    tempvar = ary(x, col1)    
    ary(x, col1) = ary(x, col2)
    ary(x, col2) = tempvar
  Next
End Sub  'SwapRows

Sub PrintAlbumArray(aTracks)
	Dim msg
	msg = ""
	idxLast = UBound(aTracks,2)
	For i = 0 to idxLast
		msg = msg & "Alb:" & aTracks(6, i) & "Song:" & chr(9) & aTracks(0, i) & chr(9) & aTracks(1, i) & chr(9) & aTracks(2, i) & chr(9) & aTracks(3, i) & chr(13)
	Next
	SDB.MessageBox msg, mtInformation, Array(mbOk)
End Sub

Sub WriteDuneFolder(filename, scalefactor)
	filecontent = "media_url=.list.m3u" & chr(10) & _
	"paint_scrollbar=no" & chr(10) & _
	"paint_path_box=no" & chr(10) & _
	"paint_help_line=no" & chr(10) & _
	"icon_path=.icon.png" & chr(10) & _
	"icon_scale_factor=" & scalefactor & chr(10) & _
	"use_icon_view=yes" & chr(10) & _
	"icon_valign=center" & chr(13) & chr(10)
	
	Set dftfso = fso.CreateTextFile(filename ,True, False) ' False creates ascii file, which Dune likes/needs
	dftfso.Write(filecontent)
	dftfso.Close ' Create DuneFolder.txt file
End Sub

Sub WriteDuneSubFolder(filename, scalefactor)
	filecontent = "icon_path=.icon.png" & chr(10) & _
	"icon_scale_factor=" & scalefactor & chr(10) & _
	"background_path=../../../.service/.listbackground.jpg" & chr(10) & _
	"use_icon_view=yes" & chr(10) & _
	"icon_valign=center" & chr(10) & _
	"background_x=0" & chr(10) & _
	"background_y=0" & chr(10) & _
	"content_box_x=0" & chr(10) & _
	"content_box_Y=0" & chr(10) & _
	"paint_path_box=no" & chr(10) & _
	"paint_help_line=no" & chr(10) & _
	"paint_scrollbar=yes" & chr(10) & _
	"paint_captions=no" & chr(10) & _
	"num_cols=4" & chr(10) & _
	"num_rows=2" & chr(10) & _
	"paint_icon_selection_box=yes" & chr(10) & _
	"paint_captions=yes" & chr(10) & _
	"paint_icons=yes" & chr(10) & _
	"icon_top=7" & chr(10) & _
	"icon_bottom=100" & chr(10) & _
	"caption_font_size=normal" & chr(13) & chr(10)
	
	Set dftfso = fso.CreateTextFile(filename ,True, False) ' False creates ascii file, which Dune likes/needs
	dftfso.Write(filecontent)
	dftfso.Close ' Create DuneFolder.txt file
End Sub

Function Max(a1, a2)
	If a1 > a2 Then
		Max = a1
	Else
		Max = a2
	End If
End Function