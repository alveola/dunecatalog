Option Explicit
'
' ==========================================================================================
' A MediaMonkey Script creates Index Files for a Dune Streamer to use it as a Music Jukebox.
' 
' Name    : DuneCatalog
' Version : 1.8
Dim dcVersion : dcVersion=1.8
' Date    : 2012-12-09
' INSTALL : See DuneCatalog.txt
' URL     : http://code.google.com/p/dunecatalog/
' ==========================================================================================
'
' added:
' - playing individual tracks is now possible
' - play albums gapless as 'slideshow'
'
' changed:
' - index structure is changed:
'    * less files are written and copied
'    * script is a bit faster. 
'
' ==========================================================================================
' Change next values to reflect your Windows/Network/Dune Setup
' ==========================================================================================
'

Dim DuneIndexFolder, DuneMusicFolderName, DuneDriveLetter, NetworkMusicFolderName, NetworkDriveLetter
Dim SortAlbumsByDefault, DefaultOverwriteFiles, OpenAdvancedOptionsByDefault, ThoroughAlbumArtScanByDefault
Dim AddTracksBranchDefault
Dim arrAlbum(), StartTime, EndTime, CurSecs
Dim GlassBubbleDefault
Dim YearBeforeAlbumDefault

' name of the ImageMagick Convert program
const strConv = """c:\Program Files (x86)\ImageMagick-6.8.0-Q16\convert.exe"""

' Location of the music index on the Dune Player
DuneIndexFolder = "J:\_index\music\"
REM DuneIndexFolder = "e:\DuneIndex\"

' Location of the local (Dune) Music files. It is written as the internal storage path.
DuneMusicFolderName = "storage_name://DuneHDD/"
' Drive Letter of the Dune in Windows
DuneDriveLetter = "J"

' Location of the network Music files, as seen by the Dune. It must be accessible (anonymous)
NetworkMusicFolderName = "smb://bat/music/"
' Drive Letter of the network music path in Windows
NetworkDriveLetter = "U"

' Some default checkboxes
' Sorting Albums
SortAlbumsByDefault = TRUE
' Overwrite checkbox
REM DefaultOverwriteFiles = FALSE
DefaultOverwriteFiles = TRUE
' Thorough AlbumArt Scan
ThoroughAlbumArtScanByDefault = TRUE
' Glass Bubble default setting
GlassBubbleDefault = FALSE
' Put Year before Album
YearBeforeAlbumDefault = TRUE
' Add Tracks Branch by default
AddTracksBranchDefault = TRUE
' Open lower Advanced Options panel by default
REM OpenAdvancedOptionsByDefault = FALSE
OpenAdvancedOptionsByDefault = TRUE
' Changes until here. Keep the rest unchanged, unless you know what you are doing.
' =============================================================================================

Dim lowform, highform, newheight
lowform = 110
highform = 380
newheight = lowform
if OpenAdvancedOptionsByDefault Then newheight = highform

Dim sScriptName, sScriptPath, DCScriptFilesFolder
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
	Form1.Caption = "Dune Catalog Creator" & " v" & dcVersion
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
	
	Dim ButtonOptions : Set ButtonOptions = SDB.UI.NewButton(Form1)
	ButtonOptions.Common.SetRect 10, 54, 70, 20
	ButtonOptions.Caption = "Options vvv"
	ButtonOptions.Common.Hint = "Open Advanced Options below"
	Script.RegisterEvent ButtonOptions.Common, "OnClick", "ButtonOptionsClick"
	Set SDB.Objects("OpenOptions") = ButtonOptions  
	
	Dim ButtonCancel : Set ButtonCancel = SDB.UI.NewButton(Form1)
	ButtonCancel.Common.SetRect 130, 45, 100, 28
	ButtonCancel.Caption = "Cancel"
	Script.RegisterEvent ButtonCancel, "OnClick", "ButtonCancelClick"
	ButtonCancel.Cancel = True
	ButtonCancel.Common.Hint = "End-Stop-Close-Cancel-Exit"
	
	Dim ButtonGo : Set ButtonGo = SDB.UI.NewButton(Form1)
	ButtonGo.Common.SetRect 250, 45, 100, 28
	ButtonGo.Caption = "Go"
	ButtonGo.Common.Hint = "Start/Run/GO!"
	Script.RegisterEvent ButtonGo.Common, "OnClick", "ButtonGoClick"
	
	Dim cbxAlbumSort : Set cbxAlbumSort = SDB.UI.NewCheckBox(Form1)
	cbxAlbumSort.Caption = "Sort Album Selection"
	cbxAlbumSort.Common.SetRect 35, 190, 315, 20
	cbxAlbumSort.Checked = SortAlbumsByDefault
	cbxAlbumSort.Common.Hint = "Sort selection by Album, then by Track number"
	SDB.Objects("SortAlbum") = cbxAlbumSort
	
	Dim cbxOverwrite : Set cbxOverwrite = SDB.UI.NewCheckBox(Form1)
	cbxOverwrite.Caption = "Overwrite Existing Files"
	cbxOverwrite.Common.SetRect 35, 210, 315, 20
	cbxOverwrite.Checked = DefaultOverwriteFiles
	cbxOverwrite.Common.Hint = "Overwrite existing files"
	SDB.Objects("OverwriteFiles") = cbxOverwrite
	cbxOverwrite.Common.Enabled = TRUE
	
	Dim cbxBetteraArtScan : Set cbxBetteraArtScan = SDB.UI.NewCheckBox(Form1)
	cbxBetteraArtScan.Caption = "Use ImageMagick to find and create alternative Album Art"
	cbxBetteraArtScan.Common.SetRect 35, 230, 315, 20
	cbxBetteraArtScan.Checked = ThoroughAlbumArtScanByDefault
	cbxBetteraArtScan.Common.Hint = "Needs ImageMagick to be installed"
	SDB.Objects("DeepImageScan") = cbxBetteraArtScan
	cbxBetteraArtScan.Common.Enabled = TRUE
	Script.RegisterEvent cbxBetteraArtScan.Common, "OnClick", "UseIMToggle"
	
	Dim cbxGlassBubble : Set cbxGlassBubble = SDB.UI.NewCheckBox(Form1)
	cbxGlassBubble.Caption = "Create GlassBubble Icons"
	cbxGlassBubble.Common.SetRect 55, 250, 315, 20
	cbxGlassBubble.Checked = GlassBubbleDefault
	cbxGlassBubble.Common.Hint = "Needs ImageMagick to be installed"
	SDB.Objects("RoundC") = cbxGlassBubble
	If ThoroughAlbumArtScanByDefault Then
		cbxGlassBubble.Common.Enabled = TRUE
	Else
		cbxGlassBubble.Common.Enabled = FALSE
	End If
	
	Dim cbxSwapAlbumYear : Set cbxSwapAlbumYear = SDB.UI.NewCheckBox(Form1)
	cbxSwapAlbumYear.Caption = "Put Year before Album"
	cbxSwapAlbumYear.Common.SetRect 35, 270, 315, 20
	cbxSwapAlbumYear.Checked = YearBeforeAlbumDefault
	cbxSwapAlbumYear.Common.Hint = "Albums will be displayed in chronological order instead of alphabetical"
	SDB.Objects("YearBeforeAlbum") = cbxSwapAlbumYear
	
	Dim cbxAddTracksBranch : Set cbxAddTracksBranch = SDB.UI.NewCheckBox(Form1)
	cbxAddTracksBranch.Caption = "Add Tracks Branch"
	cbxAddTracksBranch.Common.SetRect 35, 290, 315, 20
	cbxAddTracksBranch.Checked = AddTracksBranchDefault
	cbxAddTracksBranch.Common.Hint = "This will probably take more time"
	SDB.Objects("AddTracksBranch") = cbxAddTracksBranch
	
	Dim lblInfo : Set lblInfo = SDB.UI.Newlabel(Form1)
	lblInfo.Common.SetRect 140, 330, 315, 20
	lblInfo.Caption = "Keep mouse on any item for some more info"
	lblInfo.Common.Hint = "Not on me, you Silly. I'm just a message."
	
	Dim ButtonOpen : Set ButtonOpen = SDB.UI.NewButton(Form1)
	ButtonOpen.Common.SetRect 10, 325, 120, 20
	ButtonOpen.Caption = "Open Script in Editor"
	ButtonOpen.Common.Hint = "Opens Script in Editor"
	Script.RegisterEvent ButtonOpen.Common, "OnClick", "ButtonOpenClick"
	
	Form1.Common.Visible = True
End Sub

Sub BrowseClick(ClickedBtn)
	Dim objShell : Set objShell = CreateObject("Shell.Application")
	Dim objFolder : Set objFolder = objShell.BrowseForFolder(0, "Example", 1, "c:\Programs")
	SDB.MessageBox "folder: " & objFolder.title & " Path: " & objFolder.self.path, 2 , Array(4)
End Sub

Sub ButtonOptionsClick (Form1)
	Dim frm1 : Set frm1 = SDB.Objects("Form1")
	Dim HH : HH = frm1.Common.Height
	Dim oOptions : Set oOptions = SDB.Objects("OpenOptions")
	
	If HH = lowform Then	
		newheight = highform
		oOptions.Caption = "Options ^^^"
		oOptions.Common.Hint = "Close Advanced Options"
	ElseIf HH = highform Then
		newheight = lowform
		oOptions.Caption = "Options vvv"
		oOptions.Common.Hint = "Open Advanced Options below"
	End If
	frm1.Common.SetRect frm1.Common.Left, frm1.Common.Top, 370, newheight
End Sub

Sub ButtonOpenClick (Form1)
	Dim cmd : cmd = "notepad++ "& sScriptPath & "\" & sScriptName & ".vbs"
	Dim objShell : Set objShell = CreateObject ("WScript.Shell")
	On Error Resume Next
	objShell.Run(cmd)
	if Err.Number <> 0 Then
		cmd = "notepad "& sScriptPath & "\" & sScriptName & ".vbs"
		objShell.Run(cmd)
	End If
	Set objShell = Nothing
	ButtonCancelClick
End Sub

Sub UseIMToggle (Form1)
	Dim cbxGlassBubble : Set cbxGlassBubble = SDB.Objects("RoundC")
	Dim cbxBetteraArtScan : Set cbxBetteraArtScan = SDB.Objects("DeepImageScan")
	If cbxBetteraArtScan.checked Then
		cbxGlassBubble.Common.Enabled = TRUE
	Else
		cbxGlassBubble.Common.Enabled = FALSE
	End If
End Sub

Sub ButtonGoClick (Form1)
	StartTime = Timer()
	
	Dim maxFiles, CorrectSourceDir
	Dim albumdft, trackindex
	REM Dim m3uvar, msg	
	Dim musicfolder, netmusicfolder, tf, ntf, loc
	Dim cbxSort, index, newline, newalbum', newvaralbum
	Dim AlbumFolder
	Dim IndexFolder : Set IndexFolder = SDB.Objects("IndexFolder")

	' get data from form:
	Set musicfolder = SDB.Objects("MusicFolder")
	tf = SwapSlashes(musicfolder.Text)
	Set netmusicfolder = SDB.Objects("NetMusicFolder")
	ntf = SwapSlashes(netmusicfolder.Text)
	Dim MusicDrive : Set MusicDrive = SDB.Objects("SourceMusicDrive")
	Dim NetMusicDrive : Set NetMusicDrive = SDB.Objects("SourceNetMusicDrive")
	
	Set cbxSort = SDB.Objects("SortAlbum") 
	
	Erase arrAlbum
	LoadAlbumArray
	maxFiles = UBound(arrAlbum,2)
	if (maxFiles > 0) And (cbxSort.Checked) Then SortAlbumArray
	
	Dim Progress : Set Progress = SDB.Progress
	Progress.MaxValue = maxFiles
	
	newalbum = TRUE
	
	For index = 0 to maxFiles ' Loop through All Songs
		EndTime = Timer()
		CurSecs = EndTime - StartTime
		Progress.Text = "Processing File " & index+1 & " of " & maxFiles+1 & " (" & FormatNumber((index+1)*100/(maxFiles+1),1) & "%), " & _
			FormatNumber(CurSecs,0) & " seconds. Estimation: " & FormatNumber((CurSecs*(maxFiles+1)/(index+1))-CurSecs,0) & " seconds. Current Album: " & arrAlbum(3, index) & " by " & arrAlbum(2, index)
		' Is the Album on the Dune or the accesible network?
		CorrectSourceDir = FALSE
		if UCase(Left(arrAlbum(5, index), 1)) = UCase(MusicDrive.Text) Then 
			loc = tf
			CorrectSourceDir = TRUE
		Elseif UCase(Left(arrAlbum(5, index), 1)) = UCase(NetMusicDrive.Text) Then
			loc = ntf
			CorrectSourceDir = TRUE
		End If
		If CorrectSourceDir Then
			If newalbum Then
				trackindex = 1
				' Create Album Folder
				AlbumFolder = "Albums\" & DuneABCFolder(arrAlbum(3, index)) & "\" _
					& arrAlbum(3, index) & " - " & arrAlbum(2, index) & " (" & arrAlbum(4, index) & ")\"
				AlbumFolder = AllBackSlashes(FolderFix(AlbumFolder))
				newalbum = FALSE
			End If
			
			AddNextTrack Albumdft, index, loc, trackindex
			trackindex = trackindex + 1
			If (index = maxFiles) Then
				WriteAlbum Albumdft, index, AlbumFolder, loc, trackindex
				newalbum = TRUE
				trackindex = 1
			Else
				If (arrAlbum(6, index) < arrAlbum(6, index+1)) Then ' End of list or end of album
					WriteAlbum Albumdft, index, AlbumFolder, loc, trackindex
					newalbum = TRUE
					trackindex = 1
				End If
			End If
			Progress.Increase
		End If
	Next
	
	EndTime = Timer()
	SDB.MessageBox "Files Processed." & chr(10) & chr(13) & _
		"Time: " & FormatNumber(EndTime - StartTime, 1) & " seconds." & chr(10) & chr(13) & _
		"Bye!", mtInformation, Array(mbOK)
	SDB.Objects("Form1") = Nothing
End Sub

Function HasSpecialCharacter(iString)
	Dim i, a
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
	iString = Replace(iString, "?", "_")
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
		Dim StartY : StartY = 10 * (Y \ 10)
		Dim EndY : EndY = StartY + 9
		DuneYearFolder = StartY & "-" & EndY
	End If
End Function

Function YearSubFolder(year)
	If (year = "") Then
		YearSubFolder = "Unknown\Empty"
	ElseIf isNumeric(year) Then
		If CInt(year) < 1950 Then
			If CInt(year) = 0 Then
				YearSubFolder = "0000-1949\0000"
			Else
				YearSubFolder = "0000-1949\" & year
			End If
		Else
			YearSubFolder = DuneYearFolder(year) & "\" & year
		End If
	Else
		YearSubFolder = "Unknown\" &  year
	End If
End Function

Function isNumeric(xyz)
	isNumeric = ((vartype(xyz) = 2) Or (vartype(xyz) = 8) Or (vartype(xyz) = 3))
End Function

Function DuneABCFolder(SubFolder)
	Dim SFLetter, SFNumber, intSF
	SFLetter = Left(UCase(SubFolder),1)
	SFNumber = Asc(SFLetter)
	DuneABCFolder = "28_-"
	If ((SFNumber > 64) AND (SFNumber < 91)) Then
		intSF = SFNumber - 64
		if (intSF < 10) Then intSF = "0" & intSF' Prepend Zero
		DuneABCFolder = intSF & "_" & SFLetter
	End If
	if ((SFNumber > 48) AND (SFNumber < 58)) Then 	DuneABCFolder = "27_#"
End Function

Sub AddNextTrack(filecontent, i, source, ti)
	' create & write file
	Dim caption, mediaurl, TrackFolder
	Dim cbxAddTracksBranch : Set cbxAddTracksBranch = SDB.Objects("AddTracksBranch")
	Dim IndexFolder : Set IndexFolder = SDB.Objects("IndexFolder")
	caption = arrAlbum(0, i) & "." & arrAlbum(1, i)' number.trackname
	If HasSpecialCharacter(caption) Then caption = CharSwap(caption)
	
	If cbxAddTracksBranch.checked Then
		REM - create track folder: Tracks\T1\T2\T3  ==> T3 is trackname-artist
		Dim T2, TwoChars
		TwoChars = FolderFix(UCase(left(arrAlbum(1, i),2)))
		If (Right(TwoChars,1) = " ") Then TwoChars = Left(TwoChars,1)
		T2 = EndSlash(DuneABCFolder(arrAlbum(1, i)) & "\" & TwoChars)
		REM sdb.messagebox T2, 2, array(4)
		TrackFolder = T2 & arrAlbum(1, i) & " - " &	arrAlbum(2, i)
		REM sdb.messagebox TrackFolder, 2, array(4)
		GeneratePath EndSlash(IndexFolder.Text) & "Tracks\" & FolderFix(TrackFolder)
		CreateT2File EndSlash(EndSlash(IndexFolder.Text) & "Tracks\" & FolderFix(T2))
		CreateT3File EndSlash(EndSlash(IndexFolder.Text) & "Tracks\" & FolderFix(TrackFolder)), i, source

		mediaurl = arrAlbum(5, i)
		If HasSpecialCharacter(mediaurl) Then mediaurl = CharSwap(mediaurl)
		If HasSpecialCharacter(TrackFolder) Then TrackFolder = CharSwap(TrackFolder)
		
		mediaurl = SwapSlashes("../../../Tracks/" & TrackFolder)
		' new is browse
		filecontent = filecontent & _
		"item." & ti+10 & ".caption=" & caption & chr(13) & chr(10) & _
		"item." & ti+10 & ".media_url=" & mediaurl & chr(13) & chr(10) & _
		"item." & ti+10 & ".media_action=browse" & chr(13) & chr(10) & _
		"item." & ti+10 & ".icon_path=../../../.service/.empty.png" & chr(13) & chr(10) & _
		"item." & ti+10 & ".icon_scale_factor=1" & chr(13) & chr(10) & _
		"item." & ti+10 & ".icon_valign=center" & chr(13) & chr(10)
		
	Else' tracks in list will directly play
		mediaurl = arrAlbum(5, i)
		If HasSpecialCharacter(mediaurl) Then mediaurl = CharSwap(mediaurl)
		mediaurl = source & SwapSlashes(SkipDrive(mediaurl))
		' original=play
		filecontent = filecontent & _
		"item." & ti+10 & ".caption=" & caption & chr(13) & chr(10) & _
		"item." & ti+10 & ".media_url=" & mediaurl & chr(13) & chr(10) & _
		"item." & ti+10 & ".media_action=play" & chr(13) & chr(10) & _
		"item." & ti+10 & ".icon_path=../../../.service/.empty.png" & chr(13) & chr(10) & _
		"item." & ti+10 & ".icon_scale_factor=1" & chr(13) & chr(10) & _
		"item." & ti+10 & ".icon_valign=center" & chr(13) & chr(10)
	End If
End Sub

Sub WriteAlbum(filecontent, i, Folder, source, ti) ' index, AlbumFolder, loc, trackindex
	Dim filename, iconscalefactor
	Dim ArtistFolder, ArtistAlbumFolder, YearFolder, YearSFolder, AlbumFolder
	Dim cbxYearBeforeAlbum : Set cbxYearBeforeAlbum = SDB.Objects("YearBeforeAlbum")
	Dim IndexFolder : Set IndexFolder = SDB.Objects("IndexFolder")
	Dim ti0 : ti0 = 0
	
	AlbumFolder = EndSlash(IndexFolder.Text) & Folder
	GeneratePath AlbumFolder
	
	' Create .icon.jpg
	If OKtoOverwrite(AlbumFolder & ".icon.jpg") Then WriteCoverArt arrAlbum, i, AlbumFolder & ".icon.jpg" ' Cover art
	
	' Create ArtistAlbum string
	ArtistFolder = "Artists\" & DuneABCFolder(arrAlbum(2, i)) & "\" & arrAlbum(2, i)
	ArtistFolder = AllBackSlashes(FolderFix(ArtistFolder))
	If cbxYearBeforeAlbum.checked Then
		ArtistAlbumFolder = ArtistFolder & "\(" & arrAlbum(4, i) & ") " & arrAlbum(3, i) & "\"
	Else
		ArtistAlbumFolder = ArtistFolder & "\" & arrAlbum(3, i) & " (" & arrAlbum(4, i) & ")\"
	End If
	ArtistFolder = EndSlash(IndexFolder.Text) & EndSlash(ArtistFolder)
	ArtistAlbumFolder = AllBackSlashes(FolderFix(ArtistAlbumFolder))
	ArtistAlbumFolder = EndSlash(IndexFolder.Text) & EndSlash(ArtistAlbumFolder)
	
	' Create Year strings
	YearSFolder = YearSubFolder(arrAlbum(4, i))
	YearFolder = "Years\" & YearSFolder & "\" & arrAlbum(3, i) & " - " & arrAlbum(2, i) & "\"
	YearFolder = AllBackSlashes(FolderFix(YearFolder))
	YearFolder = EndSlash(IndexFolder.Text) & YearFolder

	If OKtoOverwrite(AlbumFolder & "\dune_folder.txt") Then
		Dim MusicFolder : Set MusicFolder = fso.GetFile(arrAlbum(5, i))
		Dim ImDim, ScaleFactor
		ImDim=ImageDimension(AlbumFolder & ".icon.jpg")
		ScaleFactor = Round(350/Max(ImDim(0), ImDim(1)),3)
		
		Dim caption, mediaurl
		
		mediaurl = EndSlash(MusicFolder.ParentFolder)
		If HasSpecialCharacter(mediaurl) Then mediaurl = CharSwap(mediaurl)
		mediaurl = source & SwapSlashes(SkipDrive(mediaurl))

		' Play full Album
		caption = arrAlbum(3, i)
		If HasSpecialCharacter(caption) Then caption = CharSwap(caption)
		filecontent = filecontent & _
			"item." & ti0 & ".caption=- " & caption & " -" & chr(13) & chr(10) & _
			"item." & ti0 & ".media_url=" & mediaurl & chr(13) & chr(10) & _
			"item." & ti0 & ".media_action=play" & chr(13) & chr(10) & _
			"item." & ti0 & ".icon_path=../../../.service/.empty.png" & chr(13) & chr(10) & _
			"item." & ti0 & ".icon_scale_factor=1" & chr(13) & chr(10) & _
			"item." & ti0 & ".icon_valign=center" & chr(13) & chr(10)

		'Jump to Artist
		caption = arrAlbum(2, i)
		If HasSpecialCharacter(caption) Then caption = CharSwap(caption)
		filecontent = filecontent & _
			"item." & ti0+1 & ".caption=- Jump to: " & caption & " -" & chr(13) & chr(10) & _
			"item." & ti0+1 & ".media_url=../../../Artists/" & DuneABCFolder(caption) & "/" & caption & "/" & chr(13) & chr(10) & _
			"item." & ti0+1 & ".media_action=browse" & chr(13) & chr(10) & _
			"item." & ti0+1 & ".icon_path=../../../.service/.empty.png" & chr(13) & chr(10) & _
			"item." & ti0+1 & ".icon_scale_factor=1" & chr(13) & chr(10) & _
			"item." & ti0+1 & ".icon_valign=center" & chr(13) & chr(10)

		' Jump to Year
		filecontent = filecontent & _
			"item." & ti0+2 & ".caption=- Jump to: " & arrAlbum(4, i) & " -" & chr(13) & chr(10) & _
			"item." & ti0+2 & ".media_url=../../../Years/" & SwapSlashes(YearSFolder) & chr(13) & chr(10) & _
			"item." & ti0+2 & ".media_action=browse" & chr(13) & chr(10) & _
			"item." & ti0+2 & ".icon_path=../../../.service/.empty.png" & chr(13) & chr(10) & _
			"item." & ti0+2 & ".icon_scale_factor=1" & chr(13) & chr(10) & _
			"item." & ti0+2 & ".icon_valign=center" & chr(13) & chr(10)

		'Jump to Top of Branch
		filecontent = filecontent & _
			"item." & ti0+3 & ".caption=- Jump to: Index Top -" & chr(13) & chr(10) & _
			"item." & ti0+3 & ".media_url=../../../" & chr(13) & chr(10) & _
			"item." & ti0+3 & ".media_action=browse" & chr(13) & chr(10) & _
			"item." & ti0+3 & ".icon_path=../../../.service/.empty.png" & chr(13) & chr(10) & _
			"item." & ti0+3 & ".icon_scale_factor=1" & chr(13) & chr(10) & _
			"item." & ti0+3 & ".icon_valign=center" & chr(13) & chr(10)

		filecontent = filecontent & _
			"system_files=*.aai,*.jpg,*.png,*.m3u,*.pls,*.txt" & chr(13) & chr(10) & _
			"background_order=first" & chr(13) & chr(10) & _
			"background_path=../../../.service/.listbackground.jpg" & chr(13) & chr(10) & _
			"background_order=before_all" & chr(13) & chr(10) & _
			"paint_path_box=no" & chr(13) & chr(10) & _
			"paint_help_line=no" & chr(13) & chr(10) & _
			"icon_path=.icon.jpg" & chr(13) & chr(10) & _
			"icon_scale_factor=" & ScaleFactor & chr(13) & chr(10) & _
			"use_icon_view=yes" & chr(13) & chr(10) & _
			"icon_valign=center" & chr(13) & chr(10) & _
			"paint_icons=no" & chr(13) & chr(10) & _
			"paint_icon_selection_box=yes" & chr(13) & chr(10) & _
			"num_cols=2" & chr(13) & chr(10) & _
			"num_rows=10" & chr(13) & chr(10) & _
			"caption_font_size=normal" & chr(13) & chr(10) & _
			"sort_field=unsorted" & chr(13) & chr(10) & _
			"sort_dir = asc" & chr(13) & chr(10)
			'sort_field:
			' unsorted takes sort order of virtual item number
			' name sortes by name

		' create & write file
		Dim dftfso : 	Set dftfso = fso.CreateTextFile(AlbumFolder & "\dune_folder.txt" ,True, False)
		dftfso.Write(filecontent)
		dftfso.Close ' Create DuneFolder.txt file
	End If
	
	' Create Artist dft
	If OKtoOverwrite(ArtistFolder & "dune_folder.txt") Then
		GeneratePath ArtistAlbumFolder
		CreateArtistFolderIcon AlbumFolder, ArtistFolder
	End If
	
	If HasSpecialCharacter(Folder) Then Folder = CharSwap(Folder)
	
	' Create ArtistAlbum dft
	If OktoOverwrite(ArtistAlbumFolder & "dune_folder.txt") Then
		filecontent = _
			"paint_scrollbar=no" & chr(13) & chr(10) & _
			"paint_path_box=no" & chr(13) & chr(10) & _
			"paint_help_line=no" & chr(13) & chr(10) & _
			"icon_path=../../../../" & SwapSlashes(Folder) & ".icon.jpg" & chr(13) & chr(10) & _
			"icon_scale_factor=" & Scalefactor & chr(13) & chr(10) & _
			"use_icon_view=yes" & chr(13) & chr(10) & _
			"icon_valign=center" & chr(13) & chr(10) & _
			"media_action=browse" & chr(13) & chr(10) & _
			"media_url=../../../../" & SwapSlashes(Folder) & chr(13) & chr(10)
			'
		Set dftfso = fso.CreateTextFile(ArtistAlbumFolder & "dune_folder.txt" ,True, False)
		dftfso.Write(filecontent)
		dftfso.Close ' Create DuneFolder.txt file
	End If
	
	' Create Year Folder & write dft
	If OKtoOverwrite(YearFolder & "dune_folder.txt") Then
		GeneratePath YearFolder
		Set dftfso = fso.CreateTextFile(YearFolder & "dune_folder.txt" ,True, False)
		dftfso.Write(filecontent)
		dftfso.Close
	End If
	filecontent = ""
End Sub

Sub CreateT2File(Folder)
	REM sdb.messagebox ">>" & Folder & "<<", 2, array(4)
	Dim T2dftfso : Set T2dftfso = fso.CreateTextFile(Folder & "dune_folder.txt" ,True, False) ' False creates ascii file, which Dune likes/needs
	Dim T2filecontent
	T2filecontent = _
		".content_box_x=20" & chr(13) & chr(10) & _
		".content_box_y=20" & chr(13) & chr(10) & _
		"background_order=before_all" & chr(13) & chr(10) & _
		"background_path=../../../.service/.listbackground.jpg" & chr(13) & chr(10) & _
		"background_x=0" & chr(13) & chr(10) & _
		"background_y=0" & chr(13) & chr(10) & _
		"caption_font_size=normal" & chr(13) & chr(10) & _
		"icon_path=../../../.service/.empty.png" & chr(13) & chr(10) & _
		"icon_scale_factor=1" & chr(13) & chr(10) & _
		"icon_top=0" & chr(13) & chr(10) & _
		"icon_valign=center" & chr(13) & chr(10) & _
		"num_cols=2" & chr(13) & chr(10) & _
		"num_rows=10" & chr(13) & chr(10) & _
		"paint_captions=yes" & chr(13) & chr(10) & _
		"paint_help_line=no" & chr(13) & chr(10) & _
		"paint_icon_selection_box=yes" & chr(13) & chr(10) & _
		"paint_path_box=no" & chr(13) & chr(10) & _
		"paint_scrollbar=no" & chr(13) & chr(10) & _
		"text_bottom=10" & chr(13) & chr(10) & _
		"use_icon_view=yes" & chr(13) & chr(10)

	T2dftfso.Write(T2filecontent)
	T2dftfso.Close ' Create DuneFolder.txt file
End Sub

Sub CreateT3File(Folder, index, sourcedir)
	Dim T3dftfso : Set T3dftfso = fso.CreateTextFile(Folder & "dune_folder.txt" ,True, False)
	Dim T3filecontent, ti0 : ti0 = 0
	Dim caption, mediaurl, jArtist
	
	caption = arrAlbum(0, index) & "." & arrAlbum(1, index)' number.trackname
	If HasSpecialCharacter(caption) Then caption = CharSwap(caption)
	mediaurl = arrAlbum(5, index)
	If HasSpecialCharacter(mediaurl) Then mediaurl = CharSwap(mediaurl)
	mediaurl = sourcedir & SwapSlashes(SkipDrive(mediaurl))
	
	T3filecontent = _
		"item." & ti0+10 & ".caption=" & caption & chr(13) & chr(10) & _
		"item." & ti0+10 & ".media_url=" & mediaurl & chr(13) & chr(10) & _
		"item." & ti0+10 & ".media_action=play" & chr(13) & chr(10) & _
		"item." & ti0+10 & ".icon_path=../../../../.service/.empty.png" & chr(13) & chr(10) & _
		"item." & ti0+10 & ".icon_scale_factor=1" & chr(13) & chr(10) & _
		"item." & ti0+10 & ".icon_valign=center" & chr(13) & chr(10)

	'Jump to Artist
	caption = arrAlbum(2, index)
	If HasSpecialCharacter(caption) Then caption = CharSwap(caption)
	jArtist = caption
	T3filecontent = T3filecontent & _
		"item." & ti0+1 & ".caption=- Jump to: " & caption & " -" & chr(13) & chr(10) & _
		"item." & ti0+1 & ".media_url=../../../../Artists/" & DuneABCFolder(caption) & "/" & caption & "/" & chr(13) & chr(10) & _
		"item." & ti0+1 & ".media_action=browse" & chr(13) & chr(10) & _
		"item." & ti0+1 & ".icon_path=../../../../.service/.empty.png" & chr(13) & chr(10) & _
		"item." & ti0+1 & ".icon_scale_factor=1" & chr(13) & chr(10) & _
		"item." & ti0+1 & ".icon_valign=center" & chr(13) & chr(10)
		
	' Jump to Album
	caption = arrAlbum(3, index)
	If HasSpecialCharacter(caption) Then caption = CharSwap(caption)
	T3filecontent = T3filecontent & _
		"item." & ti0 & ".caption=- Jump to: " & caption & " -" & chr(13) & chr(10) & _
		"item." & ti0 & ".media_url=../../../../Albums/" & DuneABCFolder(caption) & "/" & _
			caption & " - " & jArtist & " (" & arrAlbum(4, index) & ")/" & chr(13) & chr(10) & _
		"item." & ti0 & ".media_action=browse" & chr(13) & chr(10) & _
		"item." & ti0 & ".icon_path=../../../../.service/.empty.png" & chr(13) & chr(10) & _
		"item." & ti0 & ".icon_scale_factor=1" & chr(13) & chr(10) & _
		"item." & ti0 & ".icon_valign=center" & chr(13) & chr(10)

	' Jump to Year
	T3filecontent = T3filecontent & _
		"item." & ti0+2 & ".caption=- Jump to: " & arrAlbum(4, index) & " -" & chr(13) & chr(10) & _
		"item." & ti0+2 & ".media_url=../../../../Years/" & SwapSlashes(YearSubFolder(arrAlbum(4, index))) & chr(13) & chr(10) & _
		"item." & ti0+2 & ".media_action=browse" & chr(13) & chr(10) & _
		"item." & ti0+2 & ".icon_path=../../../../.service/.empty.png" & chr(13) & chr(10) & _
		"item." & ti0+2 & ".icon_scale_factor=1" & chr(13) & chr(10) & _
		"item." & ti0+2 & ".icon_valign=center" & chr(13) & chr(10)

	'Jump to Top of Branch
	T3filecontent = T3filecontent & _
		"item." & ti0+3 & ".caption=- Jump to: Index Top -" & chr(13) & chr(10) & _
		"item." & ti0+3 & ".media_url=../../../../" & chr(13) & chr(10) & _
		"item." & ti0+3 & ".media_action=browse" & chr(13) & chr(10) & _
		"item." & ti0+3 & ".icon_path=../../../../.service/.empty.png" & chr(13) & chr(10) & _
		"item." & ti0+3 & ".icon_scale_factor=1" & chr(13) & chr(10) & _
		"item." & ti0+3 & ".icon_valign=center" & chr(13) & chr(10)

	'Jump one level Up
	T3filecontent = T3filecontent & _
		"item." & ti0+4 & ".caption=- Jump to: 1 Up -" & chr(13) & chr(10) & _
		"item." & ti0+4 & ".media_url=../" & chr(13) & chr(10) & _
		"item." & ti0+4 & ".media_action=browse" & chr(13) & chr(10) & _
		"item." & ti0+4 & ".icon_path=../../../../.service/.empty.png" & chr(13) & chr(10) & _
		"item." & ti0+4 & ".icon_scale_factor=1" & chr(13) & chr(10) & _
		"item." & ti0+4 & ".icon_valign=center" & chr(13) & chr(10)

	' remainder of the file
	T3filecontent = T3filecontent & _
		"system_files=*.aai,*.jpg,*.png,*.m3u,*.pls,*.txt" & chr(13) & chr(10) & _
		"background_order=first" & chr(13) & chr(10) & _
		"background_path=../../../../.service/.listbackground.jpg" & chr(13) & chr(10) & _
		"background_order=before_all" & chr(13) & chr(10) & _
		"paint_path_box=no" & chr(13) & chr(10) & _
		"paint_help_line=no" & chr(13) & chr(10) & _
		"icon_path=../../../../.service/.empty.png" & chr(13) & chr(10) & _
		"icon_scale_factor=1" & chr(13) & chr(10) & _
		"use_icon_view=yes" & chr(13) & chr(10) & _
		"icon_valign=center" & chr(13) & chr(10) & _
		"paint_icons=no" & chr(13) & chr(10) & _
		"paint_icon_selection_box=yes" & chr(13) & chr(10) & _
		"num_cols=2" & chr(13) & chr(10) & _
		"num_rows=10" & chr(13) & chr(10) & _
		"caption_font_size=normal" & chr(13) & chr(10) & _
		"sort_field=unsorted" & chr(13) & chr(10) & _
		"sort_dir = asc" & chr(13) & chr(10) & _
		"paint_captions=yes" & chr(13) & chr(10)

	T3dftfso.Write(T3filecontent)
	T3dftfso.Close ' Create DuneFolder.txt file
End Sub

Function ImageDimension(ImageFile)
	Dim returnvalue(2)
	Dim ImFile, ImPath
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	Set ImFile = fso.GetFile(ImageFile)
	Dim UseIM : Set UseIM=SDB.Objects("DeepImageScan")
	If UseIM.Checked Then	Dim oImage, ImSize, ImDim
	
	ImPath = ImFile.Path
	If (fso.GetExtensionName(ImageFile) = "jpg") Then
		On Error Resume Next' Check if jpg is really a jpg ...
		Dim oPic : set oPic=loadpicture(IMPath)
		If Err = 0 Then
			'height and width properties return in himetric (0.01mm)
			'numeric factors are just to convert them to pixel
			returnvalue(0)=round(oPic.height/2540*96)
			returnvalue(1)=round(oPic.width/2540*96)
		Else
			If UseIM.Checked Then
				Set oIMage = CreateObject("ImageMagickObject.MagickImage.1")
				ImSize = oIMage.Identify("-format", "%w %h", ImPath)
				ImDim = Split(ImSize)
				returnvalue(0) = ImDim(0)
				returnvalue(1) = ImDim(1)
				Set oIMage = Nothing
			Else
				returnvalue(0) = 0
				returnvalue(1) = 0
			End If
		End If
		Err.Clear
		set oPic=nothing
	Else
		If UseIM.Checked Then
			Set oIMage = CreateObject("ImageMagickObject.MagickImage.1")
			ImSize = oIMage.Identify("-format", "%w %h", ImPath)
			ImDim = Split(ImSize)
			returnvalue(0) = ImDim(0)
			returnvalue(1) = ImDim(1)
			Set oIMage = Nothing
		Else
			returnvalue(0) = 0
			returnvalue(1) = 0
		End If
	End If
	ImageDimension=returnvalue
	Set fso = Nothing
End Function

REM Function RemoveSpecialCharacters(FolderName)
	REM Dim a
	REM ' remove special characters forbidden for file- and folder names
	REM ' " 	* 	/ 	: 	< 	> 	? 	\ 	|  
	REM a = Replace(FolderName, "?", "_")
	REM a = Replace(a, "*", "_")
	REM a = Replace(a, ":", "_")
	REM a = Replace(a, "<", "[")
	REM a = Replace(a, ">", "]")
	REM a = Replace(a, "|", "_")
	REM a = Replace(a, """", "_")
	REM RemoveSpecialCharacters = a
REM End Function

Function EndSlash(pPath)
	If Right(pPath,1) = "\" Then
		EndSlash = pPath
	Else
		EndSlash = pPath&"\"
	End If
End Function

Function GeneratePath(pFolderPath)
	REM sdb.messagebox pFolderPath, 2, Array(4)
  GeneratePath = False
  If Not fso.FolderExists(pFolderPath) Then
    If GeneratePath(fso.GetParentFolderName(pFolderPath)) Then
      GeneratePath = True
      fso.CreateFolder(pFolderPath)
    End If
  Else
    GeneratePath = True
  End If
End Function

Sub WriteCoverArt(aArr, i, aPath)
	' loop through images. Images are Ranked to find best suitable image:
	' preference for only copying (speed)
	' preference for files with front/folder/cover in their name
	' ranking:
	' Size:	middle=5, large=4, small=2
	' Type: JPG=3, any=2
	' Name: front/voor=6, cover/folder=5, other=1
	
	' Music Source Path
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	Dim File : Set File = fso.GetFile(aArr(5, i))
	Dim FilePath : FilePath = File.ParentFolder
	Dim UseIM : Set UseIM=SDB.Objects("DeepImageScan")
	Dim Imdim, MaxDim, CFN, CFE, SrcCoverArt
	Dim BestRank, CurrentRank, BestIm, BestAction, CurAction
	Dim cbxGlassBubble : Set cbxGlassBubble = SDB.Objects("RoundC")
	Dim SCA	
	
	If UseIM.Checked Then' Advanced Ranking AlbumArt, using IM
		' get image files in folder (jpg, png)
		Dim folder, files, NewsFile, sFolder, fileIdx
		Set folder = fso.GetFolder(FilePath)
		Set files = folder.Files
		
		For Each fileIdx In files
			CurrentRank = 0
			CurAction = 1
			CFE = UCase(fso.GetExtensionName(fileIdx))
			If CFE = "JPG" Then
				CurrentRank = 3
				ImDim = ImageDimension(fileIdx)
				MaxDim = max(ImDim(0), ImDim(1))
			ElseIf (CFE = "BMP" or CFE = "PNG") Then
				CurrentRank = 2
				CurAction = 2
				ImDim = ImageDimension(fileIdx)
				MaxDim = max(ImDim(0), ImDim(1))
			Else
				CurrentRank = 0
				CurAction = 0
				MaxDim = 0
			End If
			CFN = fso.GetFileName(fileIdx)
			If (InStr(1, CFN,"front", 1) or InStr(1, CFN,"voor", 1)) Then
				CurrentRank = 6 * CurrentRank
			ElseIf (InStr(1, CFN,"cover", 1) or InStr(1, CFN,"cover", 1)) Then
				CurrentRank = 5 * CurrentRank
			Else
				CurrentRank = 1 * CurrentRank
			End If
			If (MaxDim > 350) Then
					CurrentRank = 4 * CurrentRank
					CurAction = 2
				ElseIf (MaxDim < 200) Then
					CurrentRank = 2 * CurrentRank
					CurAction = 2
				Else
					CurrentRank = 5 * CurrentRank
			End If
			
			If (CurrentRank > BestRank) Then
				BestRank = CurrentRank
				BestIM = fileIDX.Path
				BestAction = CurAction
			End If
		Next
		
		If BestAction = 1 Then
			fso.CopyFile BestIM, aPath
			SrcCoverArt = EndSlash(fso.GetParentFolderName(BestIM)) & "cover.jpg"
			If (not fso.FileExists(SrcCoverArt)) Then Call fso.CopyFile(aPath, SrcCoverArt)
			set SCA = fso.GetFile(SrcCoverArt)
			SCA.Attributes = 0
		ElseIf BestAction = 2 Then
			resizeImage BestIM, 350, 350, aPath
			SrcCoverArt = EndSlash(fso.GetParentFolderName(BestIM)) & "cover.jpg"
			If (not fso.FileExists(SrcCoverArt)) Then Call fso.CopyFile(aPath, SrcCoverArt)
			set SCA = fso.GetFile(SrcCoverArt)
			SCA.Attributes = 0
		Else
			' default "anonymous" albumart
			fso.CopyFile DCScriptFilesFolder & "\cover.jpg", aPath
		End If
		If cbxGlassBubble.checked Then MakeGlassBubble(aPath)
	Else' Simply Copying existing jpg
		
		Dim AlbumArtFile
		AlbumArtFile = EndSlash(FilePath) & "front.jpg"
		If fso.FileExists(AlbumArtFile) Then
			If FitImage(AlbumArtFile) Then
				fso.CopyFile AlbumArtFile, aPath
				set fso = Nothing
				set files = Nothing
				Exit Sub
			End If
		End If
		AlbumArtFile = EndSlash(FilePath) & "cover.jpg"
		If fso.FileExists(AlbumArtFile) Then
			If FitImage(AlbumArtFile) Then
				fso.CopyFile AlbumArtFile, aPath
				set fso = Nothing
				set files = Nothing
				Exit Sub
			End If
		End If
		AlbumArtFile = EndSlash(FilePath) & "folder.jpg"
		If fso.FileExists(AlbumArtFile) Then
			If FitImage(AlbumArtFile) Then
				fso.CopyFile AlbumArtFile, aPath
				set fso = Nothing
				set files = Nothing
				Exit Sub
			End If
		End If
		' default "anonymous" albumart
		fso.CopyFile DCScriptFilesFolder & "\cover.jpg", aPath
	End If
	set fso = Nothing
	set files = Nothing
End Sub

Function FitImage(Img)
	FitImage=FALSE
	Dim ImDim : ImDim = ImageDimension(Img)
	Dim MaxImDim : MaxImDim = max(ImDim(0), ImDim(1))
	If (MaxImDim >= 200) AND (MaxImDim <= 350) Then FitImage = TRUE
End Function

Function OKtoOverwrite(aFile)
	Dim OverwriteFile : Set OverwriteFile = SDB.Objects("OverwriteFiles")
	OKtoOverwrite = TRUE
	' don't overwrite if file exists and overwrite is not allowed
	If fso.FileExists(aFile) And Not OverwriteFile.Checked Then OKtoOverwrite = FALSE
	'sdb.messagebox OKtoOverwrite, 2, array(4)
End Function

Sub CopyFiles(src, tgt)
	' copy m3u
	If OKtoOverwrite(tgt & ".list.m3u") Then fso.CopyFile src & ".list.m3u", tgt
	' copy png
	If OKtoOverwrite(tgt & ".icon.jpg") Then fso.CopyFile src & ".icon.jpg", tgt
	' copy dune_folder.txt
	If OKtoOverwrite(tgt & "dune_folder.txt") Then fso.CopyFile src & "dune_folder.txt", tgt
End Sub

Sub CopyFolderFiles(albumF, artistF)
	' Copy files from AlbumFolder to Artist (sub) Folder to have a nice icon on screen here as well.
	' Is the Default Icon Exists and a new one is present it will be overwritten.
	Dim CreateArtistIcon : CreateArtistIcon = FALSE
	Dim DefFile, TgtFile
	Dim UseIM : Set UseIM=SDB.Objects("DeepImageScan")

	If not fso.FileExists(artistF & ".icon.jpg") Then
		CreateArtistIcon = TRUE
	Else
		Set DefFile = fso.GetFile(DCScriptFilesFolder & "cover.jpg")
		Set TgtFile = fso.GetFile(artistF & ".icon.jpg")
		If (DefFile.Size = TgtFile.Size) AND (DefFile.DateLastModified = TgtFile.DateLastModified) Then
			CreateArtistIcon = TRUE
		End If
	End If
	
	If CreateArtistIcon Then
		If fso.FileExists(albumF & ".icon.jpg") Then
			Dim h2, w2, ScaleFactor
			If OKtoOverwrite(artistF & ".icon.jpg") Then fso.CopyFile albumF & ".icon.jpg", artistF
			If UseIM.Checked Then
				Dim img : Set img = CreateObject("ImageMagickObject.MagickImage.1")' Load ImageMagick
				w2 = img.Identify ("-format", "%w", albumF & ".icon.jpg")
				h2 = img.Identify ("-format", "%h", albumF & ".icon.jpg")
			Else
				Dim oPic : set oPic=loadpicture(albumF & ".icon.jpg")
				'height and width properties return in himetric (0.01mm)
				'numeric factors are just to convert them to pixel
				h2=round(oPic.height/2540*96)
				w2=round(oPic.width/2540*96)
				set oPic=nothing
			End If
			ScaleFactor = Round(350/Max(h2, w2),3)
			WriteDuneSubFolder EndSlash(artistF) & "dune_folder.txt", ScaleFactor
		Else
			If OKtoOverwrite(artistF & ".icon.jpg") Then fso.CopyFile DCScriptFilesFolder & ".icon.jpg", artistF
			If OKtoOverwrite(artistF & "dune_folder.txt") Then fso.CopyFile DCScriptFilesFolder & "SFdune_folder.txt", artistF & "dune_folder.txt"
		End If
	End If
End Sub

Sub CreateArtistFolderIcon(albumF, artistF)
	' Creates Artist Icon
	' If the Default Icon Exists and a new one is present it will be overwritten.
	Dim CreateArtistIcon : CreateArtistIcon = FALSE
	Dim DefFile, TgtFile
	Dim UseIM : Set UseIM=SDB.Objects("DeepImageScan")

	If not fso.FileExists(artistF & ".icon.jpg") Then
		CreateArtistIcon = TRUE
	Else
		Set DefFile = fso.GetFile(DCScriptFilesFolder & "cover.jpg")
		Set TgtFile = fso.GetFile(artistF & ".icon.jpg")
		If (DefFile.Size = TgtFile.Size) AND (DefFile.DateLastModified = TgtFile.DateLastModified) Then
			CreateArtistIcon = TRUE
		End If
	End If
	
	If CreateArtistIcon Then
		If fso.FileExists(albumF & ".icon.jpg") Then
			Dim h2, w2, ScaleFactor
			If OKtoOverwrite(artistF & ".icon.jpg") Then fso.CopyFile albumF & ".icon.jpg", artistF
			If UseIM.Checked Then
				Dim img : Set img = CreateObject("ImageMagickObject.MagickImage.1")' Load ImageMagick
				w2 = img.Identify ("-format", "%w", albumF & ".icon.jpg")
				h2 = img.Identify ("-format", "%h", albumF & ".icon.jpg")
			Else
				Dim oPic : set oPic=loadpicture(albumF & ".icon.jpg")
				'height and width properties return in himetric (0.01mm)
				'numeric factors are just to convert them to pixel
				h2=round(oPic.height/2540*96)
				w2=round(oPic.width/2540*96)
				set oPic=nothing
			End If
			ScaleFactor = Round(350/Max(h2, w2),3)
			WriteDuneSubFolder EndSlash(artistF) & "dune_folder.txt", ScaleFactor
		Else
			If OKtoOverwrite(artistF & ".icon.jpg") Then fso.CopyFile DCScriptFilesFolder & ".icon.jpg", artistF
			If OKtoOverwrite(artistF & "dune_folder.txt") Then fso.CopyFile DCScriptFilesFolder & "SFdune_folder.txt", artistF & "dune_folder.txt"
		End If
	End If
End Sub

Sub LoadAlbumArray' Loads all files into an Array
	Dim Tracks, i : Set Tracks = SDB.CurrentSongList
	Dim NumTracks : NumTracks = Tracks.Count
	Dim aSize
	For i = 0 to NumTracks-1
		Dim trk : Set trk = Tracks.Item(i)
		AddTrack trk
	Next
End Sub

Function SkipDrive(aPath)
	' remove driveletter and :\ of aPath
	' remove :\
	Dim A : A = Replace(aPath, ":\", "")
	' remove driveletter
	SkipDrive = Right(A,Len(A)-1)
End Function

Function SwapSlashes(aString)
	' Swap All backslashes to forwardslashes
	SwapSlashes = Replace(aString, "\", "/")
End Function

Function AllBackSlashes(aString)
	' Swap All backslashes to forwardslashes
	AllBackSlashes = Replace(aString, "/", "\")
End Function

Sub SortAlbumArray
	' Sort Main Array by Album and TrackNumber
	
	Dim maxFiles : maxFiles = UBound(arrAlbum,2)
	Dim LowBound, HighBound, albumindex, newalbum, i, lowerB
	
	' Sort by Album Name First
	LowBound = LBound(arrAlbum, 2)
	HighBound = UBound(arrAlbum, 2)
	QuickSortCol arrAlbum,LowBound,HighBound,3
	' albums with the same name are not identified (add artist/various/empty)
	' Question: how many albums exist with the same name and are indexed in the same run? ... Well?
	
	' Now Sort each album by tracknumber
	albumindex = 1 ' First album index
	newalbum = TRUE
	For i=0 to Ubound(arrAlbum, 2) - 1
		If newalbum Then
			lowerB = i
			newalbum = FALSE
		End If
		arrAlbum(6, i) = albumindex
		
		If (arrAlbum(3, i) <> arrAlbum(3, i+1)) Then
			newalbum = TRUE
			albumindex = albumindex + 1
			If i > lowerB Then
				QuickSortCol arrAlbum,lowerB,i,0
			End If
			If (i = maxFiles - 1) Then arrAlbum(6, i+1) = albumindex
		Else
			If (i = maxFiles - 1) Then
				arrAlbum(6, i+1) = albumindex
				QuickSortCol arrAlbum,lowerB,i+1,0
			End If
		End If
	Next
End Sub

Sub AddTrack(aTrack)
	' An "Intuitive" Array with fixed #columns and variable #rows CANNOT be REDIM'ed AND having its data PRESERVED.
	' SO: the data "matrix" is transposed
	'
	' 0 Disc.Tracknumber
	' 1 Title
	' 2 Artist
	' 3 AlbumName
	' 4 Year
	' 5 TrackPath
	' 6 "Album Index"
	' 7 AlbumArtist
	'
	Dim idxLast
	Dim discno : discno = ""
	Dim preZero : preZero = ""
	Dim i
	
	' [1] Retrieve the index number of the last element in the array
	On Error Resume Next
	idxLast = UBound(arrAlbum,2)
	If not Err = 0 Then
			idxLast = -1
			' This array is not empty.
			Err.Clear
	End If
	
	' [2] Resize the array, preserving the current content
	ReDim Preserve arrAlbum(7, idxLast + 1)
	' [3] Add the new element to the array
	if aTrack.DiscNumberStr <> "" Then discno = aTrack.DiscNumberStr & "."
	
	For i = 0 to 1 - len(aTrack.TrackOrderStr)
		preZero = preZero & "0"
	Next
	
	arrAlbum(0, idxLast + 1) = discno & preZero & aTrack.TrackOrderStr
	arrAlbum(1, idxLast + 1) = aTrack.Title
	
	' Artist
	If NOT aTrack.AlbumArtistName = "" Then
		arrAlbum(2, idxLast + 1) = FolderFix( SwapPrefix(aTrack.AlbumArtistName, "artist") )
	Else
		If aTrack.ArtistName = "" Then
			arrAlbum(2, idxLast + 1) = "Unknown"
		Else
			arrAlbum(2, idxLast + 1) = FolderFix( SwapPrefix(aTrack.ArtistName, "artist") )
		End If
	End If

	' See if AlbumName exists. If not, name it unknown
	If aTrack.AlbumName = "" Then
		arrAlbum(3, idxLast + 1) = "Unknown"
	Else
		arrAlbum(3, idxLast + 1) = FolderFix( SwapPrefix(aTrack.AlbumName, "album") )
	End If
	
	' See if Year exists
	If aTrack.Year = "" Then
		arrAlbum(4, idxLast + 1) = "0000"
	Else
		arrAlbum(4, idxLast + 1) = aTrack.Year
	End If
	
	arrAlbum(5, idxLast + 1) = aTrack.Path
	arrAlbum(7, idxLast + 1) = aTrack.AlbumArtistName
End Sub

Function FolderFix(aFolder)
	' checks folder on impossible names/characters
	Dim a : a = aFolder
	If Right(aFolder, 2) = ".." Then a = a & "_"' cannot end with a dot
	If Right(aFolder, 1) = "." Then a = Left(a,Len(a)-1)' a single dot will be removed
	a = Replace(a, "/", "_")' a slash here is a folder/subfolder separator elsewhere
	a = Replace(a, "?", "_")'
	a = Replace(a, ":", "_")' 
	a = Replace(a, "*", "_")
	a = Replace(a, "<", "[")
	a = Replace(a, ">", "]")
	a = Replace(a, "|", "_")
	a = Replace(a, """", "_")
	FolderFix = a
End Function

Function SwapPrefix(aName, aType)
	' swap make "The Beatles" look as "Beatles, The" and thus HAVE IT SORTED PROPERLY
	' for artists: take all "the,a,an,de,het,les,le,la"
	' for albums: maybe only The ...
	Dim aTmp : aTmp=aName
	IF aType="album" Then
		IF UCase(Left(aName,4))="THE " Then _
			aTmp = Right(aName,Len(aName)-4) & ", " & Left(aName, 3)
	ElseIF aType="artist" Then
		'The, A, An, De, Het, Een, Le, La, Les
		IF UCase(Left(aName,4))="THE " Then aTmp = Right(aName,Len(aName)-4) & ", " & Left(aName, 3)
		IF UCase(Left(aName,2))="A " Then aTmp = Right(aName,Len(aName)-2) & ", " & Left(aName, 1)
		IF UCase(Left(aName,3))="AN " Then aTmp = Right(aName,Len(aName)-3) & ", " & Left(aName, 2)
		IF UCase(Left(aName,3))="DE " Then aTmp = Right(aName,Len(aName)-3) & ", " & Left(aName, 2)
		IF UCase(Left(aName,4))="HET " Then aTmp = Right(aName,Len(aName)-4) & ", " & Left(aName, 3)
		IF UCase(Left(aName,4))="EEN " Then aTmp = Right(aName,Len(aName)-4) & ", " & Left(aName, 3)
		IF UCase(Left(aName,3))="LE " Then aTmp = Right(aName,Len(aName)-3) & ", " & Left(aName, 2)
		IF UCase(Left(aName,3))="LA " Then aTmp = Right(aName,Len(aName)-3) & ", " & Left(aName, 2)
		IF UCase(Left(aName,4))="LES " Then aTmp = Right(aName,Len(aName)-4) & ", " & Left(aName, 3)
	End IF
	SwapPrefix = aTmp
End Function

Sub Btn2Click
	' used in development only for quick reloading script
  SDB.Objects("Form1") = Nothing ' Remove the last reference to our form which also causes it to disappear
  Script.Reload("c:\Users\allart\AppData\Roaming\MediaMonkey\Scripts\DuneCatalog.vbs")
End Sub

Sub ButtonCancelClick
  SDB.Objects("Form1") = Nothing ' Remove the last reference to our form which also causes it to disappear
End Sub

Sub QuickSortCol(vec,loBound,hiBound,SortField)
	' copied from: http://www.4guysfromrolla.com/webtech/012799-3.shtml
	' rewritten to Columns instead of Rows by alveola
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
  '== This proc swaps two columns of an array 
  Dim x,tempvar
  For x = 0 to Ubound(ary,1)
    tempvar = ary(x, col1)    
    ary(x, col1) = ary(x, col2)
    ary(x, col2) = tempvar
  Next
End Sub  'SwapCols

Sub PrintAlbumArray(aTracks)
	Dim msg, idxLast, i
	msg = ""
	idxLast = UBound(aTracks,2)
	For i = 0 to idxLast
		msg = msg & "Alb:" & aTracks(6, i) & "Song:" & chr(9) & aTracks(0, i) & chr(9) & aTracks(1, i) & chr(9) & aTracks(2, i) & chr(9) & aTracks(3, i) & chr(13)
	Next
	SDB.MessageBox msg, mtInformation, Array(mbOk)
End Sub

Sub WriteAlbumFolderdft(filename, scalefactor)
	Dim dftfso : Set dftfso = fso.CreateTextFile(filename ,True, False) ' False creates ascii file, which Dune likes/needs
	Dim filecontent
	filecontent = _
	"icon_path=.icon.jpg" & chr(13) & chr(10) & _
	"icon_sel_path=.icon.jpg" & chr(13) & chr(10) & _
	"icon_scale_factor=" & Scalefactor & chr(13) & chr(10) & _
	"use_icon_view=yes" & chr(13) & chr(10) & _
	"icon_valign=center" & chr(13) & chr(10) & _
	"background_path=../../../.service/.listbackground.jpg" & chr(13) & chr(10) & _
	"paint_content_box_background=no" & chr(13) & chr(10) & _
	"paint_path_box=no" & chr(13) & chr(10) & _
	"paint_help_line=no" & chr(13) & chr(10) & _
	"paint_scrollbar=yes" & chr(13) & chr(10) & _
	"paint_captions=yes" & chr(13) & chr(10) & _
	"num_cols=2" & chr(13) & chr(10) & _
	"num_rows=10" & chr(13) & chr(10) & _
	"paint_icon_selection_box=yes" & chr(13) & chr(10) & _
	"paint_icons=yes" & chr(13) & chr(10) & _
	"caption_font_size=normal" & chr(13) & chr(10)
	
	dftfso.Write(filecontent)
	dftfso.Close ' Create DuneFolder.txt file
End Sub

Sub WriteDuneSubFolder(filename, scalefactor)
	REM Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	Dim dftfso : Set dftfso = fso.CreateTextFile(filename ,True, False) ' False creates ascii file, which Dune likes/needs
	Dim filecontent : filecontent = _
		"icon_path=" & ".icon.jpg" & chr(13) & chr(10) & _
		"icon_scale_factor=" & ScaleFactor & chr(13) & chr(10) & _
		"background_path=../../../.service/.listbackground.jpg" & chr(13) & chr(10) & _
		"use_icon_view=yes" & chr(13) & chr(10) & _
		"icon_valign=center" & chr(13) & chr(10) & _
		"background_x=0" & chr(13) & chr(10) & _
		"background_y=0" & chr(13) & chr(10) & _
		"content_box_x=0" & chr(13) & chr(10) & _
		"content_box_Y=0" & chr(13) & chr(10) & _
		"paint_path_box=no" & chr(13) & chr(10) & _
		"paint_help_line=no" & chr(13) & chr(10) & _
		"paint_scrollbar=yes" & chr(13) & chr(10) & _
		"num_cols=4" & chr(13) & chr(10) & _
		"num_rows=2" & chr(13) & chr(10) & _
		"paint_icon_selection_box=yes" & chr(13) & chr(10) & _
		"paint_captions=yes" & chr(13) & chr(10) & _
		"paint_icons=yes" & chr(13) & chr(10) & _
		"icon_top=7" & chr(13) & chr(10) & _
		"icon_bottom=100" & chr(13) & chr(10) & _
		"caption_font_size=normal" & chr(13) & chr(10)
	
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

Sub resizeImage(sourceFile, toWidth, toHeight, destinationFile)
	
	Dim imgWidth, imgHeight, img
	Dim xScale, yScale
	Dim newWidth, newHeight
	Dim Command
	
	' Load ImageMagick
	Set img = CreateObject("ImageMagickObject.MagickImage.1")
	
	' Get current image size
	imgWidth = img.Identify ("-format", "%w", sourceFile)
	imgHeight = img.Identify ("-format", "%h", sourceFile)
	
	' Calculate scale
	xScale = imgWidth / toWidth
	yScale = imgHeight / toHeight
	
	' Calculate new width and height
	if yScale > xScale then
		newWidth = round(imgWidth * (1/yScale))
		newHeight = round(imgHeight * (1/yScale))
	else
		newWidth = round(imgWidth * (1/xScale))
		newHeight = round(imgHeight * (1/xScale))
	end if

	' Run Convert to resize the image.
	Command = img.Convert("-resize", newWidth&"x"&newHeight&"!", sourceFile, destinationFile)
	
	set Command = nothing
	
end Sub

Sub MakeGlassBubble(sourceFile)
	Dim imgWidth, imgHeight, img
	Dim Command
	' Load ImageMagick
	Set img = CreateObject("ImageMagickObject.MagickImage.1")
	
	' Get current image size
	imgWidth = img.Identify ("-format", "%w", sourceFile)
	imgHeight = img.Identify ("-format", "%h", sourceFile)
	
	Dim wsh, FilePath
	Set wsh = CreateObject("Wscript.Shell")
	FilePath = FSO.GetParentFolderName(sourceFile)
	
	Dim tgt0, tgt1, tgt2, tgt3
	Dim DisplayedCornerDia : DisplayedCornerDia="60"
	' due to scaling the real cornerdia should be scaled as well
	Dim CD : CD = Round( DisplayedCornerDia * Max(imgWidth, imgHeight) / 350, 0)
	
	tgt1 = FilePath & "\" & "img_mask.png"
	tgt2 = FilePath & "\" & "img_lighting.png"
	tgt3 = FilePath & "\" & "img_target.png"
	
	Command = strConv & chr(34) & sourceFile & chr(34) _
		& " -alpha off -fill white -colorize 100% " _
		& " -draw ""fill black polygon 0,0 0," & CD & " " & CD & ",0 fill white circle " & CD & "," & CD & " " & CD & ",0""" _
		& "	( +clone -flip ) -compose Multiply -composite " _
		& " ( +clone -flop ) -compose Multiply -composite " _
		& " -background Gray50 -alpha Shape " & chr(34) & tgt1 & chr(34)
	
	wsh.run Command, 7, true
	
	Command = strConv & chr(34) & tgt1 & chr(34) _
		& " -bordercolor None -border 1x1 " _
		& " -alpha Extract -blur 0x10  -shade 130x30 -alpha On " _
		& " -background gray50 -alpha background -auto-level " _
		& " -function polynomial  3.5,-5.05,2.05,0.3 " _
		& " ( +clone -alpha extract  -blur 0x2 ) " _
		& " -channel RGB -compose multiply -composite " _
		& " +channel +compose -chop 1x1 " & chr(34) & tgt2 & chr(34)
	
	wsh.run Command, 7, true
	
	Command = strConv & chr(34) & sourceFile & chr(34) _
		& " -alpha Set " & chr(34) & tgt2 & chr(34) _
		& " ( -clone 0,1 -alpha Opaque -compose Hardlight -composite ) " _
		& "	-delete 0 -compose In -composite " & chr(34) & tgt3 & chr(34)
	
	wsh.run Command, 7, true

	' add crop to original dim (remove extra 1px right and bottom)
	Command = strConv & chr(34) & tgt3 & chr(34) _
		& " -crop " & imgWidth & "x" & imgHeight & "+0+0 " & chr(34) & tgt3 & chr(34)
	
	wsh.run Command, 7, true
	
	FSO.DeleteFile sourceFile
	FSO.MoveFile tgt3, sourceFile
	
	FSO.DeleteFile tgt1
	FSO.DeleteFile tgt2
	
	set Command = nothing
	
End Sub