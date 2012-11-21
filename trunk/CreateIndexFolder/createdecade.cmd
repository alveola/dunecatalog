:: Create year and decade images
@echo off

if AA==A%IMBIN%A set IMBIN="c:\Program Files (x86)\ImageMagick-6.8.0-Q16\"
if AA==A%BUTTONCOL%A set BUTTONCOL=RED
if AA==A%FONT%A set FONT=Showcard-Gothic
set PT=128

set YEAR1=%1
set YEAR2=%2

:: Rounded Rectangle
%IMBIN%convert -size 440x300 xc:black ^
	-fill white -draw "roundrectangle 3,3 437,297 120,90" ^
	-gaussian 1x1 +matte ant_mask.png

%IMBIN%convert ant_mask.png -fill %BUTTONCOL% -draw "color 0,0 reset" ^
	ant_mask.png +matte  -compose CopyOpacity -composite ^
	-font %FONT%  -pointsize %PT%  -fill white -stroke black -strokewidth 2 ^
	-gravity Center  -annotate +0+10 "%YEAR1%\n%YEAR2%" ^
	ant.png

%IMBIN%convert ant.png  -alpha extract -blur 0x10  -shade 110x30  -normalize ^
	ant.png  -compose Overlay -composite ^
	ant.png  -matte  -compose Dst_In  -composite ^
	ant_3D.png

IF AA==A%3A SET SF=%3
SET DP=%SF%\%YEAR1%-%YEAR2%
IF EXIST %DP% ( ECHO %DP% exists ) ELSE ( mkdir %DP% && ECHO %DP% created)

%IMBIN%convert ant_3D.png -trim +repage %DP%\.icon.png

