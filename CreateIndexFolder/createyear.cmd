:: Create year and decade images
@echo off

if AA==A%IMBIN%A set IMBIN="c:\Program Files (x86)\ImageMagick-6.8.0-Q16\"
if AA==A%BUTTONCOL%A set BUTTONCOL=RED
if AA==A%FONT%A set FONT=Showcard-Gothic
set PT=128
set TEXTLINE=%1

:: OuterCircle
%IMBIN%convert -size 320x320 xc:black ^
	-fill white -draw "circle    150,150   3,150" ^
	-gaussian 1x1 +matte ant_mask.png

%IMBIN%convert ant_mask.png -fill %BUTTONCOL% -draw "color 0,0 reset" ^
	ant_mask.png +matte  -compose CopyOpacity -composite ^
	-font %FONT%  -pointsize %pt%  -fill white -stroke black -strokewidth 2 ^
	-gravity Center  -annotate -5-4 "%textline%" ^
	ant.png

%IMBIN%convert ant.png  -alpha extract -blur 0x10  -shade 110x30  -normalize ^
	ant.png  -compose Overlay -composite ^
	ant.png  -matte  -compose Dst_In  -composite ^
	ant_3D.png

set DP=%2
IF AA==A%2A set DP=.
IF exist %DP% ( echo %DP% exists ) ELSE ( mkdir %DP% && echo %DP% created)

%IMBIN%convert ant_3D.png -trim +repage %DP%\.icon.png
