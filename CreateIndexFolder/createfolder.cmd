:: Create text in a Bubble of any shape

@echo off

if AA==A%IMBIN%A set IMBIN="c:\Program Files (x86)\ImageMagick-6.8.0-Q16\"
if AA==A%BUTTONCOL%A set BUTTONCOL=RED
if AA==A%FONT%A set FONT=Showcard-Gothic
set PT=82
set TEXTLINE=%1

REM :: Circle
REM %IMBIN%convert -size 320x180 xc:black ^
	REM -fill white -draw "circle    160,80   83,80" ^
	REM -gaussian 1x1 +matte ant_mask.png

REM :: Ellipse
REM %IMBIN%convert -size 320x180 xc:black ^
	REM -fill white -draw "ellipse 160,80 157,77 0,360" ^
	REM -gaussian 1x1 +matte ant_mask.png

:: Ellipse
REM %IMBIN%convert -size 320x180 xc:black ^
	REM -fill white -draw "ellipse 160,80 120,60 0,360" ^
	REM -gaussian 1x1 +matte ant_mask.png

REM :: Empty
REM %IMBIN%convert -size 320x180 xc:black ^
	REM -gaussian 1x1 +matte ant_mask.png

REM :: Bean	
REM %IMBIN%convert -size 320x180 xc:black ^
	REM -fill white -draw "circle    80,80   3,80" ^
		REM -draw "circle   220,80 297,80" ^
		REM -draw "rectangle 80,3  217,157" ^
	REM -gaussian 1x1 +matte ant_mask.png

REM :: Rounded Rectangle
%IMBIN%convert -size 380x230 xc:black ^
	-fill white -draw "roundrectangle 3,3 377,227 180,100" ^
	-gaussian 1x1 +matte ant_mask.png

%IMBIN%convert ant_mask.png -fill %BUTTONCOL% -draw "color 0,0 reset" ^
	ant_mask.png +matte  -compose CopyOpacity -composite ^
	-font %FONT%  -pointsize %PT%  -fill white -stroke black -strokewidth 2 ^
	-gravity Center  -annotate +0+10 "%textline%" ^
	ant.png

%IMBIN%convert ant.png  -alpha extract -blur 0x10  -shade 110x30  -normalize ^
	ant.png  -compose Overlay -composite ^
	ant.png  -matte  -compose Dst_In  -composite ^
	ant_3D.png

set DP=%2
IF AA==A%2A set DP=.
IF exist %DP% ( echo %DP% exists ) ELSE ( mkdir %DP% && echo %DP% created)

%IMBIN%convert ant_3D.png -trim +repage %DP%\.icon.png
