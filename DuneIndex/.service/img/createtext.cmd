:: Create year and decade images
@echo off
set IMbin="c:\Program Files (x86)\ImageMagick-6.8.0-Q16\"

REM set font=Candara-Bold
set font=Showcard-Gothic
set pt=%2
set pt=80

set textline=%1
REM del %outputfile%

REM :: Circle
REM %IMbin%convert -size 320x180 xc:black ^
	REM -fill white -draw "circle    160,80   83,80" ^
	REM -gaussian 1x1 +matte ant_mask.png

REM :: Ellipse
REM %IMbin%convert -size 320x180 xc:black ^
	REM -fill white -draw "ellipse 160,80 157,77 0,360" ^
	REM -gaussian 1x1 +matte ant_mask.png

:: Ellipse
%IMbin%convert -size 320x180 xc:black ^
	-fill white -draw "ellipse 160,80 120,60 0,360" ^
	-gaussian 1x1 +matte ant_mask.png

REM :: Empty
REM %IMbin%convert -size 320x180 xc:black ^
	REM -gaussian 1x1 +matte ant_mask.png

REM :: Bean	
REM %IMbin%convert -size 320x180 xc:black ^
	REM -fill white -draw "circle    80,80   3,80" ^
		REM -draw "circle   220,80 297,80" ^
		REM -draw "rectangle 80,3  217,157" ^
	REM -gaussian 1x1 +matte ant_mask.png

REM :: Rounded Rectangle
REM %IMbin%convert -size 320x180 xc:black ^
	REM -fill white -draw "roundrectangle 3,3 317,157 100,60" ^
	REM -gaussian 1x1 +matte ant_mask.png

%IMbin%convert ant_mask.png -fill red -draw "color 0,0 reset" ^
	ant_mask.png +matte  -compose CopyOpacity -composite ^
	-font %font%  -pointsize %pt%  -fill white -stroke black -strokewidth 2 ^
	-gravity Center  -annotate -5-4 "%textline%" ^
	ant.png

%IMbin%convert ant.png  -alpha extract -blur 0x6  -shade 110x30  -normalize ^
	ant.png  -compose Overlay -composite ^
	ant.png  -matte  -compose Dst_In  -composite ^
	ant_3D.png

REM set dp=j:\_Music\Years\1960-1969\
REM set dp=.
REM set dp=%2
REM IF exist %dp%\%textline% ( echo %dp%\%textline% exists ) ELSE ( mkdir %dp%\%textline% && echo %dp%\%textline% created)

%IMbin%convert ant_3D.png -trim +repage %dp%123.png

REM convert ori.png -channel Alpha -evaluate Divide 2 output.png
