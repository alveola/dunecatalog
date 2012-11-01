:: Create year and decade images
@echo off
set IMbin="c:\Program Files (x86)\ImageMagick-6.8.0-Q16\"

REM set font=Candara-Bold
set font=Showcard-Gothic
set pt=128
REM set w_=1000
REM set h_=300

set textline=%1
set n2=%2
set textline=0000
set n2=%2
REM del %2

REM :: InnerCircle
REM %IMbin%convert -size 320x180 xc:black ^
	REM -fill white -draw "circle    160,80   83,80" ^
	REM -gaussian 1x1 +matte ant_mask.png

REM :: OuterCircle
REM %IMbin%convert -size 320x320 xc:black ^
	REM -fill white -draw "circle    150,150   3,150" ^
	REM -gaussian 1x1 +matte ant_mask.png

REM :: Ellipse
REM %IMbin%convert -size 320x180 xc:black ^
	REM -fill white -draw "ellipse 160,80 157,77 0,360" ^
	REM -gaussian 1x1 +matte ant_mask.png

REM :: Empty
REM %IMbin%convert -size 320x180 xc:black ^
	REM -gaussian 1x1 +matte ant_mask.png

REM :: Bean	
REM %IMbin%convert -size 320x180 xc:black ^
	REM -fill white -draw "circle    80,80   3,80" ^
		REM -draw "circle   220,80 297,80" ^
		REM -draw "rectangle 80,3  217,157" ^
	REM -gaussian 1x1 +matte ant_mask.png

:: Rounded Rectangle
%IMbin%convert -size 440x300 xc:black ^
	-fill white -draw "roundrectangle 3,3 437,297 120,90" ^
	-gaussian 1x1 +matte ant_mask.png

REM :: Rounded Rectangle
REM %IMbin%convert -size 320x180 xc:black ^
	REM -fill white -draw "roundrectangle 3,3 317,157 100,60" ^
	REM -gaussian 1x1 +matte ant_mask.png

::



%IMbin%convert ant_mask.png -fill red -draw "color 0,0 reset" ^
	ant_mask.png +matte  -compose CopyOpacity -composite ^
	-font %font%  -pointsize %pt%  -fill white -stroke black -strokewidth 2 ^
	-gravity Center  -annotate +0+10 "%textline%" ^
	ant.png

REM %IMbin%convert ant_mask.png -fill white -stroke black -font %font% -pointsize %pt% ^
          REM -gravity center    label:"ImageMagick\nExamples\nby Anthony" ^
          REM label_centered.gif
					
REM %IMbin%convert ant_mask.png -fill red -draw "color 0,0 reset" ^
	REM ant_mask.png +matte  -compose CopyOpacity -composite ^
	REM -font %font%  -pointsize %pt%  -fill white -stroke black -strokewidth 2 ^
	REM -gravity Center label:"ImageMagick\nExamples\nby Anthony" ant.png

%IMbin%convert ant.png  -alpha extract -blur 0x6  -shade 110x30  -normalize ^
	ant.png  -compose Overlay -composite ^
	ant.png  -matte  -compose Dst_In  -composite ^
	ant_3D.png

set dp=j:\_Music\Years\1960-1969\
set dp=%textline%-%n2%
set dp=.
IF exist %dp% ( echo %dp% exists ) ELSE ( mkdir %dp% && echo %dp% created)

%IMbin%convert ant_3D.png -trim +repage %dp%\.icon.png
	