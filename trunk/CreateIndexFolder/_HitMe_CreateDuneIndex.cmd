@echo off

call :StartTimer

setlocal

:: Creating Dune Index Folders
:: ===========================
:: 1. Create folders and icons
:: 2. Create textfiles
:: 3. Clean up

:: set the location of your ImageMagick folder here (it contains convert.exe).
set IMBIN="c:\Program Files (x86)\ImageMagick-6.8.0-Q16\"
set MAIN=DuneIndex

:: 1. Folders & Icons :::::::::::::::::::::::::::::::::::::::::::::

::create albums
echo Create Album Folders and Icons
set BUTTONCOL=OliveDrab3
call createalldirs %MAIN%\Albums\
call createallletters %MAIN%\Albums\
call createfolder Albums %MAIN%\Albums

::create artists
echo Create Artist Folders and Icons
set BUTTONCOL=FireBrick
call createalldirs %MAIN%\Artists
call createallletters %MAIN%\Artists
call createfolder Artists %MAIN%\Artists

::create decades
echo Create Decade Folders and Icons
set BUTTONCOL=SteelBlue3
call createallyeardirs %MAIN%\Years		
call createalldecades %MAIN%\Years
call createfolder Years %MAIN%\Years

::create years
echo Create Year Folders and Icons
call createallyears %MAIN%\Years

:: Some specials
call createspecial *?#@ %MAIN%\Years\other 
::call createyear 0000 %SF%\0000-1949\0000

::create tracks
echo Create Artist Folders and Icons
set BUTTONCOL=Gold2
call createalldirs %MAIN%\Tracks
call createallletters %MAIN%\Tracks
call createfolder Tracks %MAIN%\Tracks

:: 2. Dune_folder.txt :::::::::::::::::::::::::::::::::::::::::::::

ECHO Copy Dune_folder.txt for Albums
COPY 1dune_folder.txt %MAIN%\Albums\01_A\dune_folder.txt
CALL copyABCtxt.cmd %MAIN%\Albums
COPY albumsdune_folder.txt %MAIN%\Albums\dune_folder.txt

COPY 2dune_folder.txt %MAIN%\Artists\01_A\dune_folder.txt
CALL copyABCtxt.cmd %MAIN%\Artists
COPY artistsdune_folder.txt %MAIN%\Artists\dune_folder.txt

COPY 3dune_folder.txt %MAIN%\Years\0000-1949\1910\dune_folder.txt
COPY decadesdune_folder.txt %MAIN%\Years\0000-1949\dune_folder.txt
CALL copyYEARtxt.cmd %MAIN%\Years
COPY yearsdune_folder.txt %MAIN%\Years\dune_folder.txt

COPY 4dune_folder.txt %MAIN%\Tracks\01_A\dune_folder.txt
CALL copyABCtxt.cmd %MAIN%\Tracks
COPY tracksdune_folder.txt %MAIN%\Tracks\dune_folder.txt

COPY MAINdune_folder.txt %MAIN%\dune_folder.txt

MKDIR %MAIN%\.service
ATTRIB -h .listbackground.jpg
COPY .listbackground.jpg %MAIN%\.service
ATTRIB -h .empty.png
COPY .empty.png %MAIN%\.service

:: 3. Clean up ::::::::::::::::::::::::::::::::::::::::::::::::::::

del ant.png
del ant_3D.png
del ant_mask.png

endlocal

call :StopTimer
call :DisplayTimerResult

goto :EOF

:StartTimer
:: Store start time
set StartTIME=%TIME%
for /f "usebackq tokens=1-4 delims=:., " %%f in (`echo %StartTIME: =0%`) do set /a Start100S=1%%f*360000+1%%g*6000+1%%h*100+1%%i-36610100
goto :EOF

:StopTimer
:: Get the end time
set StopTIME=%TIME%
for /f "usebackq tokens=1-4 delims=:., " %%f in (`echo %StopTIME: =0%`) do set /a Stop100S=1%%f*360000+1%%g*6000+1%%h*100+1%%i-36610100
:: Test midnight rollover. If so, add 1 day=8640000 1/100ths secs
if %Stop100S% LSS %Start100S% set /a Stop100S+=8640000
set /a TookTime=%Stop100S%-%Start100S%
set TookTimePadded=0%TookTime%
goto :EOF

:DisplayTimerResult
:: Show timer start/stop/delta
echo Started: %StartTime%
echo Stopped: %StopTime%
echo Elapsed: %TookTime:~0,-2%.%TookTimePadded:~-2% seconds
goto :EOF