@echo off

:: set subdir
IF NOT AA==A%1A set SF=%1\

:: Create decadefolders
call createdecade 0000 1949 %SF%
call createdecade 1950 1959 %SF%
call createdecade 1960 1969 %SF%
call createdecade 1970 1979 %SF%
call createdecade 1980 1989 %SF%
call createdecade 1990 1999 %SF%
call createdecade 2000 2009 %SF%
call createdecade 2010 2019 %SF%

