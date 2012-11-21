@echo off

:: set subdir
set SF=.\
IF NOT AA==A%1A set SF=%1\

:: create icons
call createletter A %SF%01_A
call createletter B %SF%02_B
call createletter C %SF%03_C
call createletter D %SF%04_D
call createletter E %SF%05_E
call createletter F %SF%06_F
call createletter G %SF%07_G
call createletter H %SF%08_H
call createletter I %SF%09_I
call createletter J %SF%10_J
call createletter K %SF%11_K
call createletter L %SF%12_L
call createletter M %SF%13_M
call createletter N %SF%14_N
call createletter O %SF%15_O
call createletter P %SF%16_P
call createletter Q %SF%17_Q
call createletter R %SF%18_R
call createletter S %SF%19_S
call createletter T %SF%20_T
call createletter U %SF%21_U
call createletter V %SF%22_V
call createletter W %SF%23_W
call createletter X %SF%24_X
call createletter Y %SF%25_Y
call createletter Z %SF%26_Z
call createletter # %SF%27_#
call createletter - %SF%28_-
