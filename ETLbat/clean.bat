@echo off
color a
echo 因sqllog文件过多，调此脚本删除多余的sqllog文件
D:
cd CRMTOOL\log
set name=%date:~0,4%
set name=%name:/=%
REM --删除去年的记录
set /a name=%name%-1
del /f %name%?*.*
REM --删除前年的记录
set /a name=%name%-1
del /f %name%?*.*
echo 删除成功
pause
