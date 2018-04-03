@echo off
cd C:\users\administrator\desktop\HZEncrypt
start HZEncrypt.exe
echo 请继续执行kettle转换
pause
rem start pan /file F:\ETL文件\每月5号.ktr
pan /file F:\ETL文件\每月5号.ktr
echo 执行成功，请解密文件，按任意键移动文件
pause
move /y  C:\users\administrator\desktop\HZEncrypt\每月5号.xlsx E:\HaoZeCRM\每月5号\
move /y C:\users\administrator\desktop\HZencrypt\解密模板.xls E:\HaoZeCRM\每月5号\
echo 移动文件成功
pause
