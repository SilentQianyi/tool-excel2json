echo off

::excel所在的目录
set PATH_EXCEL="%~dp0/../Excel"
set EXE_GENERATOR="%~dp0/excel2json.exe"

::需要自己定义客户端输出目录
set TEMPLATE_OUT_PATH="%~dp0\tables"

%EXE_GENERATOR% %PATH_EXCEL% %TEMPLATE_OUT_PATH%

pause