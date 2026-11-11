@echo off
setlocal
pushd "%~dp0"
cl /nologo /O2 /GS- /LD tablo.c /link /NOENTRY /NODEFAULTLIB /DEF:tablo.def /OUT:TabloOkuma.dll
set "sonuc=%errorlevel%"
popd
exit /b %sonuc%
