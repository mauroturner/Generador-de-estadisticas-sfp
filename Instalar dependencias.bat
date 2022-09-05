@echo off
setlocal
:PROMPT
SET /P AREYOUSURE=Estas seguro de querer instalar las dependencias (S/[N])?
IF /I "%AREYOUSURE%" NEQ "S" GOTO END

@echo off
cd %mypath% 
pip install -r requirements.txt
set /p DUMMY=Presione enter para finalizar
:END
endlocal