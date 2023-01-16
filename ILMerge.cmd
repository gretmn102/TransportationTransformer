@echo off
:: this script needs https://www.nuget.org/packages/ilmerge

:: set your target executable name (typically [projectname].exe)
SET APP_NAME=TransportationTransformer.exe

SET ILMERGE_VERSION=3.0.41

SET ILMerge=%USERPROFILE%\.nuget\packages\ilmerge\%ILMERGE_VERSION%\tools\net452\ILMerge.exe

ECHO Merging %APP_NAME% ...

%ILMerge% /wildcards /out:%APP_NAME% src\bin\Release\net461\%APP_NAME% src\bin\Release\net461\*.dll

:Done
for %%a in (%APP_NAME%) do (
    set file=%%~na
)
DEL %file%.pdb
MKDIR deploy
MOVE /Y %APP_NAME% deploy\%APP_NAME%
