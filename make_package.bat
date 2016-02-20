REM Build Solution
set PATH_SOURCE_PROJ="%cd%\TranslationLib.vbproj"

FOR %%C IN (Release20 Release40) DO (
 "C:\Program Files (x86)\MSBuild\14.0\Bin\MSBuild.exe" %PATH_SOURCE_PROJ% /p:Configuration=%%C
 if errorlevel 1 goto :eof
)

nuget pack %PATH_SOURCE_PROJ% -Prop Configuration=Release20 -IncludeReferencedProjects -Verbosity detailed
