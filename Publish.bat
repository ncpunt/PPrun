Call Day.bat
set buildNumber=1.0.1.%jd%
rd /s/q bin\Release\net8.0-windows\win-x64\publish
dotnet publish -p:PublishSingleFile=true -r win-x64 -c Release --self-contained false /p:AssemblyVersion=%buildNumber% /p:Version=%buildNumber%
cd bin\Release\net8.0-windows\win-x64\publish
dir *.*
pause