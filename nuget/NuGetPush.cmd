setlocal
set PackageVersion=1.7.3.0

nuget.exe push NetOffice.Core.%PackageVersion%.nupkg
nuget.exe push NetOffice.Core.Net20.%PackageVersion%.nupkg
nuget.exe push NetOffice.Core.Net30.%PackageVersion%.nupkg
nuget.exe push NetOffice.Core.Net35.%PackageVersion%.nupkg
nuget.exe push NetOffice.Core.Net40.%PackageVersion%.nupkg
nuget.exe push NetOffice.Core.Net45.%PackageVersion%.nupkg

nuget.exe push NetOffice.Excel.%PackageVersion%.nupkg
nuget.exe push NetOffice.Excel.Net20.%PackageVersion%.nupkg
nuget.exe push NetOffice.Excel.Net30.%PackageVersion%.nupkg
nuget.exe push NetOffice.Excel.Net35.%PackageVersion%.nupkg
nuget.exe push NetOffice.Excel.Net40.%PackageVersion%.nupkg
nuget.exe push NetOffice.Excel.Net45.%PackageVersion%.nupkg

nuget.exe push NetOffice.Word.%PackageVersion%.nupkg
nuget.exe push NetOffice.Word.Net20.%PackageVersion%.nupkg
nuget.exe push NetOffice.Word.Net30.%PackageVersion%.nupkg
nuget.exe push NetOffice.Word.Net35.%PackageVersion%.nupkg
nuget.exe push NetOffice.Word.Net40.%PackageVersion%.nupkg
nuget.exe push NetOffice.Word.Net45.%PackageVersion%.nupkg

nuget.exe push NetOffice.Outlook.%PackageVersion%.nupkg
nuget.exe push NetOffice.Outlook.Net20.%PackageVersion%.nupkg
nuget.exe push NetOffice.Outlook.Net30.%PackageVersion%.nupkg
nuget.exe push NetOffice.Outlook.Net35.%PackageVersion%.nupkg
nuget.exe push NetOffice.Outlook.Net40.%PackageVersion%.nupkg
nuget.exe push NetOffice.Outlook.Net45.%PackageVersion%.nupkg

nuget.exe push NetOffice.PowerPoint.%PackageVersion%.nupkg
nuget.exe push NetOffice.PowerPoint.Net20.%PackageVersion%.nupkg
nuget.exe push NetOffice.PowerPoint.Net30.%PackageVersion%.nupkg
nuget.exe push NetOffice.PowerPoint.Net35.%PackageVersion%.nupkg
nuget.exe push NetOffice.PowerPoint.Net40.%PackageVersion%.nupkg
nuget.exe push NetOffice.PowerPoint.Net45.%PackageVersion%.nupkg

nuget.exe push NetOffice.Access.%PackageVersion%.nupkg
nuget.exe push NetOffice.Access.Net20.%PackageVersion%.nupkg
nuget.exe push NetOffice.Access.Net30.%PackageVersion%.nupkg
nuget.exe push NetOffice.Access.Net35.%PackageVersion%.nupkg
nuget.exe push NetOffice.Access.Net40.%PackageVersion%.nupkg
nuget.exe push NetOffice.Access.Net45.%PackageVersion%.nupkg

nuget.exe push NetOffice.MSProject.%PackageVersion%.nupkg
nuget.exe push NetOffice.MSProject.Net20.%PackageVersion%.nupkg
nuget.exe push NetOffice.MSProject.Net30.%PackageVersion%.nupkg
nuget.exe push NetOffice.MSProject.Net35.%PackageVersion%.nupkg
nuget.exe push NetOffice.MSProject.Net40.%PackageVersion%.nupkg
nuget.exe push NetOffice.MSProject.Net45.%PackageVersion%.nupkg

nuget.exe push NetOffice.Visio.%PackageVersion%.nupkg
nuget.exe push NetOffice.Visio.Net20.%PackageVersion%.nupkg
nuget.exe push NetOffice.Visio.Net30.%PackageVersion%.nupkg
nuget.exe push NetOffice.Visio.Net35.%PackageVersion%.nupkg
nuget.exe push NetOffice.Visio.Net40.%PackageVersion%.nupkg
nuget.exe push NetOffice.Visio.Net45.%PackageVersion%.nupkg

pause
