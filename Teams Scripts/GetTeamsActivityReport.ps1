##After a while of trying to work out how to use the PowerShell Graph SDK commands to pull Teams reports and failing, stumbled across some articles which talked about
##the request failing as the it expected a JSON response but it got a CSV file back instead...looks like it's still broken. So found out a slightly different way
##of doing it. Instead of using this: Get-MgReportTeamUserActivityUserDetail -Period D30 -Outfile TeamsReport.csv (which fails) instead use Invoke-MgGraphRequest 
##command which works fine. Code below which worked for me

##change period to suit, can be D7,D30,D90,D180. So period='Dx'
##change OutputFilePath to where you want report saved to

##note: same issue around response seems to affect a number of the Get-MgReportxxxxx commands. Can use the same Invoke request and just replace the URI with the 
##report you want

Connect-MgGraph -Scopes Reports.Read.All

Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D30')" -OutputFilePath C:\users\username\Documents\teamsreport.csv
