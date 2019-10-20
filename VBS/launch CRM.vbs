Dim strTime, objIexplorer
Const run_time = "11:07:00"
Const link = "https://secure-ausomxeia.crmondemand.com/OnDemand/user/ReportIFrameView?SAWDetailViewURL=saw.dll?Go%26Path%3D%252fusers%252femr-ema%2523dan.petran%252fUsage%2bReports%252fuser%2bsign-in%2bhistory%2btracking&AnalyticReportName=user+sign-in+history+tracking"
Set objIexplorer = CreateObject("internetexplorer.application")

strTime = Time()

While TimeValue(strTime) <> TimeValue(run_time)
                WScript.Sleep 1000
                strTime = Time()
Wend
objIexplorer.Visible = True
objIexplorer.Navigate link
WScript.Quit
