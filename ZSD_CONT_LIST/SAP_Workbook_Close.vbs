SAP_Workbook = wscript.arguments(0)
on error resume next
do
 err.clear
 Set xclApp = GetObject(, "Excel.Application")
 If Err.Number = 0 Then exit do
 'msgbox "Wait for Excel session"
 wscript.sleep 5
 loop

do 
 err.clear
 Set xclwbk = xclApp.Workbooks.item(SAP_Workbook)
 If Err.Number = 0 Then exit do
 'msgbox "Wait for SAP workbook"
 wscript.sleep 5
loop

on error goto 0 

xclApp.Visible = True
xclapp.DisplayAlerts = false

xclwbk.Close False


Set xclwbk = Nothing
Set xclsheet = Nothing
if xclapp.Workbooks.Count=0 then xclapp.Quit
set xclapp = Nothing


