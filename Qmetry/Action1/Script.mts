
Call TestCase1()
Call CloseLog()


Function CreateLog()
Dim TRFSO, MyTRFile,TRFileName
TRFileName = "C:\Users\C49563\Output.xml"
Set TRFSO=CreateObject("Scripting.FileSystemObject")
Set MyTRFile = TRFSO.CreateTextFile(TRFileName)
MyTRFile.WriteLine("")
MyTRFile.WriteLine("")
MyTRFile.close
Set TRFSO=Nothing
End Function

'STEP-2  call log function everytime when you want to run test 
Function WriteLog(Msg)
Dim TRFSO, MyTRFile,TRFileName
TRFileName = "C:\Users\C49563\Output.xml"
Set TRFSO=CreateObject("Scripting.FileSystemObject")
Set MyTRFile=TRFSO.OpenTextFile(TRFileName,8,True)
MyTRFile.WriteLine(Msg)
MyTRFile.close
Set TRFSO=Nothing
End Function


'STEP-3  Close log at end of the test and only once
Function CloseLog()
Dim TRFSO, MyTRFile,TRFileName
TRFileName = "C:\Users\C49563\Output.xml"
Set TRFSO=CreateObject("Scripting.FileSystemObject")
Set MyTRFile=TRFSO.OpenTextFile(TRFileName,8,True)
MyTRFile.WriteLine("")
MyTRFile.close
Set TRFSO=Nothing
End Function




'How to call these functions ( create and close file will remain in script)

Function TestCase1()
Call CreateLog()

On Error Resume Next
If Browser("name:=.*").Page("title:=.*").WebElement("innertext:=Done").Exist Then
	WriteLog("Test Case 1 : Passed")
Else
	WriteLog("Test Case 1 : Failed")

'do something more and the write in log to post in output.xml

End If

End Function
