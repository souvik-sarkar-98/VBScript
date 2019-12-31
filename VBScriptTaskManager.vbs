'------------------------------------------------------------------
' This sample schedules a task to start on a daily basis.
'------------------------------------------------------------------
Option Explicit
Dim service,WshShell
Set WshShell = CreateObject("WScript.Shell")
Set service = CreateObject("Schedule.Service") ' Create the TaskService object.

Call Menu()

'***********************************************************************************
Function Menu()
	Select Case InputBox("CHOOSE AN ACTION "&vbCrLf&vbCrLf&"1-->> SCHEDULE NEW TASK" &vbCrLf&vbCrLf&"2--> VIEW EXISTING SCHEDULED TASK "&vbCrLf&vbCrLf& "3-->> DELETE EXISTING SCHEDULED TASK "&vbCrLf&vbCrLf& "4-->> EXIT ") 
		Case "1"
			Call setNewTask()
		Case "2"
			Call viewScheduleTask()
		Case "3"
			Call deleteScheduleTask()
		Case "4"
			Exit Function 
		Case Else
			wscript.echo "WRONG INPUT - PLEASE TRY AGAIN"
			Call Menu()
	End Select 
End Function

'***********************************************************************************
Function GenerateID()
  	Dim nLow,nHigh
 	nLow = 1000
    nHigh = 9999
 	Randomize
  	GenerateID=Int((nHigh - nLow + 1) * Rnd + nLow)
End Function 

'***********************************************************************************

'Function to check time
Function IsTime (str)
  if str = "" Then
    IsTime = false
  Else
    On Error Resume Next
    TimeValue(str)
    if Err.number = 0 Then
      IsTime = true
    Else
      IsTime = false
    end if
    On Error GoTo 0
  end if
end Function 

'***********************************************************************************

'Function to set a new Task
Function setNewTask()
	'Defining Variables
	Dim rootFolder,taskDefinition,regInfo,settings,triggers,trigger,Time,RunTime,startTime,endTime,Action,TaskName,TaskId,endDate,oExec,FileSelected,RunSpan,Text,repetitionPattern,reply,str
	const TriggerTypeDaily = 2
    const ActionTypeExec = 0 ' A constant that specifies an executable action.
	RunSpan=2
	call service.Connect()
	Set rootFolder = service.GetFolder("\") ' Get a folder to create a task definition in. 
	Set taskDefinition = service.NewTask(0) ' The flags parameter is 0 because it is not supported.
	TaskName=InputBox("ENTER TASK NAME :"&vbCrLf&vbCrLf&"OR ENTER '0' TO EXIT ")
	If TaskName="" Or TaskName="0" Then
		Exit Function 
	End If 
	TaskId=GenerateID()
	' Define information about the task.
	Set regInfo = taskDefinition.RegistrationInfo ' Set the registration info for the task by creating the RegistrationInfo object. 
	regInfo.Description = TaskName
	regInfo.Author = "Administrator"
	TaskName=TaskName&" ##ID## "&TaskId
    ' Set the task setting info for the Task Scheduler by
	Set settings = taskDefinition.Settings
	settings.Enabled = True
	settings.StartWhenAvailable = True
	settings.Hidden = False
    ' Create a daily trigger. Note that the start boundary 
	Set triggers = taskDefinition.Triggers
	Set trigger = triggers.Create(TriggerTypeDaily)
	str=""
	Do While Not IsTime(RunTime)
		RunTime=InputBox (str&"ENTER TIME : "&vbcrlf&"NOTE: PLEASE FOLLOW 24 HOUR FORMAT [HH:MM:SS]"&vbCrLf&vbCrLf)
		str="WRONG INPUT- TRY AGAIN"&vbCrLf&vbCrLf
	Loop 
	startTime = Year(Now)&"-"&Month(now)&"-"&Day(now)& "T"&RunTime 
	endDate=DateAdd("YYYY",RunSpan,Date())
	endTime= Year(endDate)&"-"&Month(endDate)&"-"&Day(endDate)& "T"&RunTime 
	trigger.StartBoundary = startTime
	trigger.EndBoundary = endTime
	trigger.DaysInterval = 1    'Task runs every day.
	trigger.Id = TaskName&"Id"&TaskId
	trigger.Enabled = True
	Text="WOULD YOU LIKE TO REPEAT THE TASK FOR CERTAIN DURATION OF TIME?"&vbCrLf&vbCrLf&"[MAXIMUM DURATION IS 24 HOURS	MINIMUM DURATION IS 1 MINUTE]"&vbCrLf&vbCrLf&"1-->YES"&vbCrLf&"0-->NO"
	reply=TakeNCheckInput(Text,"0","1","1","","","","")
	If reply="1" And Not reply="0" Then 
		Set repetitionPattern = trigger.Repetition
		Text="ENTER THE REPETATION DURATION FROM STARTING TIME IN HOUR OR MINUTE : "&vbCrLf&vbCrLf&"[MAXIMUM DURATION IS 24 HOURS	MINIMUM DURATION IS 1 MINUTE]"&vbCrLf&vbCrLf&"[ UNIT: HOUR--> H		MINUTE--> M ]"
		reply=TakeNCheckInput(Text,"H","M","0","24","1","59","0")
		repetitionPattern.Duration = "PT"&reply
		Text="ENTER TIME GAP: "&vbCrLf&vbCrLf&"[MAXIMUM GAP IS 24 HOURS	MINIMUM GAP IS 1 MINUTE]"&vbCrLf&vbCrLf&"[ UNIT: HOUR--> H		MINUTE--> M ]"
		reply=TakeNCheckInput(Text,"H","M","0","24","1","59","0")
		repetitionPattern.Interval = "PT"&reply
	End If 
	Set Action = taskDefinition.Actions.Create( ActionTypeExec )
	Text="CHOOSE AN OPTION "&vbCrLf&vbCrLf&"1-->SELECT A FILE FOR EXECUTION "&vbCrLf&vbCrLf&"2--> INPUT A COMMAND TO EXECUTE"
	reply=TakeNCheckInput(Text,"1","2","1","","","","")
	If reply="1" Then 
		Set oExec=WshShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
		FileSelected = oExec.StdOut.ReadLine
	ElseIf reply="2" Then 
		FileSelected=InputBox("PLEASE PASTE THE COMMAND: ")
	End If 
	Action.Path = FileSelected
	call rootFolder.RegisterTaskDefinition( _ 
	TaskName, taskDefinition, 6, , , 3)
	WshShell.popup "TASK SUBMITTED.",2,"ALERT"
End Function 

Function TakeNCheckInput(ByVal text,ByVal ip0, ByVal ip1,ByVal adj,ByVal a0,ByVal b0,ByVal a1,ByVal b1)
	Dim str,bool,rep	
	str=""
	bool=True 
	Do While bool
		rep=UCase(InputBox (str&text))
		If Mid(rep,Len(rep),1)=ip0 Then 
			If  IsNumeric(Mid(rep,1,Len(rep)-1+adj)) And Mid(rep,1,Len(rep)-1)<=a0 And Mid(rep,1,Len(rep)-1)>=b0 Then  
				bool= False 
			Else
				bool= True 
			End If 
		ElseIf 	Mid(rep,Len(rep),1)=ip1  Then  
			If IsNumeric(Mid(rep,1,Len(rep)-1+adj)) And Mid(rep,1,Len(rep)-1)<=a1 And Mid(rep,1,Len(rep)-1)>=b1 Then 
				bool= False 
			Else 
					bool= True 
			End If 
		Else 
			bool= True  
			str="WRONG INPUT-CHECK UNIT AND TRY AGAIN"&vbCrLf&vbCrLf
		End If 
	Loop
	TakeNCheckInput=rep
End Function 

'***********************************************************************************
Function viewScheduleTask()
	Dim c,rootFolder,tasks,A,Task,l,num,v
	c=0
	l=1
	num=10
	service.Connect()
	Set rootFolder = service.GetFolder("\")
	Set tasks = rootFolder.GetTasks(0)
	If tasks.Count = 0 Then 
    	Wscript.Echo "NO TASKS ARE REGISTERED."
	Else
	    A="NUMBER OF TASK REGISTERED: " & tasks.Count &vbCrLf&vbCrLf
	    
	    For Each Task In tasks
	    	c=c+1
	    	A = A & c &" -->"&vbCrLf&"TASK NAME :  "& Task.Name &vbCrLf&"NEXT RUN TIME :  " & Task.NextRunTime &vbCrLf&vbCrLf
	    	If c=(num*l) Then 
		    	l=l+1
		    	WScript.echo A
		    	A=""
	    	End If 
	    Set 	v=Task
	    Next
	    If c < num Or Not (tasks.Count/num =0) Then  
	   	 	wscript.echo A
	    End If
	End If
	Call Menu()
End Function 


'***********************************************************************************

Function deleteScheduleTask()
	On Error Resume Next
	Dim objTaskFolder,colTasks,tempName,objTask,TaskID,Cnf
	Cnf="GFDJMK"
	Call service.Connect()
	Set objTaskFolder = service.GetFolder("\")
	Set colTasks = objTaskFolder.GetTasks(0)
	TaskID=InputBox ("Please Enter Task ID OR NAME : ")
	For Each objTask In colTasks
	 	With objTask
		    If InStr(objTask.Name,TaskID) And Not TaskID="" Then
		    	Cnf=InputBox("DO YOU WANT TO DELETE '"&objTask.Name&"' ?"&vbCrLf&vbCrLf&"1 --> YES"&vbCrLf&"0 --> NO")
		   		If Cnf="1" Then 
		   		 	tempName=objTask.Name
		      		objTask.Stop(0) 
		      		WScript.Sleep(1000)
		      		objTask.Enabled = False 
		      		objTaskFolder.DeleteTask objTask.Name,0  
    				if Err.number = 0 Then
		      			WshShell.Popup ""&tempName&" SUCCESSFULLY DELETED",2,"ALERT"
		      		Else 
		      			WshShell.Popup Err.Description ,2,"ALERT"
		      		End If 
		      	End If
   			 End If
  		End With
	Next
	If Cnf="GFDJMK" Then 
		  WshShell.Popup "'"&TaskID&"' NOT FOUND",1,"ALERT"
	End If 
End Function 