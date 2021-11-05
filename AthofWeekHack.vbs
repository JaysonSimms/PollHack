'Find the polldaddy link on the web page
'Determine the option ID for the person to be chosen
'Works with any current IE

voteCount = 0
votes = 35
totalCount = 0	

Sub WaitForLoad
 Do While ie.Busy
  Wscript.Sleep 100
 Loop
End Sub


Sub VoteButton
	For Each Button In ie.Document.getElementsByTagName("a") 'Get the vote button
		If InStr(Button.getAttribute("href"), "vote") Then
			Button.Click  'click when object with an href of vote is found
			Call WaitForLoad 'make sure it completes.
			Exit For
		End If
	Next
End Sub


Do
	do while voteCount < votes	
					
		totalCount = totalCount + 1
		wscript.echo totalCount
		voteCount = voteCount + 1
		'Clear cache.  Requires a copy of runDLL32.exe set to low
			'1) Copy rundll32.exe to c:\temp\
			'2) Open command line as administrator
			'3) type in the following and run : cacls c:\temp\rundll32-low.exe /setintegritylevel low
		Set WshShell = CreateObject("WScript.Shell")
		WshShell.run "c:\temp\RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 95"
		Wscript.Sleep 4000 'estimate 4 seconds to complete
	
		Set ie = createObject("InternetExplorer.Application")
		Call WaitForLoad
	
		ie.navigate "https://poll.fm/10955827"
		Call WaitForLoad
	'	ie.visible = 1
	
		'Select choice
    ie.Document.All.Item("PDI_answer###########").Click 'poll-specific.  View source on poll and find the correct ID for the choice to be selected
		
		Wscript.echo "Voted"
		Call VoteButton
						
		ie.quit
	loop
	'reset counter
	votecount = 0
	
	'Pause 1 minutes to reset throttling
	wscript.echo "Paused for Reset"
	Wscript.Sleep 60000
Loop


wscript.quit
