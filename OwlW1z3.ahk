;^ = Ctrl
;+ = Shift
;! = Alt
;# = Windows Key

SetTitleMatchMode RegEx ;
#SingleInstance, Force

;	Main Script

^+!F1::										;Run Default.ahk
Run, C:\Users\rowdy\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Default.ahk
ExitApp

^+!F2::										;Run Owlwise.ahk
Run, C:\Users\rowdy\Desktop\OwlW1z3.ahk
ExitApp

!a::
send, Aqua
return

!l::
send, Black
return

!b::
send, Blue
return

!n::
send, Brown
return

!g::
send, Green
return

!o::
send, Orange
return

!p::
send, Pink
return

!r::
send, Red
return

!s::
send, Slate
return

!w::
send, White
return

!y::
send, Yellow
return

^b:: 				;Used if doing a 24 fiber run, exicute before you type the individual fiber color
send Blue Set{Space}
return

^f:: 				;types Fiber
send Fiber
return

					; Win+e is used after Win+h 
#e::				; Change the color name after generating report. 
send t
sleep 100
send {Up}{ctrl down}{Shift down}{Left}{Left}{ctrl up}{Shift up}
return

!j::					;Used to open the next .otd file, for the first one in a folder you have do this yourself
send {Enter}
WinActivate OwlView for
MouseMove 382, 75
click 382, 75			;Next .otd file
send {Enter}
MouseMove 99, 45	
click 99, 45			;click on tools
MouseMove 110, 124
click 110, 124			;click on meters
MouseMove 205, 315
click 230, 325			;Name field
Sleep 10
send, {Tab}
;send, {Tab}
send ^a{Delete}			;These three lines are  commented out for single mode
send, 500{Tab}			;enter the DeadZone
sleep 10 
send ^a{Delete}
send, 650{Tab}{Tab}		;Total Link Distance, needs to over the actually fiber legnth but less than double
send ^v					;Past in the name of the fiber run, e.i. 'FOTC-IT to PLC-CHEM ' 
						;	make sure to have the space at the end 
return

						;After the Win+g script above it stops in the 'Cable ID' field 
						;Type the fiber color followed by a space then Win+h (script below)
;
^g::
send Green Set
return

!h::					;This script fill out the bottom of the report and prints it to PDF
send {Space}Fiber{Tab}^a{Delete}State Utility Contractors{Tab}^a{Delete}WS Muddy Creek Archie{Tab}^a{Delete}919-710-3299{Tab}^a{Delete}CITI, LLC{Tab}^a{Delete}Rowdy@citi-llc.com
						;Change the company and contact info above 
MouseMove 93, 34	
click 93, 34		;click on tools
MouseMove 93, 98
click 93, 98		;click on feet
click, 71, 74		;click on rectangle zoom	
MouseClickDrag, L, 350, 420, 810, 80		;zoom to relevant space	
Click 30, 35		;click on file
sleep 1
MouseMove 16, 543	
Click 16, 533		;Print to PDF
sleep 50
send, {Enter}
send, {Enter}
sleep 600
send t
sleep 300
send {Up}{ctrl down}{Shift down}{Left}{Left}{ctrl up}{Shift up}
return

^n::
WinActivate OwlView for
MouseMove 375, 65
click 375, 65		;Next .otd file
send {Enter}
MouseMove 93, 34	
click 93, 34		;click on tools
MouseMove 93, 122
click 93, 122		;click on meters
MouseMove 205, 315
click 205, 315		;Name field
return

#o::				; This is used to save and rerun the AutoHotKey script after editing it
	send ^s
	sleep 10
	Run, C:\Users\rowdy\Desktop\OwlW1z3.ahk
return

^o::				;Used if doing a 24 fiber run, exicute before you type the individual fiber color
send Orange Set{Space}
return

#r::				; Used to fiber set color name between Orange and Blue
send {ctrl down}{Left}{Left}{Left}{Left}{Shift down}{Left}{ctrl up}{Shift up}
return

#t::				; Used to copy and rename report templates into fiber run folders
Loop, 16			; Number of folders 
{					; Copy template file and have the first folder selected before running
					; Make a copy of the folder before you run this JNC
	send {Enter}
	sleep 100
	send ^v
	sleep 100
	send !{Left}
	sleep 100
	send {Down}
	sleep 100
}
Loop, 16			; Go back to the top
{
	send {Up}
}
Loop, 16			; Rename the template to the fiber run
{
	send {F2}
	send ^c
	send {Enter}
	send {Enter}
	sleep 100
	send o{F2}
	send ^v
	sleep 10
	send {BACKSPACE}{BACKSPACE}{Space}ODTR Fiber Report{Enter}
	send !{Left}
	sleep 100
	send {Down}
	sleep 100
}
return

#y::				;Fill out ODTR single sheet reports **this is unfinished**
send Williams WTP{Tab}{Tab}23 March 2020{Tab}{Tab}
send FLT1-1{Tab}{Tab}FLT1-3{Tab}{Tab}{Tab}{Tab}
return

!q::
loop, 22
{
	send {Down}
}
return

#u::				;make pdf word docs
count := 0
loop, 12
{
WinActivate PLC-FLT1 to FLT1-2,,Fiber, Microsoft Word Document
sleep 300 
Switch count
	{
	Case 0:			;Aqua
		send {Enter}
	Case 1:			;Rose
		send {Down}{Down}{Down}{Down}{Down}{Down}{Down}
		send {Enter}
	Case 2:			;Violet
		send {Down}
		send {Enter}
	Case 3:			;Yellow
		send {Down}{Down}{Down}
		send {Enter}
	Case 4:			;Black
		send {Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}
		send {Enter}
	Case 5:			;Red
		send {Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}
		send {Enter}
	Case 6:			;White
		send {Down}
		send {Enter}
	Case 7:			;Slate
		send {Up}{Up}{Up}{Up}{Up}
		send {Enter}
	Case 8:			;Brown
		send {Up}{Up}
		send {Enter}
	Case 9:			;Green
		send {Down}
		send {Enter}
	Case 10:		;Orange
		send {Down}{Down}
		send {Enter}
	Case 11:		;Blue
		send {Up}{Up}{Up}{Up}
		send {Enter}
	Default:
		msgbox, Error
	}
count++
sleep 2000
send {Enter}
sleep 2500
click 1440, 640
sleep 200
send ^c
sleep 200
WinActivate Fiber Report, Microsoft Word Document
sleep 200
send ^v
sleep 200
;send {Backspace}{BACKSPACE}
;sleep 300
}
return


#i::
count := 0
loop, 4
{
	Switch count
	{
	Case 0:
		send {Up}
		sleep 10
	Case 1:
		send {Down}
		sleep 10
	Case 2:
		send {Left}
		sleep 10
	Case 3:
		send {Right}
		sleep 10
	Default:
		msgbox, Error
	}
count++
}
return


#j::				;change to feet and zoom
MouseMove 93, 34	
click 93, 34		;click on tools
MouseMove 93, 98
click 93, 98		;click on feet
click, 71, 74		;click on rectangle zoom	
MouseClickDrag, L, 375, 395, 420, 100		;zoom to relevant space	
return

#1::				;Fill in the top of Light Loss Report
click 826, 81
sleep 3000
click 1065, 520
send Williams WTP
send {Tab}{Tab}
send 19{Space}March{Space}2020
send {Tab}{Tab}
return 

#2::				;Fill in the wavelegnth of light loss
send {Tab}{Tab}
send 850{Space}nm
send {Tab}{Tab}
return 

#3::				;Fill in the bottum and go to next light loss report
send {Tab}{Tab}{Tab}{Tab}Yes{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}
send ^s
Sleep 6000
MouseMove 1903, 20
click 1903, 20
Sleep 1000
send !{Tab}
Sleep 1000
send !{Left}
Sleep 1000
send {Down}
sleep 10
send {Enter}
Sleep 1000
send {Down}
sleep 10
send {Enter}
return

^+t::
MouseMove 74, 101
return

#z::Pause			;will pause a script while it is running, good if you accidently run the wrong one
return				;	and it starts to go crazy


					;END