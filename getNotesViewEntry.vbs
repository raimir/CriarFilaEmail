%REM
	Agent teste-v1-r1
	Created 12/02/2015 14:12 by Jonatan Raimir	
	Description: testes diversos
%END REM
Option Public
Option Declare

Sub Initialize
	Dim session As New NotesSession
	Dim db As NotesDatabase
	Dim view As NotesView
	Dim entry As NotesViewEntry
	Dim entry2 As NotesViewEntry
	Dim nav As NotesViewNavigator
	
	dim nt1 as notesdatetime, nt2 as notesdatetime
	dim seconds as double
	set nt1 = new notesdatetime(cstr(now))



	Set db = session.getDatabase("myserver","xtr-data\job-v1.nsf",false)
	Set view = db.GetView("JOB-status-agente")
	Set nav = view.Createviewnavfromcategory("CONCLUIDO")
	Set entry = nav.Getfirst()
	
	While Not entry Is Nothing 
		Dim posicao As variant 
		If entry.Isdocument()  Then
			'pegando posição do NotesViewEntry anterior
			posicao = entry.Getposition(".")
			'pegando NotesViewEntry pela posição
			Set entry2 = nav.Getpos(posicao, ".")
			'.......
		End If
		Set entry = nav.Getnext(entry)
	Wend

	set nt2 = new notesdatetime(cstr(now))
	seconds = nt2.Timedifferencedouble(nt1)

End Sub
