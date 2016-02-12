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
End Sub
