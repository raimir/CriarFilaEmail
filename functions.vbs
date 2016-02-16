%REM
	Function lockDocument
	Description: Função para travar documento evitando que outros peguem para fazer alguma tarefa
%END REM

Function lockDocument()
	If (Not job Is Nothing) Then 
		If ( Not job.hasItem("lock") ) Then 
			Call job.replaceItemValue("lock", 1)
		ElseIf job.Getfirstitem("lock").Type = 768 Then 
			If job.getItemValue("lock")(0) = 0 Then
				Call job.replaceItemValue("lock", 1)
			End If	
		Else 
			Call job.replaceItemValue("lock", 1)
		End If
	End If 
End Function


%REM
	Function unlockDocument
	Description: Função para destravar documento
%END REM

Function unlockDocument()
	If (Not job Is Nothing) Then 
		Call job.replaceItemValue("lock", 0)
	End If 
End Function

%REM
	Function isLockedDocument
	Description: Função que verifica se o documento está bloqueado para acesso
%END REM

Function isLockedDocument() as boolean
	dim locked
	locked = false

	If (Not job Is Nothing) Then 
		If ( Not job.hasItem("lock") ) Then 
			If job.Getfirstitem("lock").Type = 768 Then 
				If job.getItemValue("lock")(0) = 1 Then
					locked = true	
				End If	
			End If		
		End If
	End If 

	isLockedDocument = locked
End Function