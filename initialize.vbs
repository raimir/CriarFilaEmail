
Sub Initialize
	Dim docPrincipal As NotesDocument
	Dim docModelo As NotesDocument
	Dim realDocModelo As NotesDocument
	Dim docMetrica As NotesDocument
	Dim docCampanha As NotesDocument
	Dim staJob As Boolean
	Dim auxStJob
	Dim nthdocument As Double
	
	'variaveis para verificar se o job está ativo
	Dim stDb As NotesDatabase
	Dim stView As NotesView 
	Dim stJob As NotesDocument
	
	
	'setando o nome da function
	nameFunction = "Initialize"
	
	'Iniciando o agente
	On Error GoTo catch	
	
	Set session = New NotesSession
	Set jsonRead = New JSONReader
	
	
	'tempo inicio do agente
	Set runAgent = session.createDateTime( CStr(Now) )
	seconds = 0
	
	'usado para teste
	'Dim dbJob As NotesDatabase  
	'Set dbJob = session.Getdatabase("xtr-tinto/CONSISTE", "xtr-data\job-v1.nsf", False)
	'Set job = dbJob.Getview("JOB-cod").Getdocumentbykey("job-v1-C0D052D221C556F403257F56006FCA9F", True)
	
	'JOB documento de contexto
	Set job = session.Documentcontext
	Print "Executando " + CStr(job.job_agente(0)) + " as " + Now
	
	'inserindo campo para controle
	If (Not job.Hasitem("nthdocument")) And (Not job.Hasitem("executado")) Then  
		job.executado = 1
		job.nthdocument = 0
		Call job.save( True, False )
	End if	
	
	'variaves de status JOB
	Set stDb = job.Parentdatabase
	Set stView = stDb.Getview("JOB-cod")
	
	idInstalacao = job.idInstalacao(0)
	entidadePrincipal = job.entidadePrincipal(0)
	viewPrincipal = job.viewPrincipal(0)
	filtroPrincipal = job.filtroPrincipal(0)
	selecionarTodos = job.selecionarTodos(0)
	codModelo = job.codModelo(0)
	nomeCampanha = job.nomecampanha(0)
	
	ref_modeloemail = job.ref_modeloemail(0)
	ref_blacklistcadastro = job.ref_blacklistcadastro(0)
	ref_segmentacao = job.ref_segmentacao(0)
	ref_emailcampanha = job.ref_emailcampanha(0)
	ref_metricaemail = job.ref_metricaemail(0)
	autenticado = job.autenticado(0)
	
	
	If job.docSelecionados(0) = "" Then 
		docSelecionados = ""
	else
		docSelecionados = jsonRead.Parse( job.docSelecionados(0) ).items()
	End If
	
	jsonContasSmtp = jsonRead.Parse( job.contassmtp(0) ).items() 
	entidadesBasicas = jsonRead.Parse( job.entidadesBasicas(0) ).items()
	entidadesApp = jsonRead.Parse( job.entidadesApp(0) ).items()
	entidadesArquivos = jsonRead.Parse( job.entidadesArquivos(0) ).items()
	
	'listas databases e views
	Set listasDB = xtrListaDatabase( entidadesArquivos )
	Set listasView = xtrListaView( listasDB )
	
	'buscando documento de campanha
	Set docCampanha = listasView.getItem("emailcampanha").getDocumentBykey( job.ref_emailcampanha(0) )
	
	'criando documento de metrica
	Set docMetrica = createDocumentMetrica( listasDb )
	
	'pegando o modelo e pre renderizando as variaveis do assunto e o corpo do modelo de email que
	'não necessitam das informações das pessoas  
	
	Set realDocModelo = listasView.getItem("modeloemail").getDocumentByKey( codModelo, True )
	Set docModelo = session.currentDatabase.createDocument
	Call realDocModelo.copyAllItems( docModelo, True )
	
	'Pré renderizar
	Call xtrPreRenderiza("ModeloMailForm", docModelo, listasDb )
	Call xtrPreRenderiza("AssuntoMailForm", docModelo, listasDb )
	
	'verificando se é por seleção ou por visao
	If selecionarTodos = 0 Then
		ForAll docP In docSelecionados
			'verificando o status do job para saber se está ativo
			Set stJob = stView.Getdocumentbykey(job.xtr_cod(0), True)	
			If stJob Is Nothing Then
				Exit ForAll
			End If
			
			Set docPrincipal = listasView.getItem( StrLeft(entidadePrincipal,"-") ).getDocumentByKey( docP, True )
			Call createDocumentFila( docPrincipal, docModelo, docCampanha, docMetrica )
		End ForAll
	Else 
		Dim viewP As NotesView
		Dim viewNav As NotesViewNavigator
		Dim entryP As NotesViewEntry
		
		Set viewP = listasDB.getItem( StrLeft(entidadePrincipal,"-") ).getView( viewPrincipal )
		If Not viewP Is Nothing Then
			If viewP.isCategorized Then 
				If filtroPrincipal <> "" Then 
					Set viewNav = viewP.createViewNavFromCategory(filtroPrincipal)
				Else
					Set viewNav = viewP.createViewNav()
				End If
			Else 
				Set viewNav = viewP.createViewNav()
			End If			 
		End If
		
		'pegando a posição do último documentoque foi enviado
		nthdocument = job.nthdocument(0)
		If nthdocument > 0 Then 
			Set entryP = viewNav.Getnth(nthdocument)	
			set entryP = viewNav.getNext(entryP)
		Else
			Set entryP = viewNav.getFirst()
		End If
		
		staJob = true
		While staJob
			
			staJob = false
			
			If Not entryP Is Nothing Then
				If entryP.isDocument() Then 
					Dim posicao
					
					nthdocument = CDbl( entryP.Getposition(".") )
					Set docPrincipal = entryP.Document()
					Call createDocumentFila( docPrincipal, docModelo, docCampanha, docMetrica )
					Set entryP = viewNav.getNext( entryP )
					staJob = True
				End If
			End If
			
			'verificando o status do job para saber se o documento é existente
			Set stJob = stView.Getdocumentbykey( job.xtr_cod(0), True )	
			If stJob Is Nothing Then
				staJob = false
			else
				'parando agent depois de ser executado por um determinado tempo
				Set stopAgent = session.createDateTime( CStr(Now) )
				seconds = stopAgent.Timedifferencedouble(runAgent)
				If seconds >= 60 Then
					staJob = false
					job.nthdocument = nthdocument
					job.job_status = "AGUARDANDO"
					Call job.save(True, False)
				End If
			End If
		Wend
	End If	
	
	'print de finalização do agente
	If seconds >= 60 Then
		job.job_status = "AGUARDANDO"
	Else
		job.job_status = "CONCLUIDO"	
	End If
	'job.job_status = "CONCLUIDO"	
	Call job.Save( True, False )
	Print "Finalizando " + CStr(job.job_agente(0)) + " as " + Now 
	Exit sub
	
catch:
	If Error <> "" Then 
		Call jobPrint(job , "Agent " + job.job_agente(0) + " com Erro " + Error + " na linha" + Str(Erl) + " " + nameFunction) 
		job.job_status = "ERRO"
		Call job.Save( True, False ) 
		Print "Agente " + CStr(job.job_agente(0)) + " com Erro " + Error + " na linha" + Str(Erl) + " na função" + nameFunction
		'print de finalização do agent
		Print "Finalizando " + CStr(job.job_agente(0)) + " ás " + Now + " com erro"
	End If
	Exit sub
End Sub