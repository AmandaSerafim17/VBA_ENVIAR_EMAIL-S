Sub Enviar_email()

'Enviar o mesmo anexo para varios e-mails

        'Sheet de trabalho - Ajustar conforme necessidade
        Planilha1.Select
        
        'variavel
        Dim email As String
        Dim emailcc As String
        Dim regional As String
        Dim assunto As String
        Dim corpo As String
        Dim anexo As String
        
		'Estou definindo inicio e fim das linhas de email (P = Para)
		Dim PlinhaInicial as Integer
		Dim PlinhaFinal as Integer
		Dim linha as Integer
		
        
        'Atribuindo valor às variaveis
        assunto = Range("E2").Value
        corpo = Range("F2").Value
        regional = Range("A2").Value
        anexo = Range("G2").Value
		
		'Ajustas as linhas conforme necessidade
		PlinhaInicial = 2
		PlinhaFinal = 6
        
        'Enviar email
		For linha = PlinhaInicial to PlinhaFinal  'Variaveis e Objetos estão abaixo por conta do laço de repetição
			email = Range("B" & linha).Value
			emailcc = Range("C" & linha).Value
			anexo = Range("G" & linha).Value
			regional = Range("A" & linha).Value
			
			Set objOutlook = CreateObject("Outlook.Application")
			Set novoEmail = objOutlook.CreateItem(0)
		
			With novoEmail
				.to = email
				.cc = emailcc
				.Subject = assunto & regional
				.body = corpo
				.attachments.Add anexo
				.display
			
			End With
			Application.Wait (Now + TimeValue("00:00:02")) 'Usando para o outlook ter tempo de processamento
			Set novoEmail = Nothing 'zerar o e-mail para um novo envio
			
		Next linha


End Sub
