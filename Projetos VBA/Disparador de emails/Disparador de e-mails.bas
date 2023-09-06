Attribute VB_Name = "M�dulo1"
Sub CriarEmail()

    Dim hoje As String
    
    Worksheets("Envio de e-mails").Activate
    If Range("A1048570").Value = "Fila de Projetos" Then
        'MsgBox "O processo j� est� conclu�do"
    ElseIf Range("A1048570").Value = "Declinada" Then
        MsgBox "O processo j� est� conclu�do"
    'ElseIf Range("A1048570").Value = "Conclu�da" Then
     '   MsgBox "O processo j� est� conclu�do"
    ElseIf Range("A1048575").Value = "" Then
        MsgBox "N�o foi poss�vel executar o processo, por falta do email do respons�vel. Informe um respons�vel pelo processo."
    ElseIf Range("A1048570").Value = "No Prazo" Then
        MsgBox "O processo ainda est� dentro do prazo"
    ElseIf Range("A1048570").Value = "Aguardando Prazo" Then
        MsgBox "O processo n�o possui prazo definido."
    Else
        EnviarEmail
    End If
    
    
    
    
End Sub

Sub EnviarEmail()

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim Para As String
    Dim Copia As String
    Dim Assunto As String
    Dim Corpo As String
    Dim Prazo As String
    
    
 ' Cria uma inst�ncia do aplicativo Outlook
    Worksheets("Envio de e-mails").Activate
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    Prazo = Range("A1048573").Value
    PROCESSO = Range("A1048572").Value
    Assunto = Range("C10").Value
    Para = Range("A1048575").Value
    Copia = Range("A1048576").Value
    Corpo = Range("C16").Value
    
    OutlookMail.display
    ' Preencha as informa��es do e-mail
    With OutlookMail
        .To = Para
        .CC = Copia
        .Subject = Assunto
        .Body = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Range("C15"), "[Respons�vel]", Range("A1048563")), "[C�lula]", Range("A1048561")), "[C�lula]", Range("A1048561")), "[Data da Solicita��o]", Range("A1048557")), "[Tarefa/A��o]", Range("A1048558")), "[Setor Respons�vel]", Range("A1048559")), "[Origem]", Range("A1048560")), "[Solicitante]", Range("A1048562")), "[Respons�vel]", Range("A1048563")), "[ID]", Range("A1048564")), "[�ltimo Prazo]", Range("A1048565")), "[Aging]", Range("A1048566")), "[Problema / Oportunidade]", Range("A1048567"))
        '.SEND
       
    End With
    
        
        
        ' Libera os objetos
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub

Sub AtualizaCorpo()
    
    Dim Status As String
    Dim Conclusao As String
    Dim Nova As String
    Dim Atraso As String

    Worksheets("Envio de e-mails").Activate
    Conclusao = Range("A1048569").Value
    Nova = Range("B1048569").Value
    Atraso = Range("C1048569").Value
    
    
    
    If Range("A1048570").Value = "Nova" Then
        Range("C15").Value = Nova
        Range("C10").Value = Range("A1048574").Value
    ElseIf Range("A1048570").Value = "Conclu�da" Then
        Range("C15").Value = Conclusao
        Range("C10").Value = Range("A1048574").Value
    ElseIf Range("A1048570").Value = "Atrasada" Then
        Range("C15").Value = Atraso
        Range("C10").Value = Range("A1048574").Value
    Else
        Range("C15").Value = ""
        Range("C10").Value = ""
    End If

End Sub

Sub ProcessosAtrasados()

    Linha = 2
    Dim Valor As String
    Cont = 0

    Do Until Linha = 5000
    Worksheets("A��es").Activate
    If Cells(Linha, 16) = "Atrasada" And Cells(Linha, 6) <> "" Then
        If Rows(Linha).Hidden = False Then
         Valor = Cells(Linha, 5)
         Worksheets("Envio de e-mails").Activate
         Range("C7") = Valor
         Worksheets("A��es").Activate
        Call EnvioMassa
        Cont = Cont + 1
        End If
    End If
    
    Linha = Linha + 1
    
Loop
    MsgBox ("Foram criados " & Cont & " e-mails")

End Sub

Sub EnvioMassa()

Dim hoje As String
    
    Worksheets("Envio de e-mails").Activate
    If Range("A1048570").Value = "Fila de Projetos" Then
        'MsgBox "O processo j� est� conclu�do"
    ElseIf Range("A1048570").Value = "Declinada" Then
        MsgBox "O processo j� est� conclu�do"
    'ElseIf Range("A1048570").Value = "Conclu�da" Then
     '   MsgBox "O processo j� est� conclu�do"
    ElseIf Range("A1048575").Value = "" Then
        MsgBox "N�o foi poss�vel executar o processo, por falta do email do respons�vel. Informe um respons�vel pelo processo."
    ElseIf Range("A1048570").Value = "No Prazo" Then
        MsgBox "O processo ainda est� dentro do prazo"
    ElseIf Range("A1048570").Value = "Aguardando Prazo" Then
        MsgBox "O processo n�o possui prazo definido."
    Else
        EnviarEmailMassa
    End If
End Sub

Sub EnviarEmailMassa()

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim Para As String
    Dim Copia As String
    Dim Assunto As String
    Dim Corpo As String
    Dim Prazo As String
    
    
 ' Cria uma inst�ncia do aplicativo Outlook
    Worksheets("Envio de e-mails").Activate
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    Prazo = Range("A1048573").Value
    PROCESSO = Range("A1048572").Value
    Assunto = Range("C10").Value
    Para = Range("A1048575").Value
    Copia = Range("A1048576").Value
    Corpo = Range("C16").Value
    
    OutlookMail.display
    ' Preencha as informa��es do e-mail
    With OutlookMail
        .To = Para
        .CC = Copia
        .Subject = Assunto
        .Body = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Range("C15"), "[Respons�vel]", Range("A1048563")), "[C�lula]", Range("A1048561")), "[C�lula]", Range("A1048561")), "[Data da Solicita��o]", Range("A1048557")), "[Tarefa/A��o]", Range("A1048558")), "[Setor Respons�vel]", Range("A1048559")), "[Origem]", Range("A1048560")), "[Solicitante]", Range("A1048562")), "[Respons�vel]", Range("A1048563")), "[ID]", Range("A1048564")), "[�ltimo Prazo]", Range("A1048565")), "[Aging]", Range("A1048566")), "[Problema / Oportunidade]", Range("A1048567"))
        '.SEND
       
    End With
    
        
        
        ' Libera os objetos
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub
