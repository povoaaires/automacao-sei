Option Explicit

Dim assinatura As New Selenium.ChromeDriver
Dim col As Integer
Dim lin As Integer
Dim qtd_reg As Integer
Dim valor_column As String
Dim adapt_1 As Integer
Dim tr_1 As Integer ' inicio da contagem de um tr da coluna
Dim check As String
Dim list_un As Integer
Dim item As Integer
Dim adapt_2 As Integer
Dim valor_linha As String
Dim tr_2 As Integer ' inicio da contagem de um tr da linha
Dim unidade As String
Dim nova As String
Dim checka As Integer
Dim pathcheck As String



Sub at_ass()


Set assinatura = New Selenium.ChromeDriver

With assinatura

    .Start
    .Get "https://sei.xxxxxxx.gov.br/sip/login.php?sigla_orgao_sistema=xxxxx&sigla_sistema=SEI"
    
    
    ' Login
    .FindElementByName("txtUsuario", 2000).Click
    .SendKeys Planilha10.Cells(2, 2)
    .FindElementByName("pwdSenha", 2000).Click
    .SendKeys Planilha10.Cells(3, 2)
    .FindElementByName("selOrgao", 2000).Click
    .SendKeys Planilha10.Cells(4, 2)
    .FindElementByClass("infraButton").Click

    
    'Acesso as assinaturas
    
    .FindElementByXPath("//*[@id='main-menu']/li[1]/a").Click
    .FindElementByXPath("//*[@id='main-menu']/li[1]/ul/li[1]/a").Click

    
    col = 1
    Do Until Cells(1, col) = ""
    
        .FindElementByXPath("//*[@id='txtCargoFuncao']").Clear
        .FindElementByXPath("//*[@id='txtCargoFuncao']").Click
        .SendKeys Planilha9.Cells(1, col)
        
        .FindElementByXPath("//*[@id='btnPesquisar']").Click
        
        
        qtd_reg = .FindElementByXPath("//*[@id='hdnInfraNroItens']").Value 'quantidade de registros para verificar
        valor_column = Planilha9.Cells(1, col) ' valor da coluna
        
        adapt_1 = qtd_reg + 1 'como o tr vai começar do 2 então devemos pular o 1
        
        For tr_1 = 2 To adapt_1
            
            If .FindElementByXPath("//*[@id='divInfraAreaTabela']/table/tbody/tr[" & tr_1 & "]/td[2]").Text = valor_column Then
            
                .FindElementByXPath("//*[@id='divInfraAreaTabela']/table/tbody/tr[" & tr_1 & "]/td[4]/a[1]/img").Click
                
                Exit For
            Else
            
            End If
                
        Next tr_1
        
              
      .FindElementByXPath("//*[@id='imgLupaUnidades']").Click 'Clica na lupa
      .SwitchToNextWindow (1000) ' Próxima janela
        
    lin = 2
    Do Until Cells(lin, col) = ""
        
        .FindElementByXPath("//*[@id='txtSiglaUnidade']").Clear
        .FindElementByXPath("//*[@id='txtSiglaUnidade']").Click
        .SendKeys Planilha9.Cells(lin, col)
        .FindElementByClass("infraButton").Click
        
        list_un = .FindElementByXPath("//*[@id='hdnInfraNroItens']").Value
        adapt_2 = list_un + 1
                
                 For tr_2 = 2 To adapt_2
                    
                    unidade = .FindElementByXPath("//*[@id='divInfraAreaTabela']/table/tbody/tr[" & tr_2 & "]/td[2]").Text
                    valor_linha = Planilha9.Cells(lin, col)
                    
                    
                    checka = tr_2 - 2
                    pathcheck = "//*[@id='chkInfraItem" + CStr(checka) + "']"
                    
                   
                    
                    If unidade = valor_linha Then
                            
                    
                        .FindElementByXPath(pathcheck).Click
                            
                   
                    Exit For
                    End If
                    
                    
                    
                Next tr_2
                    
        
        .FindElementByXPath("//*[@id='btnTransportarSelecao']").Click
              
        
        lin = lin + 1
        
    Loop
  
    .FindElementByXPath("//*[@id='btnFecharSelecao']").Click
        
    .SwitchToPreviousWindow ' Sai da janela
    
    .FindElementByXPath("//*[@id='divInfraBarraComandosSuperior']/button[1]").Click
    

    col = col + 1
    
    Loop



End With

End Sub