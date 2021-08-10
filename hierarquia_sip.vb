Option Explicit

	Dim hierarquia As New Selenium.ChromeDriver

	Dim col As Integer
	Dim lin As Integer
	Dim user As String
	Dim pass As String
	Dim un As String
	Dim data As String
	Dim org As String




Sub hier()


	user = Planilha9.Cells(2, 6)
	pass = Planilha9.Cells(3, 6)
	data = Planilha9.Cells(4, 6)
	org = Planilha9.Cells(5, 6)
	un = Planilha9.Cells(6, 6)



Set hierarquia = New Selenium.ChromeDriver

With hierarquia

    .Start
    .Get "https://sei.xxxxxxxxxxxx.gov.br/sip/login.php?sigla_orgao_sistema=xxxxxxxxx&sigla_sistema=SIP"
    
    
    ' Login
    .FindElementByName("txtUsuario", 2000).Click
    .SendKeys (user)
    .FindElementByName("pwdSenha", 2000).Click
    .SendKeys (pass)
    .FindElementByName("selOrgao", 2000).Click
    .SendKeys (un)
    .FindElementByClass("infraButton").Click

    
    'Acesso a hierarquia
    
    .FindElementByXPath("//*[@id='main-menu']/li[7]/a").Click
    .FindElementByXPath("//*[@id='main-menu']/li[7]/ul/li[4]/a").Click
    .FindElementByXPath("//*[@id='btnNova']").Click

    
    col = 1
    Do Until Cells(1, col) = ""
      
        lin = 2
        Do Until Cells(lin, col) = ""
        
            .FindElementById("selHierarquia").Click
            .SendKeys ("MT")
            
            
        
        
        'Cadastra primeiro a primeira coluna (que é destinado as unidades do tipo raíz) e depois cadastra as demais
        
            If col = 1 Then
            
                .FindElementByXPath("//*[@id='selOrgao']", 2000).Click
                'Application.Wait (Now + TimeValue("0:00:03"))
                .FindElementByXPath("//*[@id='selOrgao']", 2000).Click
                .SendKeys (org)
                'Application.Wait (Now + TimeValue("0:00:03"))
                .FindElementByXPath("//*[@id='selUnidade']", 2000).Click
                .FindElementByXPath("//*[@id='selUnidade']", 2000).Click
                .SendKeys Planilha8.Cells(lin, col)
                'Application.Wait (Now + TimeValue("0:00:03"))
                .FindElementByXPath("//*[@id='txtDataInicio']", 2000).Click
                .SendKeys (data)
              
                .FindElementByXPath("//*[@id='divInfraBarraComandosSuperior']/input[1]").Click
                
            Else
                
                .FindElementByXPath("//*[@id='chkRaiz']").Click
                
                
                .FindElementByXPath("//*[@id='selUnidadeSuperior']").Click
                .SendKeys Planilha8.Cells(1, col)
                .FindElementByXPath("//*[@id='selOrgao']").Click
                .FindElementByXPath("//*[text() = '" & org & "']").Click
                .FindElementByXPath("//*[@id='selUnidade']").Click
                .SendKeys Planilha8.Cells(lin, col)
                
                .FindElementByXPath("//*[@id='txtDataInicio']").Click
                .SendKeys (data)
              
                .FindElementByXPath("//*[@id='divInfraBarraComandosSuperior']/input[1]").Click
                
            End If
            
                
            .FindElementByXPath("//*[@id='btnNova']").Click
            
            lin = lin + 1
        
        Loop
  


    col = col + 1
    
    Loop



End With

End Sub
