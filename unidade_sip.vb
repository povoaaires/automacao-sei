Option Explicit

 

	Dim unidade As Selenium.ChromeDriver

'Variaveis de login
	Dim user As String
	Dim pass As String

'Variaveis de linha
	Dim Row_Start As Integer
	Dim Row_End As Integer
	Dim col_name As Integer
	Dim col_sigla As Integer
	Dim un_origem As String
	Dim un_cad As String

Sub Cad_Unidade()

	 user = Planilha2.Cells(2, 2).Value
	 pass = Planilha2.Cells(3, 2).Value
	 
	 Row_Start = Planilha2.Cells(5, 2).Value
	 Row_End = Planilha2.Cells(6, 2).Value
	 col_name = Planilha2.Cells(8, 2).Value
	 col_sigla = Planilha2.Cells(7, 2).Value
	 un_origem = Planilha2.Cells(9, 2).Value
	 un_cad = Planilha2.Cells(10, 2).Value
 
 
Set unidade = New Selenium.ChromeDriver

 

 

With unidade

 

    .Start

    .Get "https://sei.xxxxxxxxxxx.gov.br/sip/login.php?sigla_orgao_sistema=xxxxxxxxx&sigla_sistema=SIP"

     

    .FindElementByName("txtUsuario", 2000).Click

    .SendKeys (user)

    .FindElementByName("pwdSenha", 2000).Click

    .SendKeys (pass)

    .FindElementByName("selOrgao", 2000).Click

    .SendKeys (un_origem)

    .FindElementByClass("infraButton").Click

     

     

    For Row_Start = Row_Start To Row_End

     

    .FindElementByXPath("//a[contains(@title,'Unidades')]").Click

    .FindElementByXPath("//a[contains(@title,'Cadastro de Unidade')]").Click

    .FindElementByName("selOrgao", 2000).Click

    .SendKeys (un_cad)

    .FindElementByName("txtSigla", 2000).Click

    .SendKeys Planilha2.Cells(Row_Start, col_sigla).Value

    .FindElementByName("txtDescricao").Click

    .SendKeys Planilha2.Cells(Row_Start, col_name).Value

    .FindElementByName("sbmCadastrarUnidade", 2000).Click

     

     

    Next Row_Start

  

End With

End Sub
