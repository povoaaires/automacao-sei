Option Explicit

 

	Dim usuario As Selenium.ChromeDriver, intervalo As Integer

'Variaveis de login
	Dim user As String
	Dim pass As String

'Variaveis de linha
	Dim Row_Start As Integer
	Dim Row_End As Integer
	Dim col_name As Integer
	Dim col_sigla As Integer
	Dim un As String
	Dim un_origem As String


Sub Cad_Usuario()

	 user = Planilha4.Cells(2, 2).Value
	 pass = Planilha4.Cells(3, 2).Value
	 
	 Row_Start = Planilha4.Cells(5, 2).Value
	 Row_End = Planilha4.Cells(6, 2).Value
	 col_name = Planilha4.Cells(7, 2).Value
	 col_sigla = Planilha4.Cells(8, 2).Value
	 un = Planilha4.Cells(9, 2).Value
	 un_origem = Planilha4.Cells(10, 2).Value
 
Set usuario = New Selenium.ChromeDriver


 

With usuario

 

    .Start

    .Get "https://sei.xxxxxxxxx.gov.br/sip/login.php?sigla_orgao_sistema=xxxxxxxx&sigla_sistema=SIP"

     

    .FindElementByName("txtUsuario", 2000).Click

    .SendKeys (user)

    .FindElementByName("pwdSenha", 2000).Click

    .SendKeys (pass)

    .FindElementByName("selOrgao", 2000).Click

    .SendKeys (un_origem)

    .FindElementByClass("infraButton").Click

     

     

    For Row_Start = Row_Start To Row_End

     

    .FindElementByXPath("//a[contains(@title,'Usuários')]").Click

    .FindElementByXPath("//a[contains(@title,'Cadastro de Usuário')]").Click

    .FindElementByName("selOrgao", 2000).Click

    .SendKeys (un)

    .FindElementByName("txtSigla", 2000).Click

    .SendKeys Planilha4.Cells(Row_Start, col_sigla).Value

    .FindElementByName("txtNome").Click

    .SendKeys Planilha4.Cells(Row_Start, col_name).Value

    .FindElementByName("sbmCadastrarUsuario", 2000).Click

     

     

    Next Row_Start

  

End With

End Sub
