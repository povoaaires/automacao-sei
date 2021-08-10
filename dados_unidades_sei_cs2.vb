'Este código é utilizado apenas para cadastro dos dados das unidades QUE POSSUEM O ENDEREÇO SEMELHANTE AO DO ÓRGÃO 

Option Explicit

	Dim SEI As Selenium.ChromeDriver

'Variáveis para login
	Dim user As String
	Dim pass As String

 
'Variáveis de linhas e colunas
	Dim linha_start As Integer
	Dim linha_end As Integer
	Dim un As Integer
	Dim mail_un As Integer
	Dim desc_un As Integer
	Dim url_un As Integer
	Dim cnpj_un As Integer
	Dim tel_un As Integer
	Dim nup As Long
	Dim element As String
	Dim comp As String
	Dim orgao As String



'Variáveis condicionais
	Dim verify As Boolean
	Dim check As Boolean

'Variável do comando 'IsElementPresent'
	Dim By As New By

Sub automacao_dados_unidades()


	user = Planilha2.Cells(2, 2).Value
	pass = Planilha2.Cells(3, 2).Value

	linha_start = Planilha2.Cells(5, 2).Value
	linha_end = Planilha2.Cells(6, 2).Value
	un = Planilha2.Cells(12, 2).Value
	mail_un = Planilha2.Cells(9, 2).Value
	desc_un = Planilha2.Cells(10, 2).Value
	url_un = Planilha2.Cells(11, 2).Value
	cnpj_un = Planilha2.Cells(7, 2).Value
	tel_un = Planilha2.Cells(8, 2).Value
	nup = Planilha2.Cells(13, 2).Value
	orgao = Planilha2.Cells(14, 2).Value

 

Set SEI = New Selenium.ChromeDriver
With SEI

 

'Iniciando o browser e acessando a página web

 

	.Start
	.Get "https://sei.xxxxxxxxxx.gov.br/sip/login.php?sigla_orgao_sistema=xxxxxxx&sigla_sistema=SEI"

 


'Fazendo o login na página web

 

	.FindElementByName("txtUsuario", 2000).Click    'Usuário
	.SendKeys (user)   'sigla do usuário
	.FindElementByName("pwdSenha", 2000).Click      'Senha
	.SendKeys (pass)  ' senha do usuário
	.FindElementByName("selOrgao", 2000).Click   'Órgão
	.SendKeys ("MINFRA")    'Órgão do usuário
	.FindElementByClass("infraButton").Click    'Entrar

 


'Caminho para a página destino

 

	.FindElementByXPath("//a[contains(@class,'has-submenu')]").Click   'Opção Administração
	.FindElementByXPath("//*[@id='main-menu']/li[1]/ul/li[3]/a/span").Click      'Opção Unidades (Dentro da Opção Administração)
	.FindElementByXPath("//*[@id='main-menu']/li[1]/ul/li[3]/ul/li[1]/a").Click     'Opção Listar (Dentro da Opção Unidades)

 

'Selecionando a caixa de opção unidades na página Listar
	.FindElementByName("selOrgao", 2000).Click      'Selecionando a caixa de opção dentro da página Listar

 

'ATENÇÃO, este comando está selecionando a terceira opção da caixa de seleção
	.FindElementByName("selOrgao").Click
	.SendKeys (orgao)

 


'Laço com base na quantidade de linhas da planilha

 

For linha_start = linha_start To linha_end

 

'Pesquisando a primeira unidade

 

    .FindElementByName("txtSiglaUnidade", 2000).Clear  'Limpa o campo de sigla de unidade
    .FindElementByName("txtSiglaUnidade", 2000).Click   'Clica no campo de sigla de unidade
    .SendKeys Planilha2.Cells(linha_start, un)                 'Passa o valor da linha x do excel para o campo em questão
    .FindElementByXPath("//*[@id='btnPesquisar']").Click    'Seleciona o botão pesquisar
    
    
    element = .FindElementByXPath("//*[@id='divInfraAreaTabela']/table/tbody/tr[2]/td[3]").Text
    
    comp = Planilha2.Cells(linha_start, un)
    
    
'Selecionando o botão de alterar da unidade pesquisada

        If .IsElementPresent(By.XPath("//*[@id='divInfraAreaTabela']/table/tbody/tr[2]/td[6]/a[2]/img")) Then
        
            If element = comp Then
    
                .FindElementByXPath("//*[@id='divInfraAreaTabela']/table/tbody/tr[2]/td[6]/a[2]/img", 2000).Click
            Else
                .FindElementByXPath("//*[@id='divInfraAreaTabela']/table/tbody/tr[3]/td[6]/a[2]/img", 2000).Click
                
            End If
                
        Else
            MsgBox "não disponível"
            
            'procurar um comando que pare todo o código
        End If
   
   
        

'Dentro do perfil de alteração, o comando abaixo seleciona e passa um parâmetro para o campo código

 

    .FindElementByName("txtCodigoSei", 2000).ClickDouble 'Selecionando o campo código SEI
    .SendKeys (nup)         'Passando o valor para o código
    
    
'Verifica se há dados inseridos no campo E-mail e Descrição

 

    verify = .IsElementPresent(By.Class("infraTrClara"))
    
    If verify = False Then
    
            'Preenche os campos E-mail e Descrição caso tais campos não estejam preenchidos
            
            .FindElementByName("txtEmail", 2000).Click  'Campo email da unidade
            .SendKeys Planilha2.Cells(linha_start, mail_un)     '   Passando o valor contido na linha 4 coluna 3 (o valor da linha será alterado)
            .FindElementByName("txtDescricaoEmail", 2000).Click     'Campo descrição da unidade
            .SendKeys Planilha2.Cells(linha_start, desc_un)     '   Passando o valor contido na linha 4 coluna 4 (o valor da linha será alterado)
            .FindElementByName("sbmGravarEmail", 2000).Click    ' Salvando os valores passados

 

'Fim da condificional
    End If
    
'Seleciona a figura que remete à alteração dos dados do contato associado
    .FindElementByXPath("//*[@id='imgAlterarContato']", 2000).Click

 

'Permite que o Selenium acesse o pop up para fazer as alterações
    .SwitchToNextWindow (1000)
    
'Verifica se o check box do pop up está selecionado ou não

 

    check = .FindElement(By.XPath("/html/body/div[1]/div/div/form/div[3]/div/input")).IsSelected()
    
    
    If check = False Then
    
            'Comando que seleciona o check box para puxar os dados da unidade cadastrada
            .FindElementByName("chkSinEnderecoAssociado", 2000).Click
  
'Fim da condicional
    End If
    
   
'Limpa os dados que foram preenchidos e os preenche novamente
    
    .FindElementByName("txtSitioInternet", 2000).Clear
    .FindElementByName("txtSitioInternet", 2000).Click
    .SendKeys Planilha2.Cells(linha_start, url_un)
    .FindElementByName("txtCnpj", 2000).Clear
    .FindElementByName("txtCnpj", 2000).Click
    .SendKeys Planilha2.Cells(linha_start, cnpj_un)
    .FindElementByName("txtTelefoneFixoPJ", 2000).Clear
    .FindElementByName("txtTelefoneFixoPJ", 2000).Click
    .SendKeys Planilha2.Cells(linha_start, tel_un)
'Salva as alterações
    .FindElementByName("sbmAlterarContato").Click
    
'Sai do pop up
    .SwitchToPreviousWindow
    
'Salva as alterações da Unidade
    .FindElementByName("sbmAlterarUnidade", 2000).Click
              
   
'Contador
 Next linha_start
 

 

End With
End Sub
