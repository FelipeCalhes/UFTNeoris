''MM40 parametros
usuario = "T3MMUNOZ "
senhaGUI = "argo.MY21"
senhaFIORI = "argo.MY20"
endpointARIBA="https://s1.ariba.com/Sourcing/Main?realm=aegea-t&passwordadapter=ThirdPartyUser"
loginARIBA="matias.munoz@neoris.com"
senhaARIBA = "Neoris.12345"
tipoDeRequisicao = "Req Material"
c = "P"
i = ""
material = "2000127"
textoBreve = ""
qtdSolicitada = "200"
precoAvaliacao = "105,00"
t = ""
dtRemessa = "06.05.2024"
grpMercads = ""
centro = "SPMI"
deposito = "DP04"
orgC = "AEGE"
gcM = "106"
acc = "1000"
perfilbd = "000000000001"
datain = "01.01.2022"
dataout = "31.12.2030"
PEP = "U/PRC.23035.E.RC.A.00601"


' 1 crear req de compras
Call AbrirAppURL("Chrome", "http://sapvdfiodtc01.latam.corp.net:8012/sap/bc/ui5_ui5/ui2/ushell/shells/abap/Fiorilaunchpad.html")
Call LoginFiori(usuario,senhaFIORI,isSenhaCriptograda)
Call AcessarAplicativoFiori("Criar requisição de compra - Ampliado") ' REVISAR, SACARLE COSAS PARA QUE FUNCIONE MEJOR
Call CriarReqDeCompraAmpliadoFiori1(tipoDeRequisicao,c,i,material,textoBreve,qtdSolicitada,precoAvaliacao,t,dtRemessa,grpMercads,centro,deposito,orgC,gcM)	
Nreqcompra = Finalizarreqmaterial(PEP)
Msgbox Nreqcompra
Call closeAllTabs()
SystemUtil.CloseProcessByName "chrome.exe"
'numeroRequisicao = "16315357"
'
Call LogonSapGui(usuario,senhaGUI)
Call AcessarTransacaoSapGui("ME54N")
Call OutrareqcompraGUI(Nreqcompra)
Call sinteseworkflowGUI()
Call Sinteseworkflowaprovacao()

'3 3.	Verificar compromisso orçamentário

'Pegar a requisição gerada na etapa 1 = 0016315357
'reqcompra = "0016315447"
Call  IngresarTX("CJI5")

'N° doc.de referência
Call CJI5Verificarcompromissoorcamentario(acc, perfilbd, PEP, datain, dataout, reqcompra)
'preguntar 2 pasos antes si esas telas nao aparecen
Call CloseSap()
'Call FecharSapGui()

'4 4.	Criar solicitação de cotação / Agregar
'nreq = "16315447"

Call AbrirAppURL("Chrome", "http://sapvdfiodtc01.latam.corp.net:8012/sap/bc/ui5_ui5/ui2/ushell/shells/abap/Fiorilaunchpad.html")
Call LoginFiori(usuario,senhaFIORI,isSenhaCriptograda)
Call AcessarAplicativoFiori("AGREGADOR") ' REVISAR, SACARLE COSAS PARA QUE FUNCIONE MEJOR
dFecha = date
fechasap = fechaFormatoSAP(dFecha)
numeroCotacao = AgregarRequisicaoDeCompraFiori(Nreqcompra, fechasap)
msgbox numeroCotacao








 @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf7.xml_;_


















