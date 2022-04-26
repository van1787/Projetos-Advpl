//====================================================================================================
// Definicoes de Includes da Rotina.
//====================================================================================================

#Include	"Protheus.Ch"
#Include	"FWMVCDef.Ch"
#Include	"topconn.Ch"
/*
===============================================================================================================================
Programa----------: MFIS001
Autor-------------: Vanderson Azevedo Vasconcelos
Data da Criacao---: 16/04/2018
===============================================================================================================================
Descrição---------: Rotina para gerar notas fiscais SF2
===============================================================================================================================
Retorno-----------: Nenhum
===============================================================================================================================
*/

User Function MFIS_01()//U_MFIS_01()

PRIVATE MV_PAR01:=Space(200)
PRIVATE _cSalva_MV_PAR01 := MV_PAR01

DO WHILE .T.

   lSair:=MFIS001INI()

   IF lSair
      EXIT
   ENDIF
ENDDO

	
Return()

//===============================================================================================================================
Static Function MFIS001INI()
//===============================================================================================================================
Local _aParAux	:= {}
Local _aParRet	:= {}
Local _aCpos	:= MFIS001CPS()
Local _aFields	:= {}
Local _cArqTrab	:= ""
Local _nreg		:= 0
Local _adados	:= {}
//Local _apesqui	:= {}
Local nI		:= 0  , I
Local _cArq		:= ""
Local _lOK      :=.F.

Private oMarkBRW	:= Nil
Private cAliasAux	:= GetNextAlias()
Private _nTotReg	:= 0
Private _cSerie     := "U  "	 //TESTES  "1"//


MV_PAR01 := _cSalva_MV_PAR01
 
AADD( _aParAux , { 1 , "Selecione arquivo:"		, MV_PAR01, "@!"		, ""	, "DIR"		, "" , 120 , .F. } ) 	//| 01 |

For nI := 1 To Len( _aParAux )
	aAdd( _aParRet , _aParAux[nI][03] )
Next nI

IF !ParamBox( _aParAux , "Leitura de Arquivo de Notas" , @_aParRet )
	Return .T.
EndIf

_cSalva_MV_PAR01 := MV_PAR01

IF (upper(right(alltrim(MV_PAR01),4))) == ".CSV" 

	_cArq := ALLTRIM(MV_PAR01)

	If FT_FUSE(_cArq) == -1
       Aviso( "Falha ao abrir" , "Falha ao abrir o arquivo: "+_cArq , {"Fechar"} )
	   Return .F.
	Endif 
	          
  	FT_FGOTOP() //POSICIONA NO TOPO DO ARQUIVO
  	_cDados := FT_FREADLN()

  	If UPPER(alltrim(_cDados)) = UPPER("Tipo de Registro;N")///********************
	   _lOK := .T. //Layout Padrão
  	Endif
	_nReg:= FT_FLASTREC()

	//Fecha arquivo e prepara parâmetro com arquivo convertido
  	FT_FUSE()

ELSE
	Aviso(  "Arquivo inválido","O arquivo informado: "+Alltrim(MV_PAR01)+" não tem extenção [ .CSV ] "+;            
		    "Favor informar uma arquivo no formato [ .CSV ].", {"Fechar"} )
	Return .F.

EndIf

If !_lOK

	U_ITMSG("Arquivo com formato inválido" ,;
	        "O arquivo "+_cArq+" informado para relizar a importação não tem o layout correto. "+;
			'O CSV deve ter as colunas iniciais "Tipo de Registro;Nº NFS-e;Data Hora NFE;..."', {"Fechar"} )

	Return .F.
	
Endif

DBSELECTAREA("SE1")
ProcRegua(len(_adados))

_cArqTrab:= E_CriaTrab(,_aCpos,cAliasAux)
//Cria índice
IndRegua(cAliasAux, _cArqTrab, "WKSTATUS+WK_CPO02", , , "Selecionando Registros...")

_lTemNFs:=.F.
FWMSGRUN(,{|oproc| _lTemNFs := MFIS001L(_nreg,oproc,_cArq) } , "Aguarde...", "Lendo Arquivo: "+_cArq)

If !_lTemNFs

	Aviso( "Atenção!" , "Não foram encontrados registros válidos no arquivo!" , {"Fechar"} )
	Return .F.
	
EndIf

(cAliasAux)->( DBGOTOP() )
_cTexto:="Tipo de Registro;Nº NFS-e;Data Hora NFE;Código de Verificação da NFS-e;Tipo de RPS;Série do RPS;Número do RPS;Data do Fato Gerador;Inscrição Municipal do Prestador;Indicador de CPF/CNPJ do Prestador;CPF/CNPJ do Prestador;Razão Social do Prestador;Tipo do Endereço do Prestador;Endereço do Prestador;Número do Endereço do Prestador;Complemento do Endereço do Prestador;Bairro do Prestador;Cidade do Prestador;UF do Prestador;CEP do Prestador;Email do Prestador;Opção Pelo Simples;Situação da Nota Fiscal;Data de Cancelamento;Nº da Guia;Data de Quitação da Guia Vinculada a Nota Fiscal;Valor dos Serviços;Valor das Deduções;Código do Serviço Prestado na Nota Fiscal;Alíquota;ISS devido;Valor do Crédito;ISS Retido;Indicador de CPF/CNPJ do Tomador;CPF/CNPJ do Tomador;Inscrição Municipal do Tomador;Inscrição Estadual do Tomador;Razão Social do Tomador;Tipo do Endereço do Tomador;Endereço do Tomador;Número do Endereço do Tomador;Complemento do Endereço do Tomador;Bairro do Tomador;Cidade do Tomador;UF do Tomador;CEP do Tomador;Email do Tomador;Discriminação dos Serviços"
aCpos:=ITTXTARRAY(_cTexto,";",48)
FOR I := 1 TO 47
//                      TITULO               CAMPO      TIPO,PICTURE,,TAM,DEC
    IF I = 4
       AADD( _aFields,{ "Cod Municipio"    ,"WK_COD_MUN","C" ,""     ,0,LEN(CC2->CC2_CODMUN), 0 } )//CIDADE DO CLIENTE
       AADD( _aFields,{ "Valor do Credito"  , "WK_CPO32","C" ,""     ,0,20,0} )
       AADD( _aFields,{ "Aliquota do ISS"   , "WK_CPO30","C" ,""     ,0,20,0} )
       AADD( _aFields,{ "Valor dos Servicos", "WK_CPO27","C" ,""     ,0,20,0} )
       AADD( _aFields,{ "Valor das Deducoes", "WK_CPO28","C" ,""     ,0,20,0} )
       AADD( _aFields,{ "Valor ISS devido"  , "WK_CPO31","C" ,""     ,0,20,0} )
    ENDIF    
    cConteudo:=(cAliasAux)->( FIELDGET(FIELDPOS("WK_CPO"+STRZERO(I,2))) )
    AADD( _aFields,{ aCpos[I] , "WK_CPO"+STRZERO(I,2),"C" ,""     ,0,LEN(cConteudo),0} )
NEXT

AADD( _aFields,{ "Discriminação dos Serviços" ,{|| LEFT((cAliasAux)->WK_DSCSER,50) },"C","",0,050,0} )
AADD( _aFields,{ "Observações / Rejeições"    ,"WK_MOTIVO","C","",0,100,0} )
AADD( _aFields,{ "Erro do MSExecAuto()"       ,"WK_ERRO"  ,"C","",0,100,0} )

AADD( _aFields,{ "Cod. Cliente"     ,"WK_CLICOD" ,"C",""                          ,0,LEN(SF2->F2_CLIENTE), 0 } )//CODIGO DO CLIENTE
AADD( _aFields,{ "Loja Cliente"     ,"WK_CLOLOJ" ,"C",""                          ,0,LEN(SF2->F2_LOJA   ), 0 } )//LOJA DO CLIENTE
AADD( _aFields,{ "Cod Municipio"    ,"WK_COD_MUN","C",""                          ,0,LEN(CC2->CC2_CODMUN), 0 } )//LOJA DO CLIENTE
AADD( _aFields,{ "Valor do Crédito" ,"WK_VALCRE" ,"N",PesqPict('SF2',"F2_CREDNFE"),0,TamSx3("F2_CREDNFE")[1],TamSx3("F2_CREDNFE")[2] } )//Valor do Crédito
AADD( _aFields,{ "Cod. Produto"     ,"WK_CODITEM","C",""                          ,0,LEN(SD2->D2_COD    ), 0 } )//CODIGO DO PRODUTO
AADD( _aFields,{ "TES"              ,"WK_TES"    ,"C",""                          ,0,LEN(SB1->B1_TS     ), 0 } )//CODIGO DA TES
AADD( _aFields,{ "Valor Serviços"   ,"WK_SERVICO","N",PesqPict('SD2',"D2_TOTAL"  ),0,TamSx3("D2_TOTAL"  )[1],TamSx3("D2_TOTAL"  )[2] } )//Valor dos Serviços
AADD( _aFields,{ "Valor Deduções"   ,"WK_DEDUCO" ,"N",PesqPict('SD2',"D2_DESCON" ),0,TamSx3("D2_DESCON" )[1],TamSx3("D2_DESCON" )[2] } )//Valor das Deduções
AADD( _aFields,{ "Aliq ISS"         ,"WK_ALIQISS","N",PesqPict('SD2',"D2_ALIQISS"),0,TamSx3("D2_ALIQISS")[1],TamSx3("D2_ALIQISS")[2] } )//Alíquota do ISS Recebe o valor da coluna “Alíquota” da Planilha 
AADD( _aFields,{ "Base ISS"         ,"WK_BASEISS","N",PesqPict('SD2',"D2_BASEISS"),0,TamSx3("D2_BASEISS")[1],TamSx3("D2_BASEISS")[2] } )//Base do ISS Recebe o valor da coluna “Valor dos Serviços” da Planilha ( - ) o valor da coluna “Valor das Deduções” da Planilha 
AADD( _aFields,{ "Valor ISS"        ,"WK_VALISS" ,"N",PesqPict('SD2',"D2_VALISS" ),0,TamSx3("D2_VALISS" )[1],TamSx3("D2_VALISS" )[2] } )//Valor do ISS

_apesqui := {}
AADD(_apesqui,{"Status", {{"Status","C",1,0,"Status","@!"}} } )

oMarkBRW:=FWMarkBrowse():New()		   											  		// Inicializa o Browse
oMarkBRW:SetAlias( cAliasAux )			   												// Define Alias que será a Base do Browse
oMarkBRW:SetDescription( "Importação NFSe SP" )											// Define o titulo do browse de marcacao
oMarkBRW:SetFieldMark( "MARCA" )														// Define o campo que sera utilizado para a marcação
//oMarkBRW:SetMenuDef( 'MFIS_01' )														// Força a utilização do menu da rotina atual
oMarkBRW:SetAllMark( {|| oMarkBRW:AllMark() } )											// Ação do Clique no Header da Coluna de Marcação
oMarkBRW:SetValid({|| IF((cAliasAux)->WKSTATUS $ "A",.T.,.F.) })						// Indica o Code-Block executado para validar a marcação/desmarcação do registro
//oMarkBRW:SetCustomMarkRec({|| IF((cAliasAux)->WKSTATUS $ "A,1",U_MFISMUN(),.F.)})		// Indica o Code-Block executado para validar a marcação/desmarcação do registro
oMarkBRW:AddLegend( "(cAliasAux)->WKSTATUS = '1'","YELLOW","1-Pendente")                // Permite adicionar legendas no Browse
oMarkBRW:AddLegend( "(cAliasAux)->WKSTATUS = 'A'","GREEN" ,"A-Importação Liberada")     // Permite adicionar legendas no Browse
oMarkBRW:AddLegend( "(cAliasAux)->WKSTATUS = 'P'","BLUE"  ,"P-NF Integrada com Sucesso")// Permite adicionar legendas no Browse
oMarkBRW:AddLegend( "(cAliasAux)->WKSTATUS = 'R'","RED"   ,"R-Importação com Restrição")// Permite adicionar legendas no Browse
oMarkBRW:AddLegend( "(cAliasAux)->WKSTATUS = 'S'","BLACK" ,"S-NF NÃO foi integrada")    // Permite adicionar legendas no Browse
oMarkBRW:SetFields( _aFields )													 		// Campos para exibição
oMarkBRW:SetSeek(.T., _apesqui)                                                         // Indica os campos que serão adicionados as colunas do Browse 
oMarkBRW:SetSeek(.T.)
oMarkBRW:AddButton( "Confirmar"      , {|| Processa( {|| U_MFIS001B()} ) } ,, 4 )		// Adiciona um botão na área lateral do Browse
oMarkBRW:AddButton( "Inclui Cod.Mun.", {||               U_MFISMUN() }     ,, 4 )		// Adiciona um botão na área lateral do Browse
oMarkBRW:AddButton( "Obs / Rejeicao" , {|| Aviso("Observações / Rejeições"   ,(cAliasAux)->WK_MOTIVO,{"Fechar"})} ,, 4 )// Adiciona um botão na área lateral do Browse
oMarkBRW:AddButton( "Erro MSExecAuto", {|| Aviso("Erro do MSExecAuto()"      ,(cAliasAux)->WK_ERRO  ,{"Fechar"})} ,, 4 )// Adiciona um botão na área lateral do Browse
oMarkBRW:AddButton( "Discr. Servicos", {|| Aviso("Discriminação dos Serviços",(cAliasAux)->WK_DSCSER,{"Fechar"})} ,, 4 )// Adiciona um botão na área lateral do Browse
//oMarkBRW:DisableConfig()                                                                // Desabilita a utilização das configurações do Browse 
oMarkBRW:AllMark()                                                                      // Marca Todos
oMarkBRW:Activate()																		// Ativacao da classe

(cAliasAux)->( DBCloseArea() )

Return .F.

//===============================================================================================================================
//Static Function MenuDef()
//===============================================================================================================================
//Local aRotina	:= {}
//ADD OPTION aRotina TITLE 'Alterar'    ACTION 'U_MFIS001R("ALTERAR")' OPERATION 4 ACCESS 0 
//ADD OPTION aRotina Title 'Visualizar' Action 'U_MFIS001R("VISUAL")'  OPERATION 2 ACCESS 0
//Return( aRotina )

//===============================================================================================================================
Static Function MFIS001CPS()//             ESTRUTURAS
//===============================================================================================================================
Local _aCpos := {}
AADD( _aCpos , { "WKSTATUS", "C" ,01, 0 } )
aAdd( _aCpos , { "MARCA"   , "C" ,01, 0 } )
aAdd( _aCpos , { "WK_CPO01", "C" ,06, 0 } )//Tipo de Registro;
aAdd( _aCpos , { "WK_CPO02", "C" ,10, 0 } )//Nº NFS-e;
aAdd( _aCpos , { "WK_CPO03", "C" ,20, 0 } )//Data Hora NFE;
aAdd( _aCpos , { "WK_CPO04", "C" ,10, 0 } )//Código de Verificação da NFS-e;
aAdd( _aCpos , { "WK_CPO05", "C" ,05, 0 } )//Tipo de RPS
aAdd( _aCpos , { "WK_CPO06", "C" ,05, 0 } )//Série do RPS
aAdd( _aCpos , { "WK_CPO07", "C" ,05, 0 } )//Número do RPS
aAdd( _aCpos , { "WK_CPO08", "C" ,15, 0 } )//Data do Fato Gerador
aAdd( _aCpos , { "WK_CPO09", "C" ,15, 0 } )//Inscrição Municipal do Prestador
aAdd( _aCpos , { "WK_CPO10", "C" ,03, 0 } )//Indicador de CPF/CNPJ do Prestador
aAdd( _aCpos , { "WK_CPO11", "C" ,15, 0 } )//CPF/CNPJ do Prestador
aAdd( _aCpos , { "WK_CPO12", "C" ,40, 0 } )//Razão Social do Prestador
aAdd( _aCpos , { "WK_CPO13", "C" ,05, 0 } )//Tipo do Endereço do Prestador
aAdd( _aCpos , { "WK_CPO14", "C" ,99, 0 } )//Endereço do Prestador
aAdd( _aCpos , { "WK_CPO15", "C" ,15, 0 } )//Número do Endereço do Prestador
aAdd( _aCpos , { "WK_CPO16", "C" ,25, 0 } )//Complemento do Endereço do Prestador
aAdd( _aCpos , { "WK_CPO17", "C" ,50, 0 } )//Bairro do Prestador
aAdd( _aCpos , { "WK_CPO18", "C" ,50, 0 } )//Cidade do Prestador
aAdd( _aCpos , { "WK_CPO19", "C" ,03, 0 } )//UF do Prestador
aAdd( _aCpos , { "WK_CPO20", "C" ,10, 0 } )//CEP do Prestador
aAdd( _aCpos , { "WK_CPO21", "C" ,50, 0 } )//Email do Prestador
aAdd( _aCpos , { "WK_CPO22", "C" ,02, 0 } )//Opção Pelo Simples
aAdd( _aCpos , { "WK_CPO23", "C" ,02, 0 } )//Situação da Nota Fiscal
aAdd( _aCpos , { "WK_CPO24", "C" ,10, 0 } )//Data de Cancelamento
aAdd( _aCpos , { "WK_CPO25", "C" ,10, 0 } )//Nº6da Guia
aAdd( _aCpos , { "WK_CPO26", "C" ,10, 0 } )//Data de Quitação da Guia Vinculada a Nota Fiscal
aAdd( _aCpos , { "WK_CPO27", "C" ,20, 0 } )//Valor dos Serviços
aAdd( _aCpos , { "WK_CPO28", "C" ,20, 0 } )//Valor das Deduções
aAdd( _aCpos , { "WK_CPO29", "C" ,06, 0 } )//Código do Serviço Prestado na Nota Fiscal
aAdd( _aCpos , { "WK_CPO30", "C" ,10, 0 } )//Alíquota
aAdd( _aCpos , { "WK_CPO31", "C" ,15, 0 } )//ISS devido
aAdd( _aCpos , { "WK_CPO32", "C" ,15, 0 } )//Valor do Crédito
aAdd( _aCpos , { "WK_CPO33", "C" ,15, 0 } )//ISS Retido
aAdd( _aCpos , { "WK_CPO34", "C" ,02, 0 } )//Indicador de CPF/CNPJ do Tomador
aAdd( _aCpos , { "WK_CPO35", "C" ,20, 0 } )//CPF/CNPJ do Tomador
aAdd( _aCpos , { "WK_CPO36", "C" ,15, 0 } )//Inscrição Municipal do Tomador
aAdd( _aCpos , { "WK_CPO37", "C" ,15, 0 } )//Inscrição Estadual do Tomador
aAdd( _aCpos , { "WK_CPO38", "C" ,99, 0 } )//Razão Social do Tomador
aAdd( _aCpos , { "WK_CPO39", "C" ,03, 0 } )//Tipo do Endereço do Tomador
aAdd( _aCpos , { "WK_CPO40", "C" ,99, 0 } )//Endereço do Tomador
aAdd( _aCpos , { "WK_CPO41", "C" ,10, 0 } )//Número do Endereço do Tomador
aAdd( _aCpos , { "WK_CPO42", "C" ,25, 0 } )//Complemento do Endereço do Tomador
aAdd( _aCpos , { "WK_CPO43", "C" ,50, 0 } )//Bairro do Tomador
aAdd( _aCpos , { "WK_CPO44", "C" ,50, 0 } )//Cidade do Tomador
aAdd( _aCpos , { "WK_CPO45", "C" ,03, 0 } )//UF do Tomador
aAdd( _aCpos , { "WK_CPO46", "C" ,10, 0 } )//CEP do Tomador
aAdd( _aCpos , { "WK_CPO47", "C" ,50, 0 } )//Email do Tomador;
AADD( _aCpos , { "WK_DSCSER","M" ,10, 0 } )//Discriminação dos Serviços
AADD( _aCpos , { "WK_MOTIVO","M" ,10, 0 } )//OBS
AADD( _aCpos , { "WK_ERRO"  ,"M" ,10, 0 } )//ERROR.LOG
AADD( _aCpos , { "WK_CLICOD","C" ,LEN(SF2->F2_CLIENTE), 0 } )//CODIGO DO CLIENTE
AADD( _aCpos , { "WK_CLOLOJ","C" ,LEN(SF2->F2_LOJA   ), 0 } )//LOJA DO CLIENTE
AADD( _aCpos , { "WK_VALCRE","N" ,TamSx3("F2_CREDNFE")[1],TamSx3("F2_CREDNFE")[2] } )//Valor do Crédito
AADD( _aCpos , { "WK_CODITEM","C",LEN(SD2->D2_COD), 0 } )//CODIGO DO PRODUTO
AADD( _aCpos , { "WK_ALIQISS","N",TamSx3("D2_ALIQISS")[1],TamSx3("D2_ALIQISS")[2] } )//Alíquota do ISS Recebe o valor da coluna “Alíquota” da Planilha 
AADD( _aCpos , { "WK_BASEISS","N",TamSx3("D2_BASEISS")[1],TamSx3("D2_BASEISS")[2] } )//Base do ISS Recebe o valor da coluna “Valor dos Serviços” da Planilha ( - ) o valor da coluna “Valor das Deduções” da Planilha 
AADD( _aCpos , { "WK_VALISS" ,"N",TamSx3("D2_VALISS" )[1],TamSx3("D2_VALISS" )[2] } )//Valor do ISS
AADD( _aCpos , { "WK_DEDUCO" ,"N",TamSx3("D2_DESCON" )[1],TamSx3("D2_DESCON" )[2] } )//Valor das Deduções
AADD( _aCpos , { "WK_SERVICO","N",TamSx3("D2_TOTAL"  )[1],TamSx3("D2_TOTAL"  )[2] } )//Valor dos Serviços
//AADD( _aCpos , { "WK_TOTAL"  ,"N",TamSx3("D2_TOTAL"  )[1],TamSx3("D2_TOTAL"  )[2] } )//Valor dos Serviços - Valor das Deduções
AADD( _aCpos , { "WK_COD_MUN","C",LEN(SA1->A1_COD_MUN), 0 } )//CODIGO DO MUNICIPIO
AADD( _aCpos , { "WK_TES"    ,"C",LEN(SB1->B1_TS), 0 } )//CODIGO DA TES
    
Return( _aCpos )

//===============================================================================================================================
User Function MFIS001B()
//===============================================================================================================================

Local _nRegAtu	:= 0
Local _nTotReg	:= (cAliasAux)->( LASTREC() )
Local _nRegSel	:= 0
Local _nRegErro	:= 0
Local _nRegNoMar:= 0

ProcRegua(_nTotReg)
(cAliasAux)->( DBSETORDER(0) )
(cAliasAux)->( DBGoTop() )

DO While (cAliasAux)->( !EOF() )

	_nRegAtu++
	IncProc( "Processando "+ StrZero(_nRegAtu,6) +" de "+ StrZero(_nTotReg,6) +", Inc: "+STRZERO(_nRegSel,5)+", Rej: "+STRZERO(_nRegErro,5) )

  	If (cAliasAux)->WKSTATUS <> 'A'
	   (cAliasAux)->( DbSkip() )
	   LOOP
  	ENDIF
	
	oMarkBRW:GoTo( (cAliasAux)->( RECNO() ) , .T. )
	
  	If (cAliasAux)->WKSTATUS = 'A' .AND. oMarkBRW:IsMark()
	   
	   _cErro:=""
	   _cDoc :=""
	   IF MGEraNF()//GERA SF2
	      (cAliasAux)->WKSTATUS :='P'
	      (cAliasAux)->WK_MOTIVO:="NOTA GERADA: "+_cDoc+" / Serie: "+_cSerie+CHR(13)+CHR(10)+(cAliasAux)->WK_MOTIVO
	      _nRegSel++
		  oMarkBRW:MARKREC()
		  oMarkBRW:REFRESH()
	   ELSE
	      _nRegErro++
	      (cAliasAux)->WKSTATUS :='S'
	      (cAliasAux)->WK_MOTIVO:='VEJA O ERRO NO BOTÃO "Erro do MSExecAuto()" / Nota NÃO Gerada: '+_cDoc+" / Serie: "+_cSerie+CHR(13)+CHR(10)+(cAliasAux)->WK_MOTIVO
	      (cAliasAux)->WK_ERRO  :=_cErro
		  oMarkBRW:MARKREC()
		  oMarkBRW:REFRESH()
	   ENDIF	
       (cAliasAux)->MARCA := ""//Desmarca Sempre
    ELSE
       _nRegNoMar++
	ENDIF	
	
	(cAliasAux)->( DbSkip() )
	
EndDo

(cAliasAux)->( DBSETORDER(1) )
(cAliasAux)->( DBGoTop() )

_cMen:=ALLTRIM(STR(_nRegAtu))+" Nota(s) processada(s)."+CHR(13)+CHR(10)
_cMen+=ALLTRIM(STR(_nRegSel))+" Nota(s) Integrada(s) com sucesso."+CHR(13)+CHR(10)
_cMen+=ALLTRIM(STR(_nRegErro))+" Nota(s) com ERRO(S)."+CHR(13)+CHR(10)
_cMen+=ALLTRIM(STR(_nRegNoMar))+" Nota(s) não Marcadas(s)."

Aviso( "Concluído!" , _cMen , {"Fechar"} )

oMarkBRW:GoTop()
oMarkBRW:Refresh(.T.)

Return()
//===============================================================================================================================
STATIC FUNCTION MFIS001L(_nReg,oproc,cFile)
//===============================================================================================================================
Local _nCont 	:= 0
Local _cDados	:= ""
Local _aLinha	:= {}
Local _adados 	:= {}
Local lTemNFE   :=.T. , I
Local _cTes2 := GETMV("MV_D2_TESF2",,"")

oFile:=FWFileReader():New(cFile)
IF !(oFile:Open())
   RETURN .f.
ENDIF
oFile:setBufferSize(4096)
_cDados:=oFile:GetLine()//FT_FREADLN()/PULA A LINHA DE TITULOS

CC2->(DBSETORDER(4))
SA1->(DBSETORDER(3))
SF2->(DBSETORDER(1))
DO While (oFile:hasLine()) //!FT_FEOF()  //FACA ENQUANTO NAO FOR FIM DE ARQUIVO  
	
	_nCont++ 
	oproc:ccaption := ("Lendo Linha " + STRZERO(_nCont,5) + " de " + STRZERO(_nReg,5) + ".")
    ProcessMessages()
			
	_cDados :=oFile:GetLine()//FT_FREADLN()//

	//Verifica se é final de arquivo
	If UPPER(LEFT(ALLTRIM(_cDados),6)) = 'TOTAL;'
		Exit
	Endif 	

	_aLinha :=ITTXTARRAY(_cDados,";")
	
	//Verifica se é linha válida
	If LEN(_aLinha) < 48	
	   If LEN(_aLinha) > 0 .AND. (cAliasAux)->(!EOF())
	      (cAliasAux)->WK_MOTIVO:="Estouro de colunas: "+_cDados
	      (cAliasAux)->WKSTATUS := "R"
	   ELSEIf LEN(_aLinha) = 0 .AND. (cAliasAux)->(!EOF())
	      (cAliasAux)->( DBAPPEND() )
	      (cAliasAux)->WKSTATUS := "R"
	      (cAliasAux)->WK_DSCSER:=_cDados
	   Endif
	   Loop		
	Endif
   
	(cAliasAux)->( DBAPPEND() )
	(cAliasAux)->WKSTATUS 	:= "A"
     FOR I := 1 TO 47
        cConteudo:=(cAliasAux)->( FIELDPOS("WK_CPO"+STRZERO(I,2)) )
        IF !EMPTY(cConteudo)
           (cAliasAux)->( FIELDPUT(cConteudo,_aLinha[I]) )
        ENDIF
     NEXT
     cMotivo:=""

     nValServico:=VAL(STRTRAN(STRTRAN((cAliasAux)->WK_CPO27,".",""),",","."))//Valor dos Serviços da Planilha
     nValDeducao:=VAL(STRTRAN(STRTRAN((cAliasAux)->WK_CPO28,".",""),",","."))//Valor das Deduções da Planilha 
    (cAliasAux)->WK_DEDUCO :=nValDeducao//Valor das Deduções
    (cAliasAux)->WK_SERVICO:=nValServico//Valor dos Serviços
//  (cAliasAux)->WK_TOTAL  :=(nValServico-nValDeducao) //Valor dos Serviços - Valor das Deduções
    (cAliasAux)->WK_BASEISS:=(nValServico-nValDeducao)//Base do ISS Recebe o valor da coluna “Valor dos Serviços” da Planilha ( - ) o valor da coluna “Valor das Deduções” da Planilha 
    (cAliasAux)->WK_VALCRE :=VAL(STRTRAN(STRTRAN((cAliasAux)->WK_CPO32,".",""),",","."))//Valor do Crédito
    (cAliasAux)->WK_ALIQISS:=VAL(STRTRAN(STRTRAN((cAliasAux)->WK_CPO30,".",""),",","."))//Alíquota do ISS Recebe o valor da coluna “Alíquota” da Planilha 
    (cAliasAux)->WK_VALISS :=VAL(STRTRAN(STRTRAN((cAliasAux)->WK_CPO31,".",""),",","."))//Recebe o valor da coluna “ISS devido” da Planilha
    (cAliasAux)->WK_CODITEM:=Posicione("CE1",1,xFilial("CE1")+ALLTRIM((cAliasAux)->WK_CPO29),"CE1_PROISS")//Código do Serviço Prestado na Nota Fiscal
    IF EMPTY((cAliasAux)->WK_CODITEM)//Código do Serviço Prestado na Nota Fiscal
	   (cAliasAux)->WKSTATUS :="R"
       cMotivo+="Codigo ISS sem amarracao de Produto"+CHR(13)+CHR(10)
    ENDIF   

    IF EMPTY((cAliasAux)->WK_CPO34) .OR. !ALLTRIM((cAliasAux)->WK_CPO34) $ '1,2'//Indicador de CPF/CNPJ do Tomador
	   (cAliasAux)->WKSTATUS :="R"
       cMotivo+="Tipo de cliente Invalido ou Inexistente: "+(cAliasAux)->WK_CPO34+CHR(13)+CHR(10)
    ENDIF   

    IF ALLTRIM((cAliasAux)->WK_CPO23) <> 'T'//Situação da Nota Fiscal
	   (cAliasAux)->WKSTATUS :="R"
       cMotivo+="Situação do Documento Inválida: "+(cAliasAux)->WK_CPO23+CHR(13)+CHR(10)
    ENDIF   

    cCNPJ:=STRTRAN(STRTRAN(STRTRAN( ALLTRIM((cAliasAux)->WK_CPO11) ,".",""),"/",""),"-","")//CPF/CNPJ do Tomador
    IF ALLTRIM(SM0->M0_CGC) <> cCNPJ .AND. !("GOIASMINAS" $ UPPER(SM0->M0_NOMECOM))
  	   (cAliasAux)->WKSTATUS :="R" //TESTES
       cMotivo+="Empresa Incorreta para estas NFSe: "+ALLTRIM((cAliasAux)->WK_CPO11)+CHR(13)+CHR(10) //TESTES
    ENDIF   
    _cDoc := STRZERO(VAL(ALLTRIM((cAliasAux)->WK_CPO02)),Len(SD2->D2_DOC))
//  IF SF2->(DBSEEK(xFilial()+ALLTRIM((cAliasAux)->WK_CPO02)))
    IF SF2->(DBSEEK(xFilial()+_cDoc+_cSerie))
	   (cAliasAux)->WKSTATUS :="R"
       cMotivo+="Nota já existe no sistema: "+xFilial("SF2")+" "+ALLTRIM((cAliasAux)->WK_CPO02)+CHR(13)+CHR(10)
    ENDIF   

    _cTes1 := Posicione("SB1",1,xFilial("SB1")+(cAliasAux)->WK_CODITEM,"B1_TS")
    IF !EMPTY(_cTes1)
       (cAliasAux)->WK_TES:=_cTes1
    ELSEIF !EMPTY(_cTes2)
       (cAliasAux)->WK_TES:=_cTes2
    ENDIF
    IF EMPTY((cAliasAux)->WK_TES)
	   (cAliasAux)->WKSTATUS :="R"
       cMotivo+='Produto "'+ALLTRIM((cAliasAux)->WK_CODITEM)+'" e Parametro "MV_D2_TESF2" sem TES'
    ENDIF
    
    cCNPJ:=(cAliasAux)->WK_CPO35:=STRTRAN(STRTRAN(STRTRAN( ALLTRIM((cAliasAux)->WK_CPO35) ,".",""),"/",""),"-","")//CPF/CNPJ do Tomador
    IF !EMPTY(cCNPJ) .AND. SA1->(DBSEEK(xFilial()+ALLTRIM(cCNPJ)))
       (cAliasAux)->WK_CLICOD :=SA1->A1_COD
       (cAliasAux)->WK_CLOLOJ :=SA1->A1_LOJA
       (cAliasAux)->WK_CPO36  :=TRANS(SA1->A1_INSCRM,PesqPict('SA1',"A1_INSCRM"))//Inscrição Municipal do Tomador
       (cAliasAux)->WK_CPO37  :=TRANS(SA1->A1_INSCR ,PesqPict('SA1',"A1_INSCR")) //Inscrição Estadual do Tomador
       (cAliasAux)->WK_COD_MUN:=SA1->A1_COD_MUN//CODIGO DO MUNICIPIO
    ELSEIF (cAliasAux)->WKSTATUS <> "R"
       //1-CPF/CNPJ do Tomador 2-Razão Social do Tomador 3-Endereço do Tomador 4-Cidade do Tomador 5-UF do Tomador 6-CEP do Tomador
       IF EMPTY(cCNPJ)                 .OR. EMPTY((cAliasAux)->WK_CPO38) .OR. EMPTY((cAliasAux)->WK_CPO40) .OR.;
          EMPTY((cAliasAux)->WK_CPO44) .OR. EMPTY((cAliasAux)->WK_CPO45) .OR. EMPTY((cAliasAux)->WK_CPO46)
   	      (cAliasAux)->WKSTATUS  :="R"
          cMotivo+="Dados do Cliente Inválidos ou incompletos"+CHR(13)+CHR(10)
       ELSE
          //Para trazer o conteúdo do campo A1_cod_mun, você busca na CC2_UF o valor da COLUNA UF do TOMADOR + o valor da CC2_MUN da coluna CIDADE do tomador, mais quando pegar o valor desta coluna tem que transformar ele, tirando caracter especial e transformando tudo em maiúsculo... para compar certinho com a CC2 achando a referencia pega o conteúdo do campo CC2_CODMUN       
          _cLugar:=LEFT((cAliasAux)->WK_CPO45,2)+MFISLimpa((cAliasAux)->WK_CPO44)
          IF !EMPTY(_cLugar) .AND. CC2->(DBSEEK(xFilial()+_cLugar))
//           cMotivo+="REVISAR CADASTRO DO CLIENTE APOS IMPORTACAO DA NOTA"+CHR(13)+CHR(10)
             (cAliasAux)->WK_COD_MUN:=CC2->CC2_CODMUN//CODIGO DO MUNICIPIO
   	      ELSE
             cMotivo+="Codigo do Municipio Pendente: "+_cLugar+CHR(13)+CHR(10)
   	         (cAliasAux)->WKSTATUS  :="1"
   	      ENDIF
   	   ENDIF
    ENDIF   

	DO WHILE AT(_aLinha[48],"||") # 0
       _aLinha[48]:=STRTRAN(_aLinha[48],"||","|")
    ENDDO
	(cAliasAux)->WK_DSCSER:=STRTRAN(_aLinha[48],"|"," ")
	(cAliasAux)->WK_MOTIVO:=cMotivo
	
    lTemNFE:=.T.

Enddo

oFile:Close()
//Fecha arquivo e prepara parâmetro com arquivo convertido
//FT_FUSE()	
Return lTemNFE
/*
===============================================================================================================================
Programa----------: ITTXTARRAY
===============================================================================================================================
Descrição---------: Convertge o Texto recebido como parâmetro em Array
===============================================================================================================================
Parametros--------: _cTexto     = Texto a ser convertido.
                    _cSeparador = Caracter utilizado como separador de colunas. 
                    _nNrPosicArray = Numero máximo de posições do Array,
===============================================================================================================================
Retorno-----------: _aRet = Retorna o campo _cTexto no formato de Array.
===============================================================================================================================
*/
STATIC function ITTXTARRAY(_cTexto,_cSeparador,_nNrPosicArray)
Local _aRet := {}
//Local _nI
Local _nPosInc, _nInic, _cColuna
Local _nTamTexto := Len(_cTexto)
Local _nTamColuna
Local _lWhile := .T.

Default _nNrPosicArray := 48

Begin Sequence
   
   If Len(_cTexto) == 0
      Break   
   EndIf
   
   // Exemplo:
   // Numero de colunas a serem lidas do arquivo texto.
   //        1              2            3             4                 5               6
   //"CNPJ_FORNECEDOR;CODIGO_CLIENTE;NOME_CLIENTE;NUMERO_DOCUMENTO;VALOR_DOCUMENTO;NOVO_VENCIMENTO"
   
    _nPosInc := 1 
	Do While _lWhile // .T.
	   _nInic := At(_cSeparador, _cTexto , _nPosInc) 
        
       If _nInic > 0
          _nPosInc := _nInic + 1
       Else
          Exit  
       EndIf
       
       If Len(_aRet) == 0 
          _cColuna := SubStr(_cTexto,1,_nInic-1)
          Aadd(_aRet,_cColuna)
       EndIf  
       
	   _nFin  := 0
	
	   _nFin := At(_cSeparador,_cTexto,_nInic+1)
	   
	   If _nFin == 0
	      _lWhile := .F.
	   EndIf
	   
	   If _nFin > 0
	      _nTamColuna := _nFin - (_nInic + 1)
	      _cColuna := SubStr(_cTexto,_nInic+1,_nTamColuna)
	   Else
	      _nTamColuna := _nTamTexto - _nInic // (_nInic + 1)
	      _cColuna := SubStr(_cTexto,_nInic+1,_nTamColuna)
	   EndIf
	   
	   Aadd(_aRet,_cColuna)
	   
	   If Len(_aRet) == _nNrPosicArray  // Numero máximo de colunas da planilha gravada em CSV.
	      _lWhile := .F.  //Exit
	   EndIf
	   
    EndDo 

End Sequence

Return _aRet

*===============================================================================================================================*
Static function MGEraNF()
*===============================================================================================================================*
Local aCab	 := {}
Local aLinha := {}
Local aItens := {}

//-- Verifica o ultimo documento valido para um fornecedor
//SF2->(dbSetOrder(2) )
//SF2->(MsSeek(xFilial("SF2")+LEFT( (cAliasAux)->WK_CLICOD,Len(SF2->F2_CLIENTE) )+"ZZ",.T.))
//SF2->(dbSkip(-1) )
//_cDoc:= SF2->F2_DOC
//If Empty(_cDoc)
_cDoc  := STRZERO(VAL(ALLTRIM((cAliasAux)->WK_CPO02)),Len(SD2->D2_DOC))
//Else
//   _cDoc := Soma1(_cDoc)
//EndIf

SF2->(DBSETORDER(1))
IF SF2->(DBSEEK(xFilial()+_cDoc+_cSerie))
   _cErro:="Nota já existe no sistema: "+xFilial("SF2")+" "+ALLTRIM((cAliasAux)->WK_CPO02)+CHR(13)+CHR(10)
   RETURN .F.
ENDIF   

cCNPJ:=(cAliasAux)->WK_CPO35:=STRTRAN(STRTRAN(STRTRAN( ALLTRIM((cAliasAux)->WK_CPO35) ,".",""),"/",""),"-","")//CPF/CNPJ do Tomador
SA1->(DBSETORDER(3))
IF !EMPTY(cCNPJ) .AND. SA1->(DBSEEK(xFilial()+ALLTRIM(cCNPJ)))
   (cAliasAux)->WK_CLICOD :=SA1->A1_COD
   (cAliasAux)->WK_CLOLOJ :=SA1->A1_LOJA
   (cAliasAux)->WK_CPO36  :=TRANS(SA1->A1_INSCRM,PesqPict('SA1',"A1_INSCRM"))//Inscrição Municipal do Tomador
   (cAliasAux)->WK_CPO37  :=TRANS(SA1->A1_INSCR ,PesqPict('SA1',"A1_INSCR")) //Inscrição Estadual do Tomador
   (cAliasAux)->WK_COD_MUN:=SA1->A1_COD_MUN//CODIGO DO MUNICIPIO
ENDIF   

AAdd( aCab, { "F2_CLIENTE", (cAliasAux)->WK_CLICOD     	, Nil } )	
AAdd( aCab, { "F2_LOJA"   , (cAliasAux)->WK_CLOLOJ		, Nil } )	
AAdd( aCab, { "F2_SERIE"  , _cSerie						, Nil } )
AAdd( aCab, { "F2_DOC"    ,	_cDoc                       , Nil } )	  		 
AAdd( aCab, { "F2_COND"   , "001"	   					, Nil } )	 
AAdd( aCab, { "F2_EMISSAO", CTOD(ALLTRIM((cAliasAux)->WK_CPO08)), Nil } )//Data do Fato Gerador
//AAdd( aCab, { "F2_EST"  , "01"		        		, Nil } )  	
AAdd( aCab, { "F2_TIPO"   , "N"		        	       	, Nil } )  
aadd( aCab, { "F2_FORMUL" , "N"		        	       	, Nil } )  
AAdd( aCab, { "F2_ESPECIE", "RPS"   			        , Nil } ) 
//AAdd( aCab, { "F2_PREFIXO", "UNI"		        		, Nil } ) 
AAdd( aCab, { "F2_MOEDA"  , 1		            	   	, Nil } ) 
//AAdd( aCab, { "F2_TXMOEDA", 1		               		, Nil } ) 
AAdd( aCab, { "F2_FORMUL" , "S"		   		           	, Nil } ) 
AAdd( aCab, { "F2_TIPODOC", ""			     			, Nil } ) 
AAdd( aCab, { "F2_NFELETR", ALLTRIM((cAliasAux)->WK_CPO02)         , Nil } )//Nº NFS-e
AAdd( aCab, { "F2_CODNFE" , ALLTRIM((cAliasAux)->WK_CPO04)         , Nil } )//Código de Verificação da NFS-e
AAdd( aCab, { "F2_EMINFE" , CTOD(LEFT((cAliasAux)->WK_CPO03,10))   , Nil } )//Data Hora NFE
AAdd( aCab, { "F2_HORNFE" , RIGHT(ALLTRIM((cAliasAux)->WK_CPO03),5), Nil } )//Data Hora NFE
AAdd( aCab, { "F2_CREDNFE", (cAliasAux)->WK_VALCRE                 , Nil } )//Valor do Crédito
AAdd( aCab, { "F2_MENNOTA", ALLTRIM((cAliasAux)->WK_DSCSER)        , Nil } )//Discriminação dos Serviços
AAdd( aCab, { "F2_DESCONT",0})
AAdd( aCab, { "F2_FRETE"  ,0})
AAdd( aCab, { "F2_SEGURO" ,0})
AAdd( aCab, { "F2_DESPESA",0})

aLinha := {}				
AAdd( aLinha, { "D2_COD"    , (cAliasAux)->WK_CODITEM, Nil } )//Código do Serviço Prestado na Nota Fiscal
AAdd( aLinha, { "D2_QUANT"  , 1 					 , Nil } )					
//AAdd( aLinha, { "D2_PRCVEN" , (cAliasAux)->WK_SERVICO, Nil } )//Valor dos Serviços			
AAdd( aLinha, { "D2_PRCVEN" , (cAliasAux)->WK_BASEISS, Nil } )  //Valor dos Serviços			
//AAdd( aLinha, { "D2_DESCON" , (cAliasAux)->WK_DEDUCO , Nil } )  //Valor das Deduções			
//AAdd( aLinha, { "D2_TOTAL"  , (cAliasAux)->WK_SERVICO, Nil } )//Valor dos Serviços			
AAdd( aLinha, { "D2_TOTAL"  , (cAliasAux)->WK_BASEISS, Nil } )  //Valor dos Serviços			
AAdd( aLinha, { "D2_TES"    , (cAliasAux)->WK_TES	 , Nil } )
//AAdd( aLinha, { "D2_UM"     , "UN" 				 , Nil } )
AAdd( aLinha, { "D2_ESPECIE", "RPS"   		    	 , Nil } )
AAdd( aLinha, { "D2_ALIQISS", (cAliasAux)->WK_ALIQISS, Nil } )//Alíquota do ISS Recebe o valor da coluna “Alíquota” da Planilha 
AAdd( aLinha, { "D2_BASEISS", (cAliasAux)->WK_BASEISS, Nil } )//Base do ISS Recebe o valor da coluna “Valor dos Serviços” da Planilha ( - ) o valor da coluna “Valor das Deduções” da Planilha 
AAdd( aLinha, { "D2_VALISS" , (cAliasAux)->WK_VALISS , Nil } )//Valor do ISS
_cConta := Posicione("SB1",1,xFilial("SB1")+(cAliasAux)->WK_CODITEM,"B1_CONTA")
AAdd( aLinha, { "D2_CONTA", _cConta   		    	 , Nil } )

AAdd( aItens, aLinha)

Private lMsErroAuto := .F.


IF EMPTY((cAliasAux)->WK_CLICOD)

  _aCliente := DadosImpor()
  
  BEGIN TRANSACTION
	
	MSExecAuto( {|x,y,z| mata030(x,y,z) } , _aCliente ,, 3 ) // SIGA AUTO PARA A INCLUSAO DO CLIENTE
	
	aCab[1,2]:=(cAliasAux)->WK_CLICOD:=SA1->A1_COD
	aCab[2,2]:=(cAliasAux)->WK_CLOLOJ:=SA1->A1_LOJA
	
	If lMsErroAuto
		
		DisarmTransaction()
		_cErro:="Erro Cliente: ["+ALLTRIM(MostraErro())+"]"+CHR(13)+CHR(10)
		
	EndIf
  
  END TRANSACTION
	
ENDIF

If !lMsErroAuto
	
  BEGIN TRANSACTION

    MSExecAuto({|x,y,z| mata920(x,y,z)} , aCab, aItens, 3 ) //Inclusao 
	
	If lMsErroAuto
		DisarmTransaction()
		_cErro+="Erro Nota: ["+ALLTRIM(MostraErro())+"]"+CHR(13)+CHR(10)
	EndIf

  END TRANSACTION
	
EndIf


RETURN !lMsErroAuto

*===============================================================================================================================*
Static Function DadosImpor()
*===============================================================================================================================*
Local _aCliente	:= {}
Local cCEP := STRTRAN(AllTrim((cAliasAux)->WK_CPO46),"-","")
Local cEnd := AllTrim((cAliasAux)->WK_CPO39)+" "+AllTrim((cAliasAux)->WK_CPO40)+" "+AllTrim((cAliasAux)->WK_CPO41)
Local cPes := IF(ALLTRIM((cAliasAux)->WK_CPO34) = "2","J","F")
Local cNat := GETMV("MV_NATCLIF2",,"111001")

CC2->(DBSETORDER(1))
SA1->(DBSETORDER(1))

aAdd( _aCliente , { "A1_FILIAL"	, xFilial("SA1")				, Nil }) // FILIAL
aAdd( _aCliente , { "A1_NOME"	, AllTrim((cAliasAux)->WK_CPO38), Nil }) // NOME //Razão Social do Tomador
aAdd( _aCliente , { "A1_PESSOA"	, cPes							, Nil }) // PESSOA FISICA OU JURIDICA
aAdd( _aCliente , { "A1_CGC"	, ALLTRIM((cAliasAux)->WK_CPO35), Nil }) // CGC
aAdd( _aCliente , { "A1_NREDUZ"	, AllTrim((cAliasAux)->WK_CPO38), Nil }) // NOME REDUZIDO
aAdd( _aCliente , { "A1_TIPO"	, "F"							, Nil }) // TIPO DE CLIENTE
aAdd( _aCliente , { "A1_EST"	, AllTrim((cAliasAux)->WK_CPO45), Nil }) // ESTADO - UF do Tomador
aAdd( _aCliente,  { "A1_COD_MUN", (cAliasAux)->WK_COD_MUN		, Nil }) // COD.MUNICIPIO
aAdd( _aCliente , { "A1_CEP"	, cCEP							, Nil }) // CEP
aAdd( _aCliente , { "A1_END"	, cEnd							, Nil }) // ENDEREÇO
aAdd( _aCliente , { "A1_BAIRRO"	, AllTrim((cAliasAux)->WK_CPO43), Nil }) // BAIRRO
aAdd( _aCliente , { "A1_COMPLEM", AllTrim((cAliasAux)->WK_CPO42), Nil }) // Complemento do Endereço do Tomador
aAdd( _aCliente , { "A1_PAIS"	, "105"							, Nil }) // PAIS
aAdd( _aCliente , { "A1_CODPAIS", "01058"						, Nil }) // PAIS BACEN
aAdd( _aCliente , { "A1_INSCR"	, ""							, Nil }) // INSCRICAO ESTADUAL Inscrição Estadual do Tomador
aAdd( _aCliente , { "A1_INSCRM"	, ""							, Nil }) // Inscrição Municipal do Tomador
aAdd( _aCliente , { "A1_NATUREZ", cNat							, Nil }) // NATUREZA
aAdd( _aCliente , { "A1_EMAIL"	, AllTrim((cAliasAux)->WK_CPO47), Nil }) // EMAIL - Email do Tomador

IF SA1->(FIELDPOS("A1_I_CMUNC")) <> 0
   AADD( _aCliente ,{"A1_ESTC"	  , AllTrim((cAliasAux)->WK_CPO45)	, Nil }) // ESTADO COBRANCA
   AADD( _aCliente ,{"A1_I_CMUNC", (cAliasAux)->WK_COD_MUN			, Nil }) // COD.MUNICIPIO COBRANCA
   AADD( _aCliente ,{"A1_CEPC"	  , cCEP							, Nil }) // CEP COBRANCA
   AADD( _aCliente ,{"A1_ENDCOB ", cEnd						   		, Nil }) // ENDERECO COBRANCA
   AADD( _aCliente ,{"A1_BAIRROC", AllTrim((cAliasAux)->WK_CPO43)	, Nil }) // BAIRRO COBRANCA
   _cDDD :="999"
   _cTel :="99999999"
   AADD( _aCliente , { "A1_DDD"	, _cDDD						   		, Nil }) // DDD DO TELEFONE
   AADD( _aCliente , { "A1_TEL"	, _cTel						  		, Nil }) // NUMERO DO TELEFONE
   AADD( _aCliente ,{"A1_I_GRCLI", "11"						  		, Nil }) // GRUPO CLIENTE
   _cCodVend:= "000156"
   AADD( _aCliente , { "A1_VEND"		, _cCodVend					, Nil }) // CODIGO DO VENDEDOR
   AADD( _aCliente , { "A1_GRPVEN"	, "999999"						, Nil }) // GRUPO DE VENDAS
   _cRisco 	:= 	U_ITGETMV( "IT_RISCOFUN" , "B" )
   _nLimite 	:=	U_ITGETMV( "IT_LIMFUNC" , 150 )
   _dLimite 	:=	U_ITGETMV( "IT_VENCLIMFUNC" , stod("20491231") )
   AADD( _aCliente , { "A1_RISCO"	, _cRisco						, Nil }) // RISCO CLIENTE
   AADD( _aCliente , { "A1_LC"		, _nLimite						, Nil }) // VALOR DO LIMITE
   AADD( _aCliente , { "A1_VENCLC"	, _dLimite						, Nil }) // DATA DE VENCIMENTO DO LIMITE
   _cCContabil := "1102069998"
   AADD( _aCliente , { "A1_CONTA"	, _cCContabil		 			, Nil }) // Conta Contabil
   AADD( _aCliente , { "A1_COND"		, "001"		  				, Nil }) // CONDICAO DE PAGTO INCLUSÃO 
   AADD( _aCliente , { "A1_CONTRIB"	, "2"					   		, Nil }) // Contribuinte do ICMS
   AADD( _aCliente , { "A1_CLIFUN"	, "1"							, Nil }) // Funcionário
ENDIF

Return( _aCliente )

*--------------------------------------------------------------------------------------------*
Static Function MFISLimpa(cTexto)//AWF - 27/04/2018 - Tira os caracteres "estranos"
*--------------------------------------------------------------------------------------------*
   cTexto:=UPPER(ALLTRIM(cTexto))
   cTexto:=StrTran(cTexto,"Ã?","E")
   cTexto:=StrTran(cTexto,"Ã^","E")
   cTexto:=StrTran(cTexto,"^","")
   cTexto:=StrTran(cTexto,"á","a")
   cTexto:=StrTran(cTexto,"Á","A")
   cTexto:=StrTran(cTexto,"à","a")
   cTexto:=StrTran(cTexto,"À","A")
   cTexto:=StrTran(cTexto,"ã","a")
   cTexto:=StrTran(cTexto,"Ã","A")
   cTexto:=StrTran(cTexto,"â","a")
   cTexto:=StrTran(cTexto,"Â","A")
   cTexto:=StrTran(cTexto,"ä","a")
   cTexto:=StrTran(cTexto,"Ä","A")
   cTexto:=StrTran(cTexto,"é","e")
   cTexto:=StrTran(cTexto,"É","E")
   cTexto:=StrTran(cTexto,"ë","e")
   cTexto:=StrTran(cTexto,"Ë","E")
   cTexto:=StrTran(cTexto,"ê","e")
   cTexto:=StrTran(cTexto,"Ê","E")
   cTexto:=StrTran(cTexto,"í","i")
   cTexto:=StrTran(cTexto,"Í","I")
   cTexto:=StrTran(cTexto,"ï","i")
   cTexto:=StrTran(cTexto,"Ï","I")
   cTexto:=StrTran(cTexto,"î","i")
   cTexto:=StrTran(cTexto,"Î","I")
   cTexto:=StrTran(cTexto,"ý","y")
   cTexto:=StrTran(cTexto,"Ý","y")
   cTexto:=StrTran(cTexto,"ÿ","y")
   cTexto:=StrTran(cTexto,"ó","o")
   cTexto:=StrTran(cTexto,"Ó","O")
   cTexto:=StrTran(cTexto,"õ","o")
   cTexto:=StrTran(cTexto,"Õ","O")
   cTexto:=StrTran(cTexto,"ö","o")
   cTexto:=StrTran(cTexto,"Ö","O")
   cTexto:=StrTran(cTexto,"ô","o")
   cTexto:=StrTran(cTexto,"Ô","O")
   cTexto:=StrTran(cTexto,"ò","o")
   cTexto:=StrTran(cTexto,"Ò","O")
   cTexto:=StrTran(cTexto,"ú","u")
   cTexto:=StrTran(cTexto,"Ú","U")
   cTexto:=StrTran(cTexto,"ù","u")
   cTexto:=StrTran(cTexto,"Ù","U")
   cTexto:=StrTran(cTexto,"ü","u")
   cTexto:=StrTran(cTexto,"Ü","U")
   cTexto:=StrTran(cTexto,"ç","c")
   cTexto:=StrTran(cTexto,"Ç","C")
   cTexto:=StrTran(cTexto,"º","o")
   cTexto:=StrTran(cTexto,"°","o")
   cTexto:=StrTran(cTexto,"ª","a")
   cTexto:=StrTran(cTexto,"ñ","n")
   cTexto:=StrTran(cTexto,"Ñ","N")
   cTexto:=StrTran(cTexto,"§","S")
   cTexto:=StrTran(cTexto,"o","o")
   cTexto:=StrTran(cTexto,"µ","u")
   cTexto:=StrTran(cTexto,"&","e") 
   cTexto:=StrTran(cTexto,"¡","i")
   cTexto:=StrTran(cTexto,"×","x")
Return cTexto

*===============================================================================================================================*
USER FUNCTION MFISMUN()
*===============================================================================================================================*
LOCAL lREt:=.T.
PRIVATE M->A1_EST:=""

//IF !oMarkBRW:IsMark() //Testa pq do duplo clique ele chama essa função 2 veses OR botão Altertar
	
//	IF (cAliasAux)->WKSTATUS = "1"
		IF EMPTY((cAliasAux)->WK_CPO45)
			MSGSTOP("Nota sem UF do Tomador")
			RETURN .F.
		ELSE
			M->A1_EST:=AllTrim((cAliasAux)->WK_CPO45)//Inicia essa vairavel para filtra o F3 'CC2SA1'
		ENDIF
		lREt:=ConPad1(,,,'CC2SA1',,)
		IF lREt
			(cAliasAux)->WKSTATUS    := "A"
			(cAliasAux)->WK_COD_MUN  := CC2->CC2_CODMUN
			//(oMarkBRW:Alias())->MARCA:= oMarkBRW:Mark()
			//(cAliasAux)->MARCA       := oMarkBRW:Mark()
		    oMarkBRW:Refresh(.T.)
		ENDIF
/*	ELSE
		(oMarkBRW:Alias())->MARCA:= oMarkBRW:Mark()
	ENDIF
ELSE
	(oMarkBRW:Alias())->MARCA := ""
ENDIF
*/
RETURN lREt
