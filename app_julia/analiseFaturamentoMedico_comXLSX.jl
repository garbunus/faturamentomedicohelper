#carregamento de bibliotecas
using ZipFile
using DataFrames    #para trabalhar com DataFrames
#using Plots         #para plotagem de dados
#using StatsBase     #para atividades estatísticas
using Dates         #manipulação de datas
using Missings      #manipulação de dados faltantes
using XLSX          #manipulação de planilha no formato EXCEL
#using DataStructures

function leitura_de_planilhas()
	#leitura de dados
	caminho = joinpath(dirname(@__FILE__), "Honorarios_medicos_estudos_DASA_HBR_LSUL.xlsx")
	print("Usando os arquivos na pasta: " * caminho)
	#lendo arquivo xlsx
	dados = DataFrame(XLSX.readtable(caminho, "atendimentos_realizados"))
	dados_2 = DataFrame(XLSX.readtable(caminho, "dados_medicos"))
	return dados, dados_2
end


function manipulacao(dados_1, dados_2)
	dadosPagamento = dados_1
	dadosMedicos = dados_2
	print(dados_1)
	dadosMedicos = dadosMedicos[!, [:CPF, :CNPJ, :Medico, :Nome_CNPJ]]
	DadosJoin = innerjoin(dadosPagamento, dropmissing(dadosMedicos), on = :Medico)
	DadosJoin.Data_de_realizacao = Date.(DadosJoin.Data_de_realizacao)
	DadosJoin.Periodo = Dates.format.(DadosJoin.Data_de_realizacao, "mm-yyyy")

	#filtro de procedimentos ainda NÃO enviados para pagamento
	naoPagos = filter(p -> p.Enviado .== "NAO", DadosJoin)
	
	print("Total de pagamentos ainda não realizados" * count(eachrow(naoPagos)))
	
	planilha_de_pagamento = naoPagos[!, [:Centro, :Medico, :CPF, :Periodo, :CNPJ, :Estudo, :tipo_atendimento, :Valor]]
	rename!(Planilha_de_pagamento, :tipo_atendimento => :Procedimento)
	planilha_de_pagamento
	#captura de dados únicos pendentes de pagamento
	estudos = skipmissing(unique(dadosPagamento.Estudo))
	medicos = collect(skipmissing(unique(Planilha_de_pagamento.Medico)))

	#criando agrupadores importantes em dicionários
	agrupador = Dict()
	dadosFinais = Dict()
	mesAtual = Dates.format(Dates.today(), "mm-yyyy")

	#Lista de agrupamento
	lista = [:Centro, :Medico, :CPF, :Periodo, :CNPJ, :Estudo, :Procedimento, :Valor]

	#filtrando base de dados
	for medico in medicos
	    tabMedico = filter(p -> p.Medico .== medico, Planilha_de_pagamento)
	    oQueEstaSendoPago = filter(p -> p.Medico .== medico, naoPagos)
	    df_g = groupby(tabMedico, lista)
	    combinado = combine(df_g, :Procedimento => length => :Quantidade, :Valor .=> sum => :Total)
	    #println(combinado)
	    #criando pastas para os médicos
	    mkpath(pwd() * "/planilhas/" * medico * "/")
	    #escrevendo os dados em arquivos XLSX
	    nomeArquivo = dirname(@__FILE__) * "/planilhas/" * medico * "/Pagamento_medico__" * medico * "__" * mesAtual * ".xlsx"
	    
	    #adicionar a soma de procedimentos e valores
	    #configurar colunas para o tamanho do texto
	    
	    #gravando os dados na planilha
	    XLSX.writetable(nomeArquivo, "Para pagar" => combinado, "O que está sendo pago" => oQueEstaSendoPago, overwrite=true)
	end
	println("passou dentro da manipulação de arquivos")
end

function zipFileCreator(arquivos)
	# criando o buffer para receber o zip
	buff = IOBuffer()
	# criando um arquivo zip em buffer de memória
	zipFileBuffer = ZipFile(buff, "w")
	
	for arquivo in arquivos
		add!(zipFileBuffer, arquivo)
	close(zipFileBuffer)
	end
end



dados, dados_2 = leitura_de_planilhas()
manipulacao(dados, dados_2)



println("Planilhas criadas com sucesso.")

