#carregamento de bibliotecas
using DataFrames    #para trabalhar com DataFrames
using Plots         #para plotagem de dados
using OdsIO         #leitura de arquivos ODS
using StatsBase     #para atividades estatísticas
using Dates         #manipulação de datas
using Missings      #manipulação de dados faltantes
using XLSX          #manipulação de planilha no formato EXCEL
using DataStructures


#leitura de dados
dados = ods_readall("/home/egarbin/Documentos/Faturamento_medico_analise/Honorarios_medicos_estudos_DASA.ods", sheetsNames=["atendimentos_realizados", "dados_medicos"], innerType="DataFrame")
dadosPagamento = dados["atendimentos_realizados"]
dadosMedicos = dados["dados_medicos"]
dadosMedicos = dadosMedicos[!, [:CPF, :CNPJ, :Medico, :Nome_CNPJ]]
DadosJoin = innerjoin(dadosPagamento, dropmissing(dadosMedicos), on = :Medico)
DadosJoin.Data_de_realizacao = Date.(DadosJoin.Data_de_realizacao, DateFormat("y-m-d"))
DadosJoin.Periodo = Dates.format.(DadosJoin.Data_de_realizacao, "U-yyyy")

#filtro de procedimentos ainda NÃO pagamentos
naoPagos = filter(p -> p.Pago .== "NAO", DadosJoin)
Planilha_de_pagamento = naoPagos[!, [:Centro, :Medico, :CPF, :Periodo, :CNPJ, :Estudo, :tipo_atendimento, :Valor]]
rename!(Planilha_de_pagamento, :tipo_atendimento => :Procedimento)
Planilha_de_pagamento
#captura de dados unicos pendentes de pagamento
estudos = skipmissing(unique(dadosPagamento.Estudo))
medicos = collect(skipmissing(unique(dadosPagamento.Medico)))

#criando agrupadores importantes em dicionários
agrupador = Dict()
dadosFinais = Dict()

#Lista de agrupamento
lista = [:Centro, :Medico, :CPF, :Periodo, :CNPJ, :Estudo, :Procedimento, :Valor]

#filtrando base de dados
for medico in medicos
    tabMedico = filter(p -> p.Medico .== medico, Planilha_de_pagamento)
    oQueEstaSendoPago = filter(p -> p.Medico .== medico, naoPagos)
    df_g = groupby(tabMedico, lista)
    combinado = combine(df_g, :Procedimento => length => :Quantidade, :Valor .=> sum => :Total)
    println(combinado)

    #escrevendo os dados em arquivos XLSX
    nomeArquivo = "/home/egarbin/Documentos/Pagamento_medico__" * medico * ".xlsx"
    #gravando os dados na planilha
    XLSX.writetable(nomeArquivo, "Para pagar" => combinado, "O que está sendo pago" => oQueEstaSendoPago, overwrite=true)
end

