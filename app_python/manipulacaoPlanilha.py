# carregando pacotinhos
import pandas as pd
import datetime as dt
import xlsxwriter
import io
import zipfile

#algumas funções importantes
def gerar_zipfile(arquivos, nomes_arquivos):
    conter = 0
    mem_zip = io.BytesIO()
    with zipfile.ZipFile(mem_zip, mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
        for posi, arquivo in enumerate(arquivos):
            conter += 1
            zf.writestr(nomes_arquivos[posi], arquivo)
    return mem_zip.getvalue()
    
    
def receber_planilha():
    dadosPagamento = pd.read_excel('Honorarios_medicos_estudos_DASA_HBR_LSUL.xlsx', 'atendimentos_realizados')
    dadosMedicos = pd.read_excel('Honorarios_medicos_estudos_DASA_HBR_LSUL.xlsx', 'dados_medicos')
    dadosMedicosRelevantes = dadosMedicos[['Medico', 'CPF', 'Nome_CNPJ', 'CNPJ']]
    #dadosPagamento = dadosPagamento[(dadosPagamento['Enviado'] == 'NAO')]
    dadosPagamento = dadosPagamento[(dadosPagamento['Enviado'] == 'NAO')&(dadosPagamento['Confirmado'] == 'SIM')]
    dadosPagamento['Mes de Atendimento'] = pd.to_datetime(dadosPagamento['Data_de_realizacao']).dt.to_period('M')
    dadosPagamento.join(dadosMedicosRelevantes.set_index('Medico'), on="Medico", how='inner')
    # retorna um dataframe pandas
    return dadosPagamento
    
    
def gerar_planilhas(dadosPagamento):
    # data de hoje
    hoje = dt.date.today()
    # inicializando array de planilhas
    planilhas = []
    arquivos_names = []
    # capturando nomes dos médicos
    medicos = pd.unique(dadosPagamento['Medico'])
    # colunas que serao usadas para pagamento e sinalizadas como produção
    colunas_pagamento = dadosPagamento[['Medico', 'Centro', 'Estudo', 'tipo_atendimento','Mes de Atendimento', 'PID', 'Visita_decricao', 'Valor']]
    # loop para filtro de dados
    for medico in medicos:
        df_pagamento = colunas_pagamento['Medico'] == medico
        filtrado = colunas_pagamento[df_pagamento]
        agrupado = filtrado.groupby(['Medico', 'Centro', 'Estudo', 'Mes de Atendimento', 'tipo_atendimento', 'Valor']).agg(
                                                                                                                        contagem=pd.NamedAgg(column="tipo_atendimento", aggfunc="count"),
                                                                                                                        total=pd.NamedAgg(column="Valor", aggfunc="sum"))
           
        nome_arquivo = 'Honorario_medico_' + medico + '_' + str(hoje) + '.xlsx'
        arquivos_names.append(nome_arquivo)
        
        # workbook = xlsxwriter.Workbook(nome_arquivo)
        # worksheet_pagar = workbook.add_worksheet('Pagar')

        # for i in range(0,agrupado.shape[0]):
        #    worksheet_pagar.write('A'+str(i), str(agrupado.iloc[i:1]))
        # workbook.close()

        # inicializando buffer
        buffer = io.BytesIO()
        # salvando dados em planilhas
        with pd.ExcelWriter(buffer) as writer:  
            agrupado.to_excel(writer, sheet_name='Pagar')
            filtrado.to_excel(writer, sheet_name='O que está sendo pago')
        # adicionando a planilha no array
        planilhas.append(buffer.getvalue())
        buffer.close()
    
    return planilhas, arquivos_names, medicos

def gerar_texto_email(nomes):
    buffer = io.BytesIO()
    txt = open(buffer, 'w')
    txt.writer("Olá, seguem os pagamentos médicos para esses mês.\n Os pagamentos são referentes ao médicos:\n")
    for nome in nomes:
        txt.writer(nome)
    txt.writer("Favor confirmar o recebimento deste e-mail.")
    txt.close()
    return buffer.get_value()
    

    
    
dados = receber_planilha()
planilhas, nomes, medicos = gerar_planilhas(dados)
zipfile = gerar_zipfile(planilhas, nomes)

tempfile = "planilas_pagamento_medico_do_mes.zip"

with open(tempfile, 'wb') as f:
    f.write(zipfile)

        
