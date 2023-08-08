

def concatenar(dados):
    strings_conc = ", ".join(dados)
    return strings_conc


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
    valores_texto_email = []
    # capturando nomes dos médicos
    medicos = pd.unique(dadosPagamento['Medico'])
    # colunas que serao usadas para pagamento e sinalizadas como produção
    colunas_pagamento = dadosPagamento[['Medico', 'Centro', 'Estudo', 'Tipo_atendimento', 'Mes de Atendimento', 'PID', 'Visita_descricao', 'Valor']]
    # loop para filtro de dados
    for medico in medicos:
        df_pagamento = colunas_pagamento['Medico'] == medico
        filtrado = colunas_pagamento[df_pagamento]
        agrupado = filtrado.groupby(['Medico', 'Centro', 'Estudo', 'Mes de Atendimento', 'Tipo_atendimento', 'Valor']).agg(
                                                                                                                        Quantidade=pd.NamedAgg(column="Tipo_atendimento", aggfunc="count"),
                                                                                                                        Total=pd.NamedAgg(column="Valor", aggfunc="sum"),
                                                                                                                        Visitas=pd.NamedAgg(column="Visita_descricao", aggfunc=concatenar),
                                                                                                                        Participantes=pd.NamedAgg(column="PID", aggfunc=concatenar))
           
        nome_arquivo = 'Honorario_medico_' + medico + '_' + str(hoje) + '.xlsx'
        arquivos_names.append(nome_arquivo)
        
   
        # inicializando buffer
        buffer = io.BytesIO()
        # salvando dados em planilhas
        with pd.ExcelWriter(buffer) as writer:  
            agrupado.to_excel(writer, sheet_name='Pagar')
            filtrado.to_excel(writer, sheet_name='O que está sendo pago')

            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet_1 = writer.sheets['Pagar']
            worksheet_2 = writer.sheets['O que está sendo pago']

            # Set the column widths
            worksheet_1.set_column('A:A', 50)
            worksheet_1.set_column('B:B', 15)
            worksheet_1.set_column('C:D', 20)
            worksheet_1.set_column('E:E', 50)
            worksheet_1.set_column('F:F', 8)
            worksheet_1.set_column('G:G', 20)
            worksheet_1.set_column('H:H', 10)

             # Set the column widths
            worksheet_2.set_column('A:A', 50)
            worksheet_2.set_column('B:B', 15)
            worksheet_2.set_column('C:D', 20)
            worksheet_2.set_column('E:E', 35)
            worksheet_2.set_column('F:F', 8)
            worksheet_2.set_column('G:G', 20)
            worksheet_2.set_column('H:H', 50)

            # Write the sum of the 'Valor' column to the bottom of the column
            worksheet_1.write(len(df_pagamento)+1, 6, 'Total a pagar:')
            worksheet_1.write(len(df_pagamento)+1, 7, agrupado['Total'].sum())
            
        valores_texto_email.append([medico, agrupado['Total'].sum()])
        print(valores_texto_email)
         

        # adicionando a planilha no array
        planilhas.append(buffer.getvalue())
        buffer.close()
    



    return planilhas, arquivos_names, valores_texto_email

def gerar_texto_email(dados_texto):
    buffer = io.StringIO()
      # Write text to the buffer
    buffer.write('Olá, seguem os pagamentos médicos para esse mês.\nOs pagamentos são referentes a atendimento em ensaios clínicos para os médicos:\n\n')
    for info in dados_texto:
        buffer.write(f'Medico: {info[0]}. Valor a ser pago: R${info[1]}\n')
    buffer.write("\nFavor confirmar o recebimento deste e-mail.")

    return buffer.getvalue()

  

    
    
dados = receber_planilha()
arquivos, nomes, valor_texto = gerar_planilhas(dados)
arquivos.append(gerar_texto_email(valor_texto))
nomes.append('Texto_para_email.txt')

zipfile = gerar_zipfile(arquivos, nomes)

tempfile = "planilas_pagamento_medico_do_mes.zip"

with open(tempfile, 'wb') as f:
    f.write(zipfile)

        


