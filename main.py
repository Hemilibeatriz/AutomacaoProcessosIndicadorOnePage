'''
    Instalado pelo próprio PyCharm
        install package pandas
    Instalado pelo File> Setting> Project: Python Interpreter > + >:
        openpyxl
        pywin32
'''

#importar bibliotecas
import pandas as pd
import win32com.client as win32
import pathlib

#Passo 1: importar base de dados
emails = pd.read_excel(r"Bases de Dados\Emails.xlsx")
lojas = pd.read_csv(r"Bases de Dados\Lojas.csv", encoding='latin1', sep=';')
vendas = pd.read_excel(r"Bases de Dados\Vendas.xlsx")

#definição das metas
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

#na planilha de vendas está por "id loja" que vem da planilha lojas, vou mesclar(merge) para na planilha vendas aparecer o nome da loja
vendas = vendas.merge(lojas, on='ID Loja')
#print(vendas)

#Passo 2: criar um arquivo para cada loja
dicionariolojas = {}
for loja in lojas['Loja']:
    dicionariolojas[loja] = vendas.loc[vendas['Loja'] == loja, :]

#data mais recente da planilha de vendas para calcular o indicador(dia que o relatório que será enviado fará referência)
data_indicador = vendas['Data'].max()

#Passo 3: salvar a planilha na pasta de Backup
#pasta da loja já existe?
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

arquivos_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionariolojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()

    #salvar dentro da pasta
    nome_arquivo = '{}_{}_{}.xlsx'.format(data_indicador.month, data_indicador.day, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionariolojas[loja].to_excel(local_arquivo)

#Passo 4: Calcular o indicador para as lojas e Enviar e-mail para o gerente(e-mail cadastrado)
for loja in dicionariolojas:

    vendas_loja = dicionariolojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == data_indicador, :]

    # faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    # print(faturamento_ano)
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    # print(faturamento_dia)

    # diversidade de produtos
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    # print(qtde_produtos_ano)
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())
    # print(qtde_produtos_dia)

    # ticket medio
    valor_venda = vendas_loja.groupby('Código Venda').sum()
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    # print(ticket_medio_ano)
    # ticket_medio_dia
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    # print(ticket_medio_dia)

    # enviar o e-mail
    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    mail.Subject = f'OnePage Dia {data_indicador.day}/{data_indicador.month} - Loja {loja}'
    # mail.Body = 'Texto do E-mail'

    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano >= meta_qtdeprodutos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    mail.HTMLBody = f'''
    <p>Bom dia, {nome}</p>

    <p>O resultado de ontem <strong>({data_indicador.day}/{data_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
      </tr>
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou à disposição.</p>
    <p>Atenciosamente, Hemili</p>
    '''

    # Anexos:
    attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{data_indicador.month}_{data_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))

    mail.Send()
    print('E-mail da Loja {} enviado'.format(loja))