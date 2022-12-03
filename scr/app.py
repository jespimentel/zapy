#==========================================================================
# MIT License
#
# Copyright (c) 2022 Jose E S Pimentel
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#==========================================================================
#
# url = 'https://github.com/jespimentel/zapy'
# Versão 1.0.dev1
#
#==========================================================================
import os, re, sys
import requests, json
import tkinter as tk

import pandas as pd
import matplotlib.pyplot as plt

from tkinter import filedialog 
from dotenv import load_dotenv
from pathlib import Path
from email import policy
from email.parser import BytesParser

# Funções
def obtem_dados_env():
    """Lê chave-valores do arquivo .env"""
    load_dotenv()
    return os.environ['chave']

def faz_saudacao(orgao):
  """Lê arquivo .env e saúda o usuário"""
  print (f'Bem vindo ao(à) {orgao}!\n')
  print("Selecione a pasta de e-mails (arquivos 'eml') ou arquivos texto ('txt')")
  return

def seleciona_pasta():
  """Seleciona a pasta de trabalho com Tkinter"""
  root=tk.Tk()
  root.withdraw() # Esconde a janela root do Tkinter
  path = filedialog.askdirectory(title = 'Selecione a pasta de trabalho...', initialdir = '.')
  return path

def verifica_pasta (caminho_para_arquivos):
  """Cria a pasta de trabalho caso não exista"""
  nome_da_pasta = '\\arquivos_gerados' # (Para Windows)
  caminho_de_gravacao = caminho_para_arquivos + nome_da_pasta
  if not os.path.exists(caminho_de_gravacao):
      os.makedirs(caminho_para_arquivos + nome_da_pasta)
  return caminho_de_gravacao

# Incício do programa
chave = obtem_dados_env()
faz_saudacao('Promotoria de Justiça de Piracicaba')
caminho_para_arquivos = seleciona_pasta()
path = Path(caminho_para_arquivos)
eml_files = list(path.glob('*.eml'))
txt_files = list(path.glob('*.txt'))

texts = []

# Prioriza arquivos "txt" se existentes. Não sendo o caso, lê os arquivos "eml".
if len(txt_files) !=0:
  print(f'Foram encontrados {len(txt_files)} arquivos "txt" na pasta escolhida.')
  print('Arquivos encontrados:')
  for arq in txt_files:
    print(arq)
  print('Processando os arquivos texto...')
  for file in txt_files:
    with open (file, 'r', encoding='ISO-8859-1') as f:
      text = f.read()
    texts.append(text)

elif len(eml_files) !=0:
  print(f'Foram encontrados {len(eml_files)} e-mails na pasta escolhida.')
  print('Arquivos encontrados:')
  for arq in eml_files:
    print(arq)
  print('Processando os arquivos de e-mail...')
  for file in eml_files:
    with open (file, 'rb') as fp:
      msg = BytesParser(policy=policy.default).parse(fp)
    text = msg.get_body(preferencelist=('plain')).get_content()
    texts.append(text)

else:
  print('Não encontrei arquivos para o processamento.')
  print('Até breve!')
  os.system('pause')
  sys.exit()

texto = ''
for text in texts:
  texto = texto + text

# Tratamento do texto
texto = texto.replace('Message Id', 'Message_Id')
texto = texto.replace('Group Id', 'Group_Id')
texto = texto.replace('Sender Ip', 'Sender_Ip')
texto = texto.replace('Sender Port', 'Sender_Port')
texto = texto.replace('Sender Device', 'Sender_Device')
texto = texto.replace('Message Style', 'Message_Style')
texto = texto.replace('Message Size', 'Message_Size')

features = ['Timestamp', 'Message_Id', 'Sender', 'Recipients', 'Group_Id', 'Sender_Ip', 'Sender_Port', \
            'Sender_Device', 'Type', 'Message_Style', 'Message_Size']

# REGEX
# Cada conjunto de parenteses () define um grupo, gerando uma lista como resultado do "findall". 
# O uso do "?:" anula essa funcionalidade, fazendo com que cada item da lista corresponda a uma mensagem.
# RegEx podem ser testadas em https://regex101.com/
pattern = r"\sTimestamp.+(?:\n.+)+\s(?:Message_Size\s+[0-9]+)\n" # Cada mensagem
mensagens = re.findall(pattern, texto, re.MULTILINE) # Gera a lista de mensagens

# Criação da lista de dicionários (um para cada mensagem individual)
lista_msg=[]
for mensagem in mensagens:
  item = {}
  for feature in features:
    padrao = fr'({feature}\s.+)' # f-string com raw
    res = re.findall(padrao, mensagem, re.MULTILINE)
    if res !=[]:
      elementos = res[0].split(' ', maxsplit=1)
      item[elementos[0].strip()] = elementos[1].strip() 
  lista_msg.append(item)

# Criação do dataframe com as mensagens extraídas
df = pd.DataFrame(lista_msg, columns=features)
print(f'Foram encontradas {len(df)} mensagens.')

# Identificação do alvo
criterio = (df.Message_Style == 'individual')&(df.Sender_Ip.notnull())
alvo = df[criterio]['Sender'].value_counts() 
alvo = alvo.to_string().split()[0]

# Cálculos
# Mensagens invidivuais enviadas por tipo
criterio = (df.Message_Style == 'individual')
qtd_msg_env = df[criterio].groupby(['Recipients', 'Type'])['Type'].count()

# Mensagens individuais recebidas por tipo
criterio = (df.Message_Style == 'individual')
qtd_msg_receb = df[criterio].groupby(['Sender', 'Type'])['Type'].count()

# Mensagens individuais enviadas por tipo após o 'unstack' e totalização
qtd_msg_env_desempilhada = qtd_msg_env.unstack(fill_value=0)
qtd_msg_env_desempilhada = qtd_msg_env_desempilhada.drop(qtd_msg_env_desempilhada.columns[0], axis=1)
qtd_msg_env_desempilhada['total'] = qtd_msg_env_desempilhada.sum(axis=1)
qtd_msg_env_desempilhada = qtd_msg_env_desempilhada.sort_values(by='total', ascending=False)

# Mensagens individuais recebidas por tipo após o 'unstack' e totalização
qtd_msg_receb_desempilhada = qtd_msg_receb.unstack(fill_value=0)
qtd_msg_receb_desempilhada = qtd_msg_receb_desempilhada.drop(qtd_msg_receb_desempilhada.columns[0], axis=1)
qtd_msg_receb_desempilhada['total'] = qtd_msg_receb_desempilhada.sum(axis=1)
qtd_msg_receb_desempilhada = qtd_msg_receb_desempilhada.sort_values(by='total', ascending=False)

# Quantidade de mensagens em grupo
df_part_grupos = df['Group_Id'].value_counts().to_frame()

# Criação de Dataframe de participação nos grupos
criterio = df['Group_Id'].notna()
recipients_grupos = df[criterio].Recipients
cels_grupos = []
for i, cel in recipients_grupos.iteritems():
  cels = cel.split(',')
  for num in cels:
    cels_grupos.append(num)
cels_grupos_dict = {}
cels_unicos_grupos = set(cels_grupos)
for n in cels_unicos_grupos:
  cels_grupos_dict[n] = cels_grupos.count(n) 

df_grupos = pd.DataFrame.from_dict(cels_grupos_dict, orient='index', columns=['Ocorrências']).sort_values(by='Ocorrências', \
  ascending=False)

# DF com Sender_IP
df_com_ips = df[df['Sender_Ip'].notna()]
ips = df_com_ips.Sender_Ip.value_counts()
ips_lista = ips.index.to_list()

print(f'Foram encontrados {len(ips_lista)} IPs diversos.')
resposta = input('Deseja restringir a consulta à API? <s/n>')
if resposta.lower() == 's':
  cond = True
  while (cond):
    num = input ('Qtde. consultas: ')
    if num.isdigit() and int(num)>0 and int(num)<=len(ips_lista):
      cond = False
      num = int(num)
      ips_lista = ips_lista[:num]

# Consulta à API da IPAPI
# Cria a lista com as informações de IP obtidas nas requisições
# Documentação da API: https://ipapi.com/quickstart 
# Para uma versão futura: criar função para a consulta à API
operadoras = []
for ip in ips_lista:
  elemento = {}
  try:
    dados = requests.get (f'http://api.ipapi.com/api/{ip}?access_key={chave}&hostname=1')
    dados_json = json.loads(dados.content)
    elemento = {'ip': dados_json['ip'], 'hostname' : dados_json['hostname'], 'latitude': dados_json['latitude'], 
              'longitude': dados_json['longitude'],'city': dados_json['city'], 'region_name': dados_json['region_name']}
    operadoras.append(elemento)
  except:
    resposta = 'API s/ resp.'
    elemento = {'ip': ip, 'hostname' : resposta, 'latitude': resposta, 'longitude': resposta,'city': resposta, 'region_name': resposta}
    operadoras.append(elemento)

# Criação do dataframe de operadoras com os dados obtidos da API
df_operadoras = pd.DataFrame(operadoras)

# Merge de df com df_operadoras
merged = pd.merge(df, df_operadoras, how='outer', left_on = 'Sender_Ip', right_on = 'ip')
merged = merged.drop(columns='ip', axis=1)
merged = merged.set_index('Timestamp')

# Ajuste no dataframe das operadoras para a impressão
df_operadoras = df_operadoras.set_index('ip')

# Criação de planilha Excel para resumir o trabalho
path_do_arquivo = os.path.join(verifica_pasta(caminho_para_arquivos),'resumo.xlsx' )
print('Gravando a planilha...')
with pd.ExcelWriter(path_do_arquivo) as writer:
    merged.to_excel(writer, sheet_name='Geral')
    qtd_msg_env_desempilhada.to_excel(writer, sheet_name='Msg enviadas')
    qtd_msg_receb_desempilhada.to_excel(writer, sheet_name='Msg recebidas')
    df_part_grupos.to_excel(writer, sheet_name='Part. grupos')
    df_grupos.to_excel(writer, sheet_name='Msg nos grupos')
    ips.to_excel(writer, sheet_name='IPs - n. de acessos')
    df_operadoras.to_excel(writer, sheet_name='Provedores')
print('Planilha gerada.')
resp = input('Deseja gerar os gráficos?(s/n)')
if resp != 's':
  sys.exit()
else:
  # Gráfico de pizza: Tipos de mensagens enviadas pelo alvo
  conta_msg_graf = qtd_msg_env_desempilhada.drop(columns='total')
  criterio = (conta_msg_graf.index == alvo)
  conta_msg_serie = conta_msg_graf[criterio].loc[alvo]
  conta_msg_serie.plot(kind='pie', title= f'Tipos de mensagens enviadas pelo alvo {alvo}', legend=True, figsize =(10, 10), autopct='%1.0f%%')
  path_do_arquivo = os.path.join(verifica_pasta(caminho_para_arquivos),'tipos-msg-alvo.png' )
  plt.savefig(path_do_arquivo)
  plt.show()

  # Gráfico de barras: Quantidade e tipos de mensagens individuais
  criterio = (conta_msg_graf.index != alvo)
  if len(conta_msg_graf[criterio])>30:
    conta_msg_graf[criterio].head(30).plot.bar(stacked=True, title='Top 30 - qtde e tipos de msg individuais', figsize=(16,8))
  else:
    conta_msg_graf[criterio].plot.bar(stacked=True, title='Qtde e tipos de msg individuais', figsize=(16,8))
  path_do_arquivo = os.path.join(verifica_pasta(caminho_para_arquivos),'qtd-tipo-msg-ind.png')
  plt.savefig(path_do_arquivo)
  plt.show()

  # Gráfico de barras: Participação nos grupos (quantidades de mensagens)
  df['Group_Id'].value_counts().plot(kind = 'bar', title='Participação nos Grupos (qtde. de msg)', figsize=(16,8),legend=True)
  path_do_arquivo = os.path.join(verifica_pasta(caminho_para_arquivos),'part-grupos.png')
  plt.savefig(path_do_arquivo)
  plt.show()

print('Programa concluído.')
