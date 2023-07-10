import pandas as pd

base_involves = pd.read_excel("bases//base_involves_PE3C.xlsx")

base_trade = pd.read_excel("bases//base_trade_PE3C.xlsx")

base_involves = base_involves[['PDV', 'Notificante', 'Data e hora da pesquisa', 'QUAL A QUANTIDADE DE PONTO EXTRA 3 CORÇÕES ?', 'QUAL A QUANTIDADE DE PONTO EXTRA KIMIMO', 'QUAL A QUANTIDADE DE PONTO EXTRA SANTA CLARA ?']]

base_involves['UF'] = base_involves['PDV'].str[:2]

base_involves = base_involves[['PDV', 'UF', 'Notificante', 'Data e hora da pesquisa', 'QUAL A QUANTIDADE DE PONTO EXTRA 3 CORÇÕES ?', 'QUAL A QUANTIDADE DE PONTO EXTRA KIMIMO', 'QUAL A QUANTIDADE DE PONTO EXTRA SANTA CLARA ?']]

# ------------------------------------------------------ Separando DATA da hora e formatando-a
base_involves['Data e hora da pesquisa'] = base_involves['Data e hora da pesquisa'].astype(str)

base_involves[['DATA', 'hora']] = base_involves['Data e hora da pesquisa'].str.split(" ", expand=True)

base_involves['DATA'] = pd.to_datetime(base_involves['DATA'], format='%Y/%m/%d')

base_involves = base_involves[['PDV', 'Notificante', 'UF', 'DATA', 'QUAL A QUANTIDADE DE PONTO EXTRA 3 CORÇÕES ?', 'QUAL A QUANTIDADE DE PONTO EXTRA KIMIMO', 'QUAL A QUANTIDADE DE PONTO EXTRA SANTA CLARA ?']]

base_involves[['QUAL A QUANTIDADE DE PONTO EXTRA 3 CORÇÕES ?', 'QUAL A QUANTIDADE DE PONTO EXTRA KIMIMO', 'QUAL A QUANTIDADE DE PONTO EXTRA SANTA CLARA ?']] = base_involves[['QUAL A QUANTIDADE DE PONTO EXTRA 3 CORÇÕES ?', 'QUAL A QUANTIDADE DE PONTO EXTRA KIMIMO', 'QUAL A QUANTIDADE DE PONTO EXTRA SANTA CLARA ?']].fillna(int(0)).astype(int)

base_involves = base_involves.rename({'Notificante':'PROMOTOR', 'QUAL A QUANTIDADE DE PONTO EXTRA 3 CORÇÕES ?':'QUANTOS PONTOS EXTRAS 3 CORAÇÕES', 
                                'QUAL A QUANTIDADE DE PONTO EXTRA KIMIMO':'QUANTOS PONTOS EXTRAS KIMIMO', 'QUAL A QUANTIDADE DE PONTO EXTRA SANTA CLARA ?':'QUANTOS PONTOS EXTRAS SANTA CLARA'}, axis=1
                                )

base_trade = base_trade[['PDV', 'Colaborador', 'Estado', 'Data do Preenchimento', 'Pergunta', 'Resposta Numérica']]

base_trade = base_trade.loc[base_trade['Pergunta']!="Informe quantos ponto extra do KIMIMO tem na sua loja"]

base_trade = base_trade.loc[base_trade['Pergunta']!="TEM PONTO EXTRA SANTA CLARA?"]

base_trade = base_trade.loc[base_trade['Pergunta']!="TEM PONTO EXTRA TRÊS CORAÇÕES (MELITTA)"]

base_trade = base_trade.reset_index(drop=True)

base_trade = base_trade.pivot(index=['PDV', 'Colaborador', 'Estado', 'Data do Preenchimento'], columns="Pergunta", values="Resposta Numérica")

base_trade = base_trade.reset_index(drop=False)

base_trade['Colaborador'] = base_trade['Colaborador'].str.upper()

base_trade[['QUAL A QUANTIDADE DE PONTO EXTRA 3 CORÇÕES?', 'QUAL A QUANTIDADE DE PONTO EXTRA KIMIMO', 'QUAL A QUANTIDADE DE PONTO EXTRA SANTA CLARA?']] = base_trade[['QUAL A QUANTIDADE DE PONTO EXTRA 3 CORÇÕES?', 'QUAL A QUANTIDADE DE PONTO EXTRA KIMIMO', 'QUAL A QUANTIDADE DE PONTO EXTRA SANTA CLARA?']].fillna(int(0)).astype(int)

base_trade = base_trade.rename({'Colaborador':'PROMOTOR', 'Estado':'UF', 'Data do Preenchimento':'DATA', 'QUAL A QUANTIDADE DE PONTO EXTRA 3 CORÇÕES?':'QUANTOS PONTOS EXTRAS 3 CORAÇÕES', 
                                'QUAL A QUANTIDADE DE PONTO EXTRA KIMIMO':'QUANTOS PONTOS EXTRAS KIMIMO', 'QUAL A QUANTIDADE DE PONTO EXTRA SANTA CLARA?':'QUANTOS PONTOS EXTRAS SANTA CLARA'}, axis=1
                                )

RPE3C = pd.merge(base_trade, base_involves, how='outer')

RPE3C.to_excel('Relatório de pontos extras 3 corações.xlsx', sheet_name='Relat_PE3C', index=False)