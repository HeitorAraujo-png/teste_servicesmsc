from Utilidades import *
import pandas as pd


class Xlsx:
    
    def __init__(self):
        query = Querys()
        path, xlsx = query.MakeCSV()
        colunas = query.Linhas()
        arquivo = pd.read_csv(path)
        Lista_dia = query.Lista_Dias()
        Lista_cc = query.Centro_Custos()
        # * CRIA O NOMES DO CENTRO E A QUANTIDADES
        # ! CRIA O PRINCIPAL DO ARQUIVO!

        coluna = 2
        line = {'DPT': 0}
        for cc in Lista_cc:
            centro = query.Nomeclatura(cc)
            line['DPT'] = centro
            for dia in Lista_dia:
                qtd = query.Filtro(cc=cc, dia=dia)
                line[dia] = qtd

            # ! COLUNAS DE CALCULO
            line["Total"] = (f"=SUM({query.TakeIndex(Lista_dia[0])}{coluna}:{query.TakeIndex(Lista_dia[-1])}{coluna})")
            line["Valor total"] = (f"=DOLLAR({query.TakeIndex('Total')}{coluna} * 20)")
            line["%"] = (f"={query.TakeIndex('Valor total')}{coluna}/{query.TakeIndex('Valor total')}{len(Lista_cc) + 2}")
            coluna += 1
            arquivo = pd.concat((arquivo, pd.DataFrame([line])), ignore_index=True)
        col = {'DPT': None}


        # ! LINHAS DE CALCULO
        for i in colunas:
            if i == 'Total':
                index = i
            if i == 'Valor total':
                col[i] = f'=DOLLAR(SUM(SUM({query.TakeIndex(index)}2:{query.TakeIndex(index)}{len(Lista_cc) + 1}) * 20))'
            elif i == '%':
                col[i] = f'=SUM({query.TakeIndex(i)}2:{query.TakeIndex(i)}{len(Lista_cc) + 1}) * 100'
            elif i != 'DPT':
                col[i] = f'=SUM({query.TakeIndex(i)}2:{query.TakeIndex(i)}{len(Lista_cc) + 1})'
            else:
                col['DPT'] = 'SOMAS: '
        arquivo = pd.concat((arquivo, pd.DataFrame([col])), ignore_index=True)
        arquivo.to_csv(path, index=False, encoding='latin1')
        arquivo.to_excel(xlsx, index=False)
        query.Espaco()