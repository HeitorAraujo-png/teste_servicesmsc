from settings import *
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

class Relatorio:
    
    def __init__(self, arquivo1, arquivo2):
        """Adicione o relatorio de 1 a 15 dias e de 15 a 31"""
        self.caminho1 = arquivo1
        self.caminho2 = arquivo2
        self.LinkPathPTD = fr'{BASE_DIR}\{UPLOADCSV}\DPTDIA.csv'
        self.Concatena()
        self.departamento = pd.read_csv(self.LinkPathPTD, encoding='latin1')
        self.dias = self.Dias()
        self.centros = self.CentroCustos()
        self.relatorio = pd.read_csv(fr'{BASE_DIR}\{UPLOADCSV}\relatorio.csv', encoding='latin1')    
        
    def Concatena(self):
        relatorio1 = pd.read_excel(self.caminho1, parse_dates=["DATA"])
        relatorio2 = pd.read_excel(self.caminho2, parse_dates=["DATA"])
        relatorio1['DATA'] = relatorio1['DATA'].dt.strftime('%d/%m/%y')
        relatorio2['DATA'] = relatorio2['DATA'].dt.strftime('%d/%m/%y')
        self.departamento = pd.concat([relatorio1, relatorio2], ignore_index=True)
        self.departamento.to_csv(self.LinkPathPTD, encoding='latin1', index=False)


    def Add(self):
        with open(fr'{BASE_DIR}\{UPLOADCSV}\relatorio.csv', 'r', encoding='latin1') as arq:
            arq = arq.readlines()
            with open(fr'{BASE_DIR}\{UPLOADCSV}\relatorio.csv', 'w', encoding='latin1') as csv:
                for i in arq:
                    csv.write(f'{i[:i.find(',') + 1]}{','.join(self.dias)},{i[i.find('Total'):]}')

    def remove(self, lista):
        return list(dict.fromkeys(sorted(lista)))

    def Dias(self):
        return self.remove(list(self.departamento['DATA']))
    
    def TakeIndex(self, search):
        index = self.relatorio.columns.to_list().index(search)
        result = ''
        while index >= 0:
            result = chr(index % 26 + 65) + result
            index = index // 26 - 1
        return result

    def CentroCustos(self):
        centro = self.departamento['C.C']
        centro = self.remove(list(centro))
        lista = []
        for i in range(0, len(centro)):
            pre = str(centro[i])
            ic = f'{''.join((pre[0], pre[1]))}'
            cc = f'{''.join((pre[2], pre[3]))}'
            if ic == '81':
                emp = 'TVCA'
            if ic == '65':
                emp = 'Portal MT'
            if ic == '69':
                emp = 'FMCA'
            if ic == '85':
                emp = 'On Line(MT)'
            if cc == '03':
                tipo = 'ADM'
            if cc == '06':
                tipo = 'RH'
            if cc == '10':
                tipo = 'COMERCIA'
            if cc == '11':
                tipo = 'OPEC'
            if cc == '7':
                tipo = 'MKT'
            if cc == '14':
                tipo = 'PROGRAMAÇÃO'
            if cc == '15':
                tipo = 'JORNALISMO'
            if cc == '21':
                tipo = 'TECNOLOGIA'
            lista.append(f'{centro[i]} - {tipo} - {emp}')
        return lista


    def Converte(self):
        self.Add()
        coluna = 2
        atual = {'DPT': 0}
        for dia in self.dias:
            total_dia = 0
            for cc in self.centros:
                for i in range(0, 1):
                    nm = f'{''.join((cc[0], cc[1], cc[2], cc[3]))}'
                    cc = int(nm)
                qtd = len(self.departamento[(self.departamento['C.C'] == cc) & (self.departamento['DATA'] == dia)])
                total_dia += qtd
            atual[dia] = f'=SUM({self.TakeIndex(dia)}{coluna}:{self.TakeIndex(dia)}{len(self.centros) + 1})'
        atual['Valor total'] = f'=SUM({self.TakeIndex('Valor total')}{coluna}:{self.TakeIndex('Valor total')}{len(self.centros) + 1})'
        atual['Total'] = f'=SUM({self.TakeIndex('Total')}{coluna}:{self.TakeIndex('Total')}{len(self.centros) + 1})'
        for cc in self.centros:
            nova_linha = {'DPT': cc}
            for i in range(0, 1):
                nm = f'{''.join((cc[0], cc[1], cc[2], cc[3]))}'
                cc = int(nm)
            soma = 0
            for dia in self.dias:
                qtd = len(self.departamento[(self.departamento['C.C'] == cc) & (self.departamento['DATA'] == dia)])
                nova_linha[dia] = qtd
                soma += qtd
            nova_linha['%'] = f'=SUM(({self.TakeIndex('Valor total')}{coluna}/{self.TakeIndex('Valor total')}{len(self.centros) + 2}) * 100)'
            nova_linha['Total'] = f'=SUM({self.TakeIndex(self.dias[0])}{coluna}:{self.TakeIndex(self.dias[-1])}{coluna})'
            nova_linha['Valor total'] = f'={self.TakeIndex('Total')}{coluna} * 20'
            self.relatorio = pd.concat([self.relatorio, pd.DataFrame([nova_linha])], ignore_index=True)
            coluna = coluna + 1
        atual['%'] = f'=SUM({self.TakeIndex('%')}{2}:{self.TakeIndex('%')}{len(self.centros) + 1})'
        self.relatorio = pd.concat([self.relatorio, pd.DataFrame([atual])], ignore_index=True)
        print(fr'{BASE_DIR}\{UPLOADXLSX}\relatorio_atualizado.xlsx')
        self.relatorio.to_excel(fr'{BASE_DIR}\{UPLOADXLSX}\relatorio_atualizado.xlsx', index=False)
        
    def Espaco(self):
        wb = load_workbook(fr'{BASE_DIR}\{UPLOADXLSX}\relatorio_atualizado.xlsx')
        sheet = wb.sheetnames[0]
        ws = wb[sheet]
        for col in ws.columns:
            max_l = 0
            coluna = col[0].column
            coluna_letra = get_column_letter(coluna)
            for cell in col:
                try:
                    if cell.value:
                        max_l = max(max_l, len(str(cell.value)))
                except Exception as e:
                    pass
            ajuste = max_l + 2
            ws.column_dimensions[coluna_letra].width = ajuste
        wb.save(fr'{BASE_DIR}\{UPLOADXLSX}\relatorio_atualizado.xlsx')