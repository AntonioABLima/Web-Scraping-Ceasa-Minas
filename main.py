from bs4 import BeautifulSoup # Lib para pegar dados de um codigo html
import requests # Lib para fazer o request de um link
from openpyxl import Workbook # Lib para fazer o a planilha em excel 
from openpyxl.styles import Alignment

def main():
    # Recebendo o html da site ceasa minas
    url = 'http://minas1.ceasa.mg.gov.br/ceasainternet/cst_precosmaiscomumMG/cst_precosmaiscomumMG.php'
    html_text = requests.get(url).text

    soup = BeautifulSoup(html_text, 'lxml')

    # 48 Alimentos no total, 24 tem são da class scGridFieldOdd e outros 24 scGridFieldEven
    classeImpar = 'scGridFieldOdd'
    foodsOdd = soup.find_all('tr', class_ = classeImpar)
    
    classePar = 'scGridFieldEven'
    foodsEven = soup.find_all('tr', class_ = classePar)

    # Separando os dados de cada alimento um matrizes, cada linha um alimento novo
    produtosImpares = getData(foodsOdd, 'scGridFieldOddFont')
    produtosPares = getData(foodsEven, 'scGridFieldEvenFont')
    
    # Unindo os as duas matrizes
    alimentos = []
    for i in range(len(produtosImpares)):
        alimentos.append(produtosImpares[i])
        alimentos.append(produtosPares[i])
        
    # A lista alimentos possui todos os alimentos encontrados no site
    criandoTabela(alimentos)


def getData(soup ,classe):
    # Nos interessa o nome (primeira coluna), embalagens ou unidade de peso (segunda coluna) e o preço em uberlandia (quarta coluna)
    # Classes:
    # Nome: ... css_produto_grid_line
    # Embalagens: ... css_unidade_grid_line
    # Preço Uberlandia: ... css_precosub_grid_line
    lista = []
    for food in soup:
        temp = []
        nome = food.find('td', class_ = classe + " css_produto_grid_line").text.replace(' ', '')
        embalagem = food.find('td', class_ = classe + " css_unidade_grid_line").text.replace(' ', '')
        precoUdi = food.find('td', class_ = classe + " css_precosub_grid_line").text.replace(' ', '')

        temp.append(nome)
        temp.append(embalagem)
        temp.append(precoUdi)
        lista.append(temp)

    return lista

def criandoTabela(lista):
    wb = Workbook()
    planilha = wb.worksheets[0]

    alignment = Alignment(horizontal='center')
    
    planilha.column_dimensions['B'].width = 20
    planilha.column_dimensions['C'].width = 20
    planilha.column_dimensions['D'].width = 20

    # Criando e formatando cabeçalhos
    planilha['B1'] = 'Produtos'
    planilha['B1'].alignment = alignment

    planilha['C1'] = 'Embalagens'
    planilha['C1'].alignment = alignment

    planilha['D1'] = 'Uberlândia'
    planilha['D1'].alignment = alignment

    # Inserindo produtos contidos na lista
    for produto in lista:
        planilha['B' + str(lista.index(produto) + 2)] = produto[0]
        planilha['B' + str(lista.index(produto) + 2)].alignment = alignment

        planilha['C' + str(lista.index(produto) + 2)] = produto[1]
        planilha['C' + str(lista.index(produto) + 2)].alignment = alignment

        planilha['D' + str(lista.index(produto) + 2)] = produto[2]
        planilha['D' + str(lista.index(produto) + 2)].alignment = alignment

    # Salvando planilha com o nome de aliementos
    wb.save("alimentos.xlsx")

if __name__ == "__main__":
    main()
