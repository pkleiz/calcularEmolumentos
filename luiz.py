import pylightxl as xl

def Louis():
    # Variaveis
    actualRow = 3 #demarca a linha atual da planilha output

    # Pegando a cotação do dolar bmf clean coupon
    dolBMF = actualDolarCotation()

    # Lendo a planilha input
    db = xl.readxl('operacoes.xlsx')

    #criando a planilha output
    output = createNewSheet()

    # Escrevendo a cotação atual do dolar na nova planilha
    output.ws(ws='Main').update_index(row=1, col=1, val='Dolar BMF')
    output.ws(ws='Main').update_index(row=1, col=2, val=dolBMF)

    # Criando os cabeçalhos da planilha output
    output.ws(ws='Main').update_index(row=2, col=1, val='Dt. Op')
    output.ws(ws='Main').update_index(row=2, col=2, val='Dt. Liq')
    output.ws(ws='Main').update_index(row=2, col=3, val='Fundo')
    output.ws(ws='Main').update_index(row=2, col=4, val='Valor Calculado em R$')


    # Varrendo a coluna do ativo e inserindo os dados na planilha de output
    for i in range (len(db.ws(ws='Main').col(col=10))):
        if ((db.ws('Main').index(row=i+1,col=10)) == 'DOLJ21'):
            output.ws(ws='Main').update_index(row=actualRow, col=6, val = db.ws('Main').index(row=i+1,col=10))
            output.ws(ws='Main').update_index(row=actualRow, col=1, val = db.ws('Main').index(row=i+1,col=5))
            output.ws(ws='Main').update_index(row=actualRow, col=2, val = db.ws('Main').index(row=i+1,col=6))
            output.ws(ws='Main').update_index(row=actualRow, col=3, val = db.ws('Main').index(row=i+1,col=7))
            output.ws(ws='Main').update_index(row=actualRow, col=4, val = calcEmoluments('DOLJ21',db.ws('Main').index(row=i+1,col=14)) * dolBMF ) #realiza o cálculo
            actualRow +=1
        if ((db.ws('Main').index(row=i+1,col=10)) == 'WDOJ21'):
            output.ws(ws='Main').update_index(row=actualRow, col=7, val = db.ws('Main').index(row=i+1,col=10))
            output.ws(ws='Main').update_index(row=actualRow, col=1, val = db.ws('Main').index(row=i+1,col=5))
            output.ws(ws='Main').update_index(row=actualRow, col=2, val = db.ws('Main').index(row=i+1,col=6))
            output.ws(ws='Main').update_index(row=actualRow, col=3, val = db.ws('Main').index(row=i+1,col=7))
            output.ws(ws='Main').update_index(row=actualRow, col=4, val = calcEmoluments('WDOJ21',db.ws('Main').index(row=i+1,col=14)) * dolBMF ) #realiza o cálculo
            actualRow +=1
            
    #escrevendo os dados na planilha output
    xl.writexl(db=output, fn="output.xlsx")
                                              
def actualDolarCotation():
    return 5.6264

def createNewSheet():
    # Criando uma planilha vazia para o output
    output = xl.Database()
    output.add_ws(ws="Main")
    return output

# Calculo de emulumentos de acordo com a BM&F - acesso em: 25/03/2021
def calcEmoluments(typeOfActive, quantity):
    if(typeOfActive == "DOLJ21"):
        if(quantity >= 1 and quantity <= 10):
            valueUSS = 0.53 * quantity
        elif(quantity >= 11 and quantity <= 150):
            valueUSS = 0.50 * quantity
        elif(quantity >= 151 and quantity <= 360):
            valueUSS = 0.45 * quantity
        elif(quantity >= 361 and quantity <= 1500):
            valueUSS = 0.42 * quantity
        elif(quantity >= 1501 and quantity <= 12500):
            valueUSS = 0.39 * quantity
        elif(quantity > 12500):
            valueUSS = 0.34 * quantity
        else:
            # erro! Quantidade negativa
            valueUSS = 0
            
    if(typeOfActive == "WDOJ21"):
        if(quantity >= 1 and quantity <= 10):
            valueUSS = 0.53*0.22 * quantity
        elif(quantity >= 11 and quantity <= 150):
            valueUSS = 0.50*0.22 * quantity
        elif(quantity >= 151 and quantity <= 360):
            valueUSS = 0.45*0.22 * quantity
        elif(quantity >= 361 and quantity <= 1500):
            valueUSS = 0.42*0.22 * quantity
        elif(quantity >= 1501 and quantity <= 12500):
            valueUSS = 0.39*0.22 * quantity
        elif(quantity > 12500):
            valueUSS = 0.34*0.22 * quantity
        else:
            # erro! Quantidade negativa
            valueUSS = 0
    return valueUSS

Louis()
