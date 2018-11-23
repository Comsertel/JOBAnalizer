# -*- coding: utf-8 -*-
"""
Marcelo Farinaro, 14/12/2016
Versao Python 3.5

O Programa vai ler todos os arquivos .P10 presentes nesta pasta e achar o mais recente para ira criar um unico arquivo Resultado.xls, o arquivo p10 utilizado vai ser renomeado para last.txt
p10 eh gerado automaticamente pelo SAP.

A ideia eh que a primeira linha que comece com o pipe | sera o cabecalho e todas as demais linhas que serao linhas


modelo de cabecalho =  Material;Material Description;Bin;Unrestricted;SLoc;Plnt; quebra de linha
Indice =						0		1 			2		3		4	 	  5	       6   

"""

import os, re, openpyxl

def findFile():
    """
    Procura na pasta do script todos os arquivos .P10, e retorna apenas o mais recente.
    """
    filePattern = re.compile(r"(.*)\.[pP]10") #regex
    arquivosPasta = os.listdir()
    arquivosP10 = []
    
    for arquivo in arquivosPasta:
        if filePattern.search(arquivo):
            arquivosP10.append(arquivo)
        
    return(max(arquivosP10, key = os.path.getctime))#localiza o arquivo com a data de criacao mais recente e retorna

def lineCleaner(line):
    """
    Recebe uma String, o delimitador eh o | e retorna uma lista
    """
    splitLine = line.split("|")
    
    splitLine.pop(0) #limpa primeiro item inutil
    splitLine.pop(len(splitLine)-1) #limpa ultimo item inutil
    
    for i in range(len(splitLine)):
        splitLine[i] = splitLine[i].strip()             
    
    return splitLine

def isNumber(word):
    """
    Confere se o valor do campo eh numerico
    """
    filePattern = re.compile(r"([0-9]+),([0-9]+)")
    
    result = filePattern.search(word)
    
    if filePattern.search(word):
        return int(result.group(1))
    else:
        return word
        
            

row = 1 #recebe a linha

header = []

fileName = findFile()
fhand = open(fileName,encoding='utf8') #abre arquivo P10

xlHand = openpyxl.Workbook() #abre o arquivo
xlHand.create_sheet(index = 0, title ="Main") #cria aba
shHand = xlHand.get_sheet_by_name("Main") #abre a aba
                             

for line in fhand: #para cada linha
    if line.startswith("|"):
        
        words = lineCleaner(line)
        if header != [] and len(words) != len(header): #trata o que for diferente do padrao do header
            words = header
        
        if words != header: #grava o que for diferente do header
            col = 1 
            for word in words:
                cellHand =  shHand.cell(row = row, column = col)
                cellHand.value = isNumber(word)
                col += 1
                
            if header == []:#grava o header
                header = lineCleaner(line)
            row += 1
   

fhand.close()

try:
    os.remove("last.txt")
except:
    print("Nao encontrado")

os.rename(fileName, "last.txt")    

xlHand.save("Resultado.xlsx") #salva arquivo 
