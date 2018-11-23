# JOBAnalizer

Marcelo Farinaro, 14/12/2016
Versao Python 3.5

O Programa vai ler todos os arquivos .P10 presentes na pasta e achar o mais recente para ira criar um unico arquivo Resultado.xls, o arquivo p10 utilizado vai ser renomeado para last.txt
p10 eh gerado automaticamente pelo SAP.

A ideia eh que a primeira linha que comece com o pipe | sera o cabecalho e todas as demais linhas que serao linhas


modelo de cabecalho =  Material;Material Description;Bin;Unrestricted;SLoc;Plnt; quebra de linha
Indice =						0		1 			2		3		4	 	  5	       6   
