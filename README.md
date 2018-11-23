# JOBAnalizer

Marcelo Farinaro, 14/12/2016
Versao Python 3.5

O Programa irá ler todos os arquivos .P10 presentes na pasta do executavel e achar o mais recente para criar um unico arquivo  excel "Resultado.xls", por fim o arquivo p10 utilizado vai será renomeado para "last.txt".

P10 é um arquivo plain text gerado gerado automaticamente pelo SAP.

A ideia eh que a primeira linha que comece com o pipe | será o cabecalho e todas as demais linhas que serao linhas.

modelo de cabecalho =  Material;Material Description;Bin;Unrestricted;SLoc;Plnt; quebra de linha
Indice =						0		1 			2		3		4	 	  5	       6   

Linhas sem Pipe serão ignoradas.
