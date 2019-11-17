# Gerador de Rotinas para Detecção de Amomalias

Programa gerador de simulação de rotinas para validação de algoritmos de detecção de anomalias.

Para executar o programa, utilizar a seguinte linha de comando em ambiente Windows:

        GeraRotina.exe numDias NomeArquivo numAlgoritmo [percAnomalias]
          numDia:       quantidade de dias a serem simulados.
          nomeArquivo:  nome do arquivo a ser gerado (o arquivo será gerado na pasta output, no diretório onde a 
                        aplicação está sendo executada).
          umAlgoritmo:  id do algoritmo a ser utilizado para pré-processamento dos dados (1 = DBSCAN; 2 = LOF; 
                        3 = OPTICS).
          percAnomlais: porcentagem de anomalias a serem geradas.(opcional)
          
        exemplo: GeraRotina.exe 365 Dataset_01 1 5
