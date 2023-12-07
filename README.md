# Projeto automatizacao-de-planilha-viagens

## A empresa que solicitou este projeto tinha uma rotina para a solicitação de viagens para motoristas com os seguintes passos:
1 - Registrar as viagens informando:
- setor solicitante
- tipo de veículo,
- placa do veículo,
- roteiro da viagem,
- nome dos passageiros,
- dara início da viagem,
- data fim da viagem,
- quantidade de diárias a serem pagas,
- nome do motorista e
- justificativa para o deslocamento.
##
2 - Gerar um número de pedido de viagem
##
3 - Somar as diárias de cada motorista para controle
##
4 - Criar um documento de texto para informar à empresa que paga aos motoristas a lista de diarias a serem pagas para cada um deles
##
5 - Transformar este documento de texto em um documento .pdf
##
6 - Enviar este documeto .pdf via e-mail, para a empresa que paga as diárias aos motoristas.

### Estes passos resultavam em um trabalho operacional enorme no dia a dia de trabalho desta empresa.

## Como solução criei uma planilha Google Sheet com as abas:

## - Banco de Dados, onde temos todas as informações possíveis para os campos:
### motorista, 
### setor solicitante, 
### tipo de veículo, 
### placa do veículo e 
### justificativa da viagem.
## - Programação de viagen, onde nas linhas da planilha obteos nas colunas respectivas as informações que estão na aba Banco de Dados, e complementamos com data início e fim da viagem. A planilha calcula
automaticamente a quantidade de diárias. No final da inclusão de todas as diárias do momento, o usuário registra o caractere "X" ou "x" na coluna Solicitar Diária e clica no botão Azul no cabeçalho da aba. O script JS vai somar a quantidade de diárias para cada motodista, e atualizar na aba Motoristas x Diárias, gerar registro na aba Solicita Pagamento de Diária. Esta última aba tem um script do AUTOCRAT que dispara, automaticamente, a geração de documento .pdf contendo todas as informações do pedido de diária que deverá ser enviado à empresa que paga as diárias dos motoristas. Ao final, este script envia este documento .pdf por email.

## Esta automação da rotina inicial otimizou muito o trabalho da empresa, que liberou tempo para que o usuário fizesse novas atividades.

## As telas das abas das planilhas podem ser vistas nos arquivos .pdf anexos.

## Esta aplicação roda quando a planilha no Google Sheets é executada.



  

