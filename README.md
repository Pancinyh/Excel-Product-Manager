# Excel Product Manager
## Sobre o Projeto

O Excel Product Manager é uma aplicação desenvolvida em Python para manipulação e análise de planilhas Excel (.xlsx).

O sistema permite realizar operações comuns de tratamento de dados, como ordenação de produtos, cálculo de estatísticas e remoção de registros inválidos ou duplicados, gerando novas planilhas sem alterar os dados originais.

Este projeto foi desenvolvido com foco em automação de tarefas de escritório e manipulação de dados utilizando Python.

## Funcionalidades
Manipulação de Valores
Soma de todos os valores válidos da planilha
Cálculo da média dos valores dos produtos
Identificação do produto mais caro
Identificação do produto mais barato
Ordenação dos produtos do menor para o maior valor
Exportação dos resultados para uma nova planilha
Limpeza de Dados
Remoção de produtos duplicados
Remoção de produtos com nome vazio
Remoção de produtos com valores nulos
Remoção de produtos com valores inválidos (menores ou iguais a zero)
Geração de uma nova planilha contendo apenas dados válidos

## Estrutura Esperada da Planilha
ID	Nome	Valor
1	Produto A	10.50
2	Produto B	25.00
3	Produto C	15.75

A primeira linha deve conter os cabeçalhos da planilha.

## Tecnologias Utilizadas
Python 3
OpenPyXL
Rich
Instalação

Clone o repositório:

git clone https://github.com/SEU-USUARIO/SEU-REPOSITORIO.git

Acesse a pasta:

cd SEU-REPOSITORIO

Instale as dependências:

pip install openpyxl rich
Como Utilizar
Coloque o arquivo .xlsx na mesma pasta do programa.
Execute o script principal.
Informe o nome da planilha.
Escolha a operação desejada através do menu.
Arquivos Gerados
produtos_new.xlsx

Planilha contendo os dados após remoção de duplicatas e registros inválidos.

produtos_ordenados.xlsx

Planilha contendo os produtos ordenados pelo valor em ordem crescente.

## Conceitos Praticados
Manipulação de arquivos Excel
Leitura e escrita de planilhas
Estruturas de dados (listas, dicionários e conjuntos)
Funções
Tratamento de exceções
Ordenação de dados
Estatística básica
Modularização de código
Melhorias Futuras
Interface gráfica
Geração de gráficos automáticos
Filtros personalizados
Exportação para CSV
Relatórios detalhados
Sistema orientado a objetos
Testes automatizados
## Autor

### Desenvolvido por Nicolas como projeto de estudos em Python e automação de planilhas.
