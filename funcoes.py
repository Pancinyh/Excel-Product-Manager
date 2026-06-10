from rich import print
from rich.panel import Panel
import json

def leia(msg, tipo=int):
    while True:
        try:
            opcao = tipo(input(msg))
        except ValueError, TypeError:
            print('[red]Erro. Tente novamente.[/]')
        else:
            return opcao

def menu(*opcoes):
    conteudo = ''
    for i, op in enumerate(opcoes):
        conteudo += f'[blue]{i+1}[/] - [bold white]{op}[/]\n'
    tabela = Panel.fit(
        conteudo,
        title='[bold blue]OPÇÕES[/]',
    )
    print(tabela)
    while True:
        opcao = leia('Sua Opção: ')
        if len(opcoes)>= opcao > 0:
            return opcao
        print(f'[red]Erro. O número deve ser entre 1 e {len(opcoes)}.[/]')

from datetime import datetime

def exibir_movimentacoes(lista, titulo):
    if not lista.lista:
        print('[bold red]Nenhuma movimentacao encontrada.[/]')
        return
    
    print(f'{titulo.upper()}'.center(40, '-'))
    conteudo = ''
    for dados in lista.lista:
        data = dados.data
        if isinstance(data, str):
            data = datetime.strptime(data, '%d/%m/%Y').date()
        conteudo += f"[bold white]• {data.strftime('%d/%m/%Y')} | R$[bold green]{dados.valor:<10,.2f}[/] | [bold white]Motivo: [bold blue]{dados.descricao}[/]\n"
    painel = Panel.fit(
        conteudo,
        title=titulo.upper()
    )
    print(painel)