import sys
from pathlib import Path
file = Path(__file__).resolve()
parent, root = file.parent, file.parents[1]
sys.path.append(str(root))

from controllers.pessoa import PessoaController
from models.pessoa import Pessoa

decisao = 1

while decisao != 0:
    decisao = int(input('digite 1 para salvar, 2 para listar e 0 para sair'))
    if decisao == 1:
        nome = input('Nome: ')
        sobrenome = input('Sobrenome: ')
        idade = input('Idade: ')
        cpf = input('CPF: ')
        p1 = Pessoa(nome=nome, sobrenome=sobrenome, idade=idade, cpf=cpf)
        PessoaController.salvar_pessoa(p1)
    elif decisao == 2:
        for pessoa in PessoaController.listar_pessoas():
            print(pessoa.nome)