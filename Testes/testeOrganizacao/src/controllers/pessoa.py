import sys
from pathlib import Path
file = Path(__file__).resolve()
parent, root = file.parent, file.parents[1]
sys.path.append(str(root))

from models.pessoa import Pessoa
from typing import List

class PessoaController:
    pessoa = []
    
    @classmethod
    def salvar_pessoa(cls, pessoa: Pessoa):
        cls.pessoa.append(pessoa) #insere a Pessoa criada na lista de pessoas
    
    @classmethod
    def listar_pessoas(cls) -> List[Pessoa]: #Retorna uma lista de pessoas
        return cls.pessoa