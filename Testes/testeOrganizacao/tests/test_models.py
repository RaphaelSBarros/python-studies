import sys
sys.path.append(".")
from src.models import pessoa

def test_concatenacao_nome_sobrenome(): #é pro nome ser gigantrolho mesmo. Tem que ser descritivo
    p1 = pessoa.Pessoa('Caio', 'Sampaio', 22, "12312312321")
    assert p1.nome_completo() == 'Caio Sampaio'