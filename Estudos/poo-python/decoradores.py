from typing import Any


def meu_decorador(func):
    def wrapper():
        print("Antes da função")
        func()
        print("Depois da função")
    return wrapper

@meu_decorador
def minha_funcao():
    print("Minha função foi chamada")

minha_funcao()

class MeuDecoradorDeClasse:
    def __init__(self, func) -> None:
        self.func = func
    def __call__(self) -> Any:
        print("Antes da classe")
        self.func()
        print("Depois da classe")

@MeuDecoradorDeClasse
def minha_classe():
    print("Minha classe foi chamada")

minha_classe()