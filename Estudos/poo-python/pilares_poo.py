"""_summary_"""

from abc import ABC, abstractmethod

class Animal(ABC):
    """_summary_"""

    def __init__(self, nome) -> None:
        self.__nome = nome

    @abstractmethod
    def emitir_som(self):
        """_summary_"""

class Cachorro(Animal):
    """_summary_"""

    def emitir_som(self):
        """_summary_"""
        return f"{self.__nome}: Au Au"

class Gato(Animal):
    """_summary_"""

    def emitir_som(self):
        """_summary_"""
        return f"{self.__nome}: Miau"
