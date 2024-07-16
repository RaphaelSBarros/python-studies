class Personagem:
    def __init__(self, nome, vida, nivel):
        self.__nome = nome
        self.__vida = vida
        self.__nivel = nivel
    
    def get_nome(self):
        return self.__nome
    
    def get_vida(self):
        return self.__vida
    
    def get_nivel(self):
        return self.__nivel
    
    def exibir_detalhes(self):
        return f"Nome: {self.get_nome()}\nVida: {self.get_vida()}\nNível: {self.get_nivel()}"
    
    def receber_ataque(self, dano):
        self.__vida -= dano
        if self.__vida <= 0:
            self.__vida = 0

    def atacar(self, alvo):
        dano = self.__nivel * 2
        alvo.receber_ataque(dano)
        print(f"{self.__nome} Atacou {alvo.get_nome()} e causou {dano} de dano!\n")

class Heroi(Personagem):
    def __init__(self, nome, vida, nivel, habilidade):
        super().__init__(nome, vida, nivel)
        self.__habilidade = habilidade
    
    def get_habilidade(self):
        return self.__habilidade
    
    def exibir_detalhes(self):
        return f"{super().exibir_detalhes()}\nHabilidade: {self.get_habilidade()}\n"
    
    def ataque_especial(self, alvo):
        dano = self.get_nivel() * 5
        alvo.receber_ataque(dano)
        print(f"{self.get_nome()} usou a habilidade especial {self.get_habilidade()} em {alvo.get_nome()} e causou {dano} de dano!\n")
    
class Inimigo(Personagem):
    def __init__(self, nome, vida, nivel, tipo):
        super().__init__(nome, vida, nivel)
        self.__tipo = tipo
    
    def get_tipo(self):
        return self.__tipo
    
    def exibir_detalhes(self):
        return f"{super().exibir_detalhes()}\nHabilidade: {self.get_tipo()}\n"
    
class Jogo:
    def __init__(self):
        self.heroi = Heroi("Heroi", 100, 5, "Super-Força")
        self.inimigo = Inimigo("Morcego", 50, 3, "Voador")

    def iniciar_batalha(self):
        print("Iniciando Batalha!")
        while self.heroi.get_vida() > 0 and self.inimigo.get_vida() > 0:
            print("Detalhes dos personagens:")
            print(self.heroi.exibir_detalhes())
            print(self.inimigo.exibir_detalhes())

            input("Pressione Enter para atacar ...")
            escolha = input("Escolha\n[1] Ataque Normal  [2] Ataque Especial\n")

            if escolha == "1":
                self.heroi.atacar(self.inimigo)
            elif escolha == "2":
                self.heroi.ataque_especial(self.inimigo)
            else:
                print("Escolha inválida")

        if self.heroi.get_vida() > 0:
            print("Parabéns! Você venceu.")
        else:
            print("Você foi derrotado!")

jogo = Jogo()
jogo.iniciar_batalha()