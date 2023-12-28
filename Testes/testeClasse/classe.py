class ControleRemoto: #criando classe / self. = this.
    def __init__(self, cor, altura, profundidade, largura):
        self.cor=cor
        self.altura=altura
        self.profundidade=profundidade
        self.largura=largura


controle_remoto = ControleRemoto("preto", "10cm", "2cm", "2cm") #instanciando uma clas
print(controle_remoto.cor) #exibindo característica da classe

controle_remoto2 = ControleRemoto("azul", "12cm", "3cm", "4cm")
print(controle_remoto2.cor)

controle_remoto3 = ControleRemoto("roxo", "11cm", "5cm", "3cm")
print(controle_remoto3.cor)
