try: # operação
    a = int(input("Numerador: "))
    b = int(input("Denominador: "))
    r = a / b
except (ZeroDivisionError):
    print("Não é possíve dividir um número por zero")
except (ValueError, TypeError):
    print("O valor não está no formato correto")
else: # caso nao dê erro exibe o resultado
    print(f'O resultado é: {r:.1f}')
finally: # exibe independente de erro ou não
    print("Volte sempre!")