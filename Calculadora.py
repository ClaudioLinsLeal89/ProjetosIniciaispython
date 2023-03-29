def soma():
    x = int(input("Primeiro numero: "))
    y = int(input("Segundo numero: "))
    print("Soma: ",x+y)

def subtracao():
    x = int(input("Primeiro numero: "))
    y = int(input("Segundo numero: "))
    print("Subtracao: ",x-y)

def multiplicacao():
    x = int(input("Primeiro numero: "))
    y = int(input("Segundo numero: "))
    print("Multiplicacao: ",x*y)

def divisao():
    x = int(input("Primeiro numero: "))
    y = int(input("Segundo numero: "))
    print("Divisao: ",x/y)

opcao=1

while opcao:
    print("0. Sair")
    print("1. Somar")
    print("2. Subtrair")
    print("3. Multiplicação")
    print("4. Divisão ")

    opcao = int(input("Opção: "))

    if(opcao==1):
        soma()
    if(opcao==2):
        subtracao()
    if(opcao==3):
        multiplicacao()
    if(opcao==4):
        divisao()