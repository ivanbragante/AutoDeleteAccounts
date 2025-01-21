import time
from navegador import abrir_nav, logar
from planilha import rodar
#Este arquivo é responsável por chamar as funções de abrir o navegador, logar na conta e rodar a automação da analise das contas

def main():
    abrir_nav()
    time.sleep(5)
    logar()
    rodar()

if __name__ == "__main__":
    main()