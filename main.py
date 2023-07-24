import openpyxl
from openpyxl.styles import Font

def cadastrar_item(nome, categoria, preco, quantidade, arquivo):
    try:
        workbook = openpyxl.load_workbook(arquivo)
        sheet = workbook.active
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Nome", "Categoria", "Preço", "Quantidade"])
        # Deixar a primeira coluna em negrito
        for cell in sheet["1:1"]:
            cell.font = Font(bold=True)

    sheet.append([nome, categoria, preco, quantidade])
    workbook.save(arquivo)
    print(f"Item '{nome}' cadastrado com sucesso!")

# Nome do arquivo onde os dados serão armazenados
arquivo = "estabelecimento.xlsx"

while True:
    print("\nCadastro de Itens")
    print("1 - Cadastrar novo item")
    print("0 - Sair")
    opcao = input("Digite a opção desejada: ")

    if opcao == "1":
        nome = input("Digite o nome do item: ")
        categoria = input("Digite a categoria do item: ")
        preco = float(input("Digite o preço do item: "))
        quantidade = int(input("Digite a quantidade disponível: "))
        cadastrar_item(nome, categoria, preco, quantidade, arquivo)
    elif opcao == "0":
        print("Encerrando o programa.")
        break
    else:
        print("Opção inválida. Tente novamente.")
