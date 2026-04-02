import os

def main():
    base_path = input("Digite o caminho da pasta: ")

    # listar pastas (empresas)
    for empresa in os.listdir(base_path):
        empresa_path = os.path.join(base_path, empresa)

        if os.path.isdir(empresa_path):
            print(f"\nEmpresa: {empresa}")

            # listar arquivos dentro da pasta
            for arquivo in os.listdir(empresa_path):
                nome_arquivo = os.path.splitext(arquivo)[0]

                print(f"Arquivo bruto: {nome_arquivo}")

                # separar nome e votos
                partes = [p.strip() for p in nome_arquivo.split(",")]

                nome = partes[0]
                votos = partes[1:]

                print(f"Nome: {nome}")
                print(f"Votos: {votos}")

if __name__ == "__main__":
    main()