import os
from collections import defaultdict
from openpyxl import load_workbook


def obter_caminhos():
    print("=== CONFIGURAÇÃO ===")
    base_path = input("Digite o caminho da pasta com as empresas: ").strip()
    excel_path = input("Digite o caminho da planilha Excel: ").strip()

    return base_path, excel_path


def listar_empresas(base_path):
    return [
        pasta for pasta in os.listdir(base_path)
        if os.path.isdir(os.path.join(base_path, pasta))
    ]


def listar_arquivos_empresa(base_path, empresa):
    empresa_path = os.path.join(base_path, empresa)

    return [
        arquivo for arquivo in os.listdir(empresa_path)
        if arquivo.lower().endswith(".pdf")
    ]


def extrair_dados(nome_arquivo):
    nome_limpo = os.path.splitext(nome_arquivo)[0]
    partes = [p.strip() for p in nome_limpo.split("-")]

    cpf = partes[0]
    votos = partes[1:]

    return cpf, votos


def processar_dados(base_path):
    dados = []

    empresas = listar_empresas(base_path)

    for empresa in empresas:
        arquivos = listar_arquivos_empresa(base_path, empresa)

        for arquivo in arquivos:
            cpf, votos = extrair_dados(arquivo)

            dados.append({
                "empresa": empresa,
                "cpf": cpf,
                "votos": votos,
                "arquivo": arquivo
            })

    return dados


def normalizar_cpf(cpf):
    return str(cpf).replace(".", "").replace("-", "").strip()


def construir_mapa(dados):
    mapa = defaultdict(list)

    for item in dados:
        cpf = normalizar_cpf(item["cpf"])
        mapa[cpf].append(item)

    # coment rick: detecção de duplicados
    duplicados = {cpf: itens for cpf, itens in mapa.items() if len(itens) > 1}

    if duplicados:
        print("\nCPFs duplicados encontrados:")
        for cpf, itens in duplicados.items():
            print(f"\nCPF: {cpf}")
            for i in itens:
                print(f" - Empresa: {i['empresa']} | Arquivo: {i['arquivo']}")

    return mapa


def escrever_excel(excel_path, mapa_dados):
    wb = load_workbook(excel_path)
    ws = wb["COMITENTES"]

    linha = 2

    while True:
        cpf_planilha = ws[f"E{linha}"].value

        if cpf_planilha is None:
            break

        cpf_norm = normalizar_cpf(cpf_planilha)

        if cpf_norm in mapa_dados:
            ws[f"I{linha}"] = "ok"

            # pega o PRIMEIRO registro (critério atual)
            registro = mapa_dados[cpf_norm][0]
            votos = registro["votos"]

            col_inicio = 12  # coluna L

            for i, voto in enumerate(votos):
                ws.cell(row=linha, column=col_inicio + i).value = voto.lower()

        else:
            ws[f"I{linha}"] = "fora"

        linha += 1

    wb.save(excel_path)
    print("\n Planilha atualizada com sucesso!")


def main():
    base_path, excel_path = obter_caminhos()

    if not os.path.exists(base_path):
        print("Caminho da pasta inválido!")
        return

    if not os.path.exists(excel_path):
        print("Caminho do Excel inválido!")
        return

    dados = processar_dados(base_path)

    mapa_dados = construir_mapa(dados)

    escrever_excel(excel_path, mapa_dados)


if __name__ == "__main__":
    main()