import openpyxl
import math
import os

# === Configuração das lojas ===
LOJAS = {
    "02 JANAUBA": r"\\192.168.1.185\Compras\ANDERSON - JOSIANE\PEDIDO LOJA\METOD NOVA\02 JANAUBA ABASTECIMENTO PADRAO.xlsx",
    "04 MAJOR PRATES": r"\\192.168.1.185\Compras\ANDERSON - JOSIANE\PEDIDO LOJA\METOD NOVA\04 MAJOR PRATES ABASTECIMENTO PADRAO.xlsx",
    "05 SAO JOSE": r"\\192.168.1.185\Compras\ANDERSON - JOSIANE\PEDIDO LOJA\METOD NOVA\05 SAO JOSE ABASTECIMENTO PADRAO.xlsx",
    "07 SALINAS": r"\\192.168.1.185\Compras\ANDERSON - JOSIANE\PEDIDO LOJA\METOD NOVA\07 SALINAS ABASTECIMENTO PADRAO.xlsx",
    "08 PIRAPORA": r"\\192.168.1.185\Compras\ANDERSON - JOSIANE\PEDIDO LOJA\METOD NOVA\08 PIRAPORA ABASTECIMENTO PADRAO.xlsx",
}

# === Colunas fixas ===
COLUNAS = {
    "embalagem": 3,      # C
    "estoque_matriz": 4, # D
    "dias_matriz": 5,    # E
    "estoque_loja": 24,  # X
    "dias_loja": 25,     # Y
    "pedir": 26,         # Z
    "media_venda": 32,   # AF
}

# === Pergunta dias de cobertura ===
dias_cobertura_frios = int(input("Informe os dias de cobertura para FRIOS: "))
dias_cobertura_condimentos = int(input("Informe os dias de cobertura para CONDIMENTOS: "))

# === Pergunta lojas ===
print("\nSelecione as lojas para analisar (separe por vírgula):")
for i, loja in enumerate(LOJAS.keys(), start=1):
    print(f"{i} - {loja}")

opcoes = input("Opção(s): ").split(",")
selecionadas = [list(LOJAS.keys())[int(i.strip()) - 1] for i in opcoes]

for loja in selecionadas:
    caminho = LOJAS[loja]

    if not os.path.exists(caminho):
        print(f"❌ Arquivo não encontrado: {caminho}")
        continue

    print(f"\n📂 Processando {loja}...")

    # === Abrir planilha com valores congelados das fórmulas ===
    wb = openpyxl.load_workbook(caminho, data_only=True)

    # Criar/limpar aba de relatório
    if "RELATORIO" in wb.sheetnames:
        ws_rel = wb["RELATORIO"]
        wb.remove(ws_rel)
    ws_rel = wb.create_sheet("RELATORIO")
    ws_rel.append(["ABA", "CODIGO", "DESCRICAO", "PEDIR"])

    # === Processar abas Frios e Condimentos ===
    for aba in ["FRIOS", "CONDIMENTOS"]:
        if aba not in wb.sheetnames:
            print(f"⚠ Aba {aba} não encontrada, pulando...")
            continue

        ws = wb[aba]
        dias_cobertura = dias_cobertura_frios if aba == "FRIOS" else dias_cobertura_condimentos

        for row in range(3, ws.max_row + 1):  # começa na linha 3
            try:
                embalagem = ws.cell(row=row, column=COLUNAS["embalagem"]).value or 0
                estoque_matriz = ws.cell(row=row, column=COLUNAS["estoque_matriz"]).value or 0
                dias_matriz = ws.cell(row=row, column=COLUNAS["dias_matriz"]).value or 0
                estoque_loja = ws.cell(row=row, column=COLUNAS["estoque_loja"]).value or 0
                dias_loja = ws.cell(row=row, column=COLUNAS["dias_loja"]).value or 0
                media_venda = ws.cell(row=row, column=COLUNAS["media_venda"]).value or 0

                # Garantir numéricos
                try:
                    embalagem = int(embalagem)
                except:
                    embalagem = 1
                try:
                    estoque_matriz = float(estoque_matriz)
                except:
                    estoque_matriz = 0
                try:
                    dias_matriz = float(dias_matriz)
                except:
                    dias_matriz = 0
                try:
                    estoque_loja = float(estoque_loja)
                except:
                    estoque_loja = 0
                try:
                    dias_loja = float(dias_loja)
                except:
                    dias_loja = 0
                try:
                    media_venda = float(media_venda)
                except:
                    media_venda = 0

                pedir_final = 0

                # === Regra 1: estoque da matriz insuficiente ===
                if estoque_matriz <= 0 or dias_matriz < 10:
                    pedir_final = 0
                else:
                    # === Regra 2: pedido normal ===
                    demanda_dia = media_venda / 30
                    estoque_ideal = demanda_dia * dias_cobertura
                    pedir_bruto = estoque_ideal - estoque_loja

                    if pedir_bruto > 0:
                        pedir_final = math.ceil(pedir_bruto / embalagem) * embalagem

                # === Regra 3: exceções ===
                if estoque_loja <= 0 or dias_loja < 15:
                    pedir_final = max(1, embalagem)

                # Grava no campo "PEDIR"
                ws.cell(row=row, column=COLUNAS["pedir"], value=int(pedir_final))

                # Adiciona ao relatório
                codigo = ws.cell(row=row, column=1).value
                descricao = ws.cell(row=row, column=2).value
                ws_rel.append([aba, codigo, descricao, pedir_final])

            except Exception as e:
                print(f"⚠ Erro na linha {row} da aba {aba}: {e}")

    # Salvar
    wb.save(caminho)
    print(f"✅ Arquivo atualizado: {caminho}")
