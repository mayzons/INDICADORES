import os
import pandas as pd

# Caminho da pasta onde está o consolidado
pasta_base = r"C:\proj_pessoal\INDICADORES"
# Nome do arquivo consolidado
arquivo_entrada = os.path.join(pasta_base, "consolidado.xlsx")

# Lê o arquivo consolidado
df = pd.read_excel(arquivo_entrada)

# Lista de operações possíveis
operacoes = {
    "1": "Soma",
    "2": "Média",
    "3": "Mínimo",
    "4": "Máximo",
    "5": "Contagem",
    "6": "Top 5"
}


def formatar_valor(valor):
    # Formata números e percentuais
    if isinstance(valor, (int, float)):
        if 0 <= valor <= 1:  # interpreta como percentual
            return f"{valor:.2%}"
        else:
            return f"{valor:,.2f}"
    return valor


def executar_operacao():
    global df
    print("\n📊 Colunas disponíveis no arquivo:")
    for i, col in enumerate(df.columns, 1):
        print(f"{i}. {col}")

    escolha_coluna = input("\nDigite o número da coluna que deseja analisar: ")
    try:
        coluna = df.columns[int(escolha_coluna) - 1]
    except (ValueError, IndexError):
        print("❌ Opção inválida!")
        return True

    print("\n🔧 Operações disponíveis:")
    for key, val in operacoes.items():
        print(f"{key}. {val}")

    escolha_operacao = input("\nEscolha a operação desejada: ")

    try:
        if escolha_operacao == "1":  # Soma
            if pd.api.types.is_numeric_dtype(df[coluna]):
                resultado = df[coluna].sum()
            else:
                print("❌ Não é possível somar uma coluna não numérica.")
                return True

        elif escolha_operacao == "2":  # Média
            if pd.api.types.is_numeric_dtype(df[coluna]):
                resultado = df[coluna].mean()
            else:
                print("❌ Não é possível calcular média de uma coluna não numérica.")  # noqa
                return True

        elif escolha_operacao == "3":  # Mínimo
            resultado = df[coluna].min()

        elif escolha_operacao == "4":  # Máximo
            resultado = df[coluna].max()

        elif escolha_operacao == "5":  # Contagem
            resultado = df[coluna].count()

        elif escolha_operacao == "6":  # Top 5
            if pd.api.types.is_numeric_dtype(df[coluna]):
                resultado = df.nlargest(5, coluna)
            else:
                print("❌ Top 5 só funciona com colunas numéricas.")
                return True

        else:
            print("❌ Operação inválida!")
            return True

        # Exibir o resultado
        print("\n✅ Resultado da operação:")
        if isinstance(resultado, pd.DataFrame):
            print(resultado)
        else:
            print(formatar_valor(resultado))

        # Menu de próxima ação
        print("\nO que deseja fazer agora?")
        print("1. Exportar resultado")
        print("2. Fazer outra operação")
        print("3. Limpar tela")
        print("4. Encerrar")

        opcao = input("Escolha: ")
        if opcao == "1":
            nome_saida = input(
                "Digite o nome do arquivo de saída (sem extensão): ")
            caminho_saida = f"{nome_saida}.xlsx"
            if isinstance(resultado, pd.DataFrame):
                resultado.to_excel(caminho_saida, index=False)
            else:
                pd.DataFrame(
                    [{"Coluna": coluna,
                        "Resultado": resultado}]).to_excel(caminho_saida,
                                                           index=False)
            print(f"📂 Resultado exportado para {caminho_saida}")
            return True
        elif opcao == "2":
            return True
        elif opcao == "3":
            os.system('cls' if os.name == 'nt' else 'clear')
            return True
        else:
            return False

    except Exception as e:
        print(f"❌ Erro: {e}")
        return True


# Loop principal
continuar = True
while continuar:
    continuar = executar_operacao()
