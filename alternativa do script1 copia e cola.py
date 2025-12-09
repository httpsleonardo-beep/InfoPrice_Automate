import pandas as pd
import pyperclip  # Biblioteca que controla o Ctrl+C / Ctrl+V
import os

# === CONFIGURAÇÃO ===
ARQUIVO_CURVA_A = "C:\\Users\\Leonardo.Galdino\\Desktop\\scripts\\InfoPrice\\Curva A Compradores 01-01 á 25-11.xlsx"
# ====================

print(">>> Lendo sua planilha e gerando lista de busca (SEM VÍRGULAS)...")


def carregar_planilha_robusta(arquivo):
    # Tenta ler de todas as formas possíveis
    try:
        return pd.read_csv(arquivo, sep=None, engine='python', encoding='latin1')
    except:
        pass
    try:
        return pd.read_csv(arquivo, sep=',', engine='python', encoding='utf-8')
    except:
        pass
    try:
        return pd.read_csv(arquivo, sep=';', engine='python', encoding='latin1')
    except:
        pass
    try:
        return pd.read_excel(arquivo)
    except:
        pass
    return None


try:
    # 1. Carrega o arquivo
    df = carregar_planilha_robusta(ARQUIVO_CURVA_A)

    if df is None:
        print("❌ ERRO: Não consegui ler o arquivo 'Curva A'. Verifique se ele está na pasta.")
        input("Pressione Enter para sair...")
        exit()

    # 2. Identifica a coluna
    if 'Código Acesso' in df.columns:
        coluna = 'Código Acesso'
    else:
        # Pega a 5ª coluna se não achar o nome
        coluna = df.columns[4]

        # 3. Limpa os dados (Tira .0, tira espaços, tira letras)
    lista_eans = df[coluna].dropna().astype(str).str.replace(r'\.0$', '', regex=True).str.strip().unique()

    # Filtra só o que tem cara de código de barras
    lista_limpa = [x for x in lista_eans if len(x) > 6 and x.isdigit()]

    # 4. AQUI ESTÁ A MUDANÇA: Junta tudo com ESPAÇO, sem vírgula
    texto_final = " ".join(lista_limpa)

    # 5. Mágica: Copia para o seu Clipboard
    pyperclip.copy(texto_final)

    print("\n" + "=" * 50)
    print(f"✅ SUCESSO! {len(lista_limpa)} CÓDIGOS COPIADOS!")
    print("   Formato: '789... 789... 789...' (Espaços)")
    print("=" * 50)
    print("O que fazer agora:")
    print("1. Vá no InfoPrice > Filtro de Produtos.")
    print("2. Clique na busca e aperte CTRL+V.")
    print("3. O site deve reconhecer todos de uma vez agora!")

except Exception as e:
    print(f"❌ Erro: {e}")

input("\nPressione Enter para fechar...")