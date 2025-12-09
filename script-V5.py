# import pandas as pd
# import glob
# import os
# import warnings
#
# # Ignora avisos
# warnings.simplefilter(action='ignore', category=UserWarning)
#
# # ================= CONFIGURAÇÕES =================
# CAMINHO_INFOPRICE = r"C:\\Users\\Leonardo.Galdino\\Desktop\\scripts\\InfoPrice\\Dados Painel InfoPrice.xlsx"
# ARQUIVO_CURVA_A = r"C:\\Users\\Leonardo.Galdino\\Desktop\\scripts\\InfoPrice\\Curva A Compradores 01-01 á 25-11.xlsx"
#
# # === LISTAS DE BLOQUEIO ===
# REDES_IGNORAR = ['CONTRUMIL', 'CONSTRUMIL', 'SUPERMERCADOS BRAMIL', 'BRAMIL']
#
# # === LISTAS ESTRATÉGICAS ===
#
# # 1. RIO DE JANEIRO (AGRESSIVOS)
# REDES_RIO_AGRESSIVAS = [
#     'AMOEDO', 'CHATUBA', 'MASTERSON', 'LEROY MERLIN', 'OBRAMAX',
#     'CAÇULA', 'VILAREJO', 'TELHA LAR', 'REDE CONSTRUIR', 'SHOW BRASIL',
#     'CENCOSUD', 'GUANABARA_CONSTRUCOES', 'CASA SHOW'
# ]
#
# # 2. CLUSTER CONSTRUMIL (LOCAL)
# CIDADES_MIGUEL = ['TRÊS RIOS', 'TRE RIOS', 'TRES RIOS', 'PARAIBA DO SUL', 'PARAÍBA DO SUL', 'BARRA MANSA', 'PATY',
#                   'PATY DO ALFERES']
# REDES_IMPORTANTES_MIGUEL = [
#     'HELIO DUTRA', 'ARTE CONSTRUIR', 'WBRUN', 'GALPÃO', 'GALPAO',
#     'VIMACOL', 'JALEV', 'ABC', 'CASA WALTER', 'OBRA PRIMA'
# ]
#
# # 3. CLUSTER PRETÓPOLIS (LOCAL)
# CIDADES_PETROPOLIS = ['PETRÓPOLIS', 'PETROPOLIS']
# REDES_IMPORTANTES_PETROPOLIS = [
#     'AMOEDO', 'NA CERTA', 'ABC', 'GERMASIL', 'SÓ SUCESSO', 'SO SUCESSO',
#     'SHOW PISOS', 'PETROSAMPA', 'FERRAGENS NOEL', 'CARRAPETA',
#     'FERRAGENS VITRINE', 'MASTERSON', 'CHATUBA'
# ]
#
# # =================================================
#
# print(">>> 1. Iniciando Inteligência V5 (Petrópolis Agressiva)...")
#
#
# def limpar_ean(valor):
#     if pd.isna(valor) or valor == '': return None
#     return str(valor).replace('.0', '').strip()
#
#
# def aplicar_trava_com_status(preco_sugerido, preco_atual, margem_percentual):
#     if pd.isna(preco_sugerido) or pd.isna(preco_atual) or preco_atual == 0:
#         return round(preco_atual, 2), "Sem alteração"
#
#     teto = preco_atual * (1 + margem_percentual)
#     piso = preco_atual * (1 - margem_percentual)
#     porcentagem_txt = int(margem_percentual * 100)
#
#     valor_final = preco_sugerido
#     status = ""
#
#     if preco_sugerido > teto:
#         valor_final = teto
#         status = f"TRAVA TETO (+{porcentagem_txt}%)"
#     elif preco_sugerido < piso:
#         valor_final = piso
#         status = f"TRAVA PISO (-{porcentagem_txt}%)"
#     else:
#         if round(preco_sugerido, 2) != round(preco_atual, 2):
#             diff = ((preco_sugerido / preco_atual) - 1) * 100
#             valor_final = preco_sugerido
#             status = f"Livre ({diff:.1f}%)"
#         else:
#             valor_final = preco_atual
#             status = "Mantido"
#
#     return round(valor_final, 2), status
#
#
# # --- CARREGAR DADOS ---
# print(">>> Lendo planilhas...")
# try:
#     try:
#         df_construmil = pd.read_csv(ARQUIVO_CURVA_A, encoding='latin1', sep=None, engine='python')
#     except:
#         df_construmil = pd.read_excel(ARQUIVO_CURVA_A)
#
#     if 'Código Acesso' in df_construmil.columns:
#         df_construmil['EAN_BUSCA'] = df_construmil['Código Acesso'].apply(limpar_ean)
#     else:
#         df_construmil['EAN_BUSCA'] = df_construmil.iloc[:, 4].apply(limpar_ean)
#
#     if os.path.isdir(CAMINHO_INFOPRICE):
#         arquivos = glob.glob(os.path.join(CAMINHO_INFOPRICE, "*.xlsx"))
#         arquivo_final = max(arquivos, key=os.path.getctime)
#     else:
#         arquivo_final = CAMINHO_INFOPRICE
#
#     try:
#         df_infoprice = pd.read_csv(arquivo_final, encoding='latin1', sep=None, engine='python')
#     except:
#         df_infoprice = pd.read_excel(arquivo_final)
#
#     # Limpezas
#     cols_desejadas = ["Cidade", "Rede", "Identificador Produto", "Preço Pago"]
#     cols_existentes = [c for c in cols_desejadas if c in df_infoprice.columns]
#     df_infoprice = df_infoprice[cols_existentes]
#     df_infoprice['EAN_BUSCA'] = df_infoprice['Identificador Produto'].apply(limpar_ean)
#
#     if df_infoprice['Preço Pago'].dtype == object:
#         df_infoprice['Preço Pago'] = df_infoprice['Preço Pago'].astype(str).str.replace(',', '.').astype(float)
#
#     df_infoprice['Cidade'] = df_infoprice['Cidade'].astype(str).str.upper()
#     df_infoprice['Rede'] = df_infoprice['Rede'].astype(str).str.upper()
#
#     mask_ignorar = df_infoprice['Rede'].apply(lambda x: any(ignorar in x for ignorar in REDES_IGNORAR))
#     df_infoprice = df_infoprice[~mask_ignorar]
#
#     print(f"✓ Dados carregados.")
#
# except Exception as e:
#     print(f"❌ Erro leitura: {e}")
#     exit()
#
# # --- CÁLCULO INTELIGENTE ---
# print(">>> Calculando preços...")
#
# res_miguel, status_miguel, regra_miguel = [], [], []
# res_petropolis, status_petropolis, regra_petropolis = [], [], []
# res_conveniencia, status_conveniencia = [], []
# obs_list = []
#
# for index, row in df_construmil.iterrows():
#     ean = row['EAN_BUSCA']
#     preco_atual = row['Preço Vda Unitário']
#
#     df_match = df_infoprice[df_infoprice['EAN_BUSCA'] == ean]
#
#     # Preços Base
#     df_rio = df_match[(df_match['Cidade'] == 'RIO DE JANEIRO') & (df_match['Rede'].isin(REDES_RIO_AGRESSIVAS))]
#     p_rio = df_rio['Preço Pago'].min() if not df_rio.empty else None
#
#     df_local_miguel = df_match[df_match['Cidade'].isin(CIDADES_MIGUEL)]
#     df_redes_miguel = df_local_miguel[df_local_miguel['Rede'].isin(REDES_IMPORTANTES_MIGUEL)]
#     if df_redes_miguel.empty: df_redes_miguel = df_local_miguel
#     p_miguel = df_redes_miguel['Preço Pago'].min() if not df_redes_miguel.empty else None
#
#     df_local_petro = df_match[df_match['Cidade'].isin(CIDADES_PETROPOLIS)]
#     df_redes_petro = df_local_petro[df_local_petro['Rede'].isin(REDES_IMPORTANTES_PETROPOLIS)]
#     if df_redes_petro.empty: df_redes_petro = df_local_petro
#     p_petro = df_redes_petro['Preço Pago'].min() if not df_redes_petro.empty else None
#
#     # === LÓGICA CLUSTER 1 (MIGUEL) - CASCATA ===
#     # Prioriza defender a casa: Local > Vizinho > Rio
#
#     sugestao_m = preco_atual
#     margem_m = 0.05
#     regra_m = "Manter"
#
#     if p_miguel:
#         sugestao_m = p_miguel
#         margem_m = 0.05
#         regra_m = "1. Local (5%)"
#     elif p_petro:
#         sugestao_m = p_petro
#         margem_m = 0.05
#         regra_m = "2. Petrópolis (5%)"
#     elif p_rio:
#         sugestao_m = p_rio
#         margem_m = 0.02
#         regra_m = "3. Rio (2%)"
#     else:
#         regra_m = "Sem Ref"
#
#     val_m, stat_m = aplicar_trava_com_status(sugestao_m, preco_atual, margem_m)
#
#     # === LÓGICA CLUSTER 2 (PETRÓPOLIS) - COMPETIÇÃO ===
#     # Aqui queremos ser baratos. Compara Rio vs Petrópolis e pega o MENOR.
#
#     sugestao_p = preco_atual
#     margem_p = 0.05
#     regra_p = "Manter"
#
#     # Caso 1: Tem Rio e Tem Petrópolis -> Duelo de Preços
#     if p_rio and p_petro:
#         if p_rio < p_petro:
#             sugestao_p = p_rio
#             margem_p = 0.02  # Rio ganhou, margem segura
#             regra_p = "Rio < Pet (2%)"
#         else:
#             sugestao_p = p_petro
#             margem_p = 0.05  # Pet ganhou, margem normal
#             regra_p = "Pet < Rio (5%)"
#
#     # Caso 2: Só tem Rio
#     elif p_rio:
#         sugestao_p = p_rio
#         margem_p = 0.02
#         regra_p = "Só Rio (2%)"
#
#     # Caso 3: Só tem Petrópolis
#     elif p_petro:
#         sugestao_p = p_petro
#         margem_p = 0.05
#         regra_p = "Só Pet (5%)"
#
#     # Caso 4: Não tem nenhum dos dois -> Fallback para Miguel
#     else:
#         sugestao_p = val_m
#         regra_p = "Fallback Miguel"
#
#     val_p, stat_p = aplicar_trava_com_status(sugestao_p, preco_atual, margem_p)
#
#     # === CLUSTER 3 (CONVENIÊNCIA) ===
#     val_c, stat_c = aplicar_trava_com_status(preco_atual * 1.05, preco_atual, 0.05)
#
#     # Salvando Listas
#     res_miguel.append(val_m)
#     status_miguel.append(stat_m)
#     regra_miguel.append(regra_m)
#
#     res_petropolis.append(val_p)
#     status_petropolis.append(stat_p)
#     regra_petropolis.append(regra_p)
#
#     res_conveniencia.append(val_c)
#     status_conveniencia.append(stat_c)
#
#     obs_list.append(f"Rio:{p_rio} | Mig:{p_miguel} | Pet:{p_petro}")
#
# # --- SALVAR ---
# print(">>> Salvando planilha final...")
# df_construmil['Cluster Construmil e Região Miguel'] = res_miguel
# df_construmil['Status_Miguel'] = status_miguel
# df_construmil['Regra_Miguel'] = regra_miguel
#
# df_construmil['Cluster Pretópolis'] = res_petropolis
# df_construmil['Status_Petropolis'] = status_petropolis
# df_construmil['Regra_Petropolis'] = regra_petropolis
#
# df_construmil['Cluster Conveniência'] = res_conveniencia
# df_construmil['Status_Conveniencia'] = status_conveniencia
#
# df_construmil['Origem_Dados'] = obs_list
#
# pasta_saida = os.path.dirname(ARQUIVO_CURVA_A)
# nome_final = os.path.join(pasta_saida, "Curva_A_PRECIFICADA_FINAL_V5.xlsx")
#
# df_construmil.to_excel(nome_final, index=False)
# print(f"✅ Sucesso! Arquivo gerado: {nome_final}")


import pandas as pd
import glob
import os
import warnings
import numpy as np

# Ignora avisos
warnings.simplefilter(action='ignore', category=UserWarning)

# ================= CONFIGURAÇÕES =================
CAMINHO_INFOPRICE = r"C:\\Users\\Leonardo.Galdino\\Desktop\\scripts\\InfoPrice\\Dados Painel InfoPrice.xlsx"
ARQUIVO_CURVA_A = r"C:\\Users\\Leonardo.Galdino\\Desktop\\scripts\\InfoPrice\\Curva A Compradores 01-01 á 25-11.xlsx"

# === LISTAS DE BLOQUEIO ===
REDES_IGNORAR = ['CONTRUMIL', 'CONSTRUMIL', 'SUPERMERCADOS BRAMIL', 'BRAMIL']

# === LISTAS ESTRATÉGICAS ===
REDES_RIO_AGRESSIVAS = [
    'AMOEDO', 'CHATUBA', 'MASTERSON', 'LEROY MERLIN', 'OBRAMAX',
    'CAÇULA', 'VILAREJO', 'TELHA LAR', 'REDE CONSTRUIR', 'SHOW BRASIL',
    'CENCOSUD', 'GUANABARA_CONSTRUCOES', 'CASA SHOW'
]

CIDADES_MIGUEL = ['TRÊS RIOS', 'TRE RIOS', 'TRES RIOS', 'PARAIBA DO SUL', 'PARAÍBA DO SUL', 'BARRA MANSA', 'PATY',
                  'PATY DO ALFERES']
REDES_IMPORTANTES_MIGUEL = [
    'HELIO DUTRA', 'ARTE CONSTRUIR', 'WBRUN', 'GALPÃO', 'GALPAO',
    'VIMACOL', 'JALEV', 'ABC', 'CASA WALTER', 'OBRA PRIMA'
]

CIDADES_PETROPOLIS = ['PETRÓPOLIS', 'PETROPOLIS']
REDES_IMPORTANTES_PETROPOLIS = [
    'AMOEDO', 'NA CERTA', 'ABC', 'GERMASIL', 'SÓ SUCESSO', 'SO SUCESSO',
    'SHOW PISOS', 'PETROSAMPA', 'FERRAGENS NOEL', 'CARRAPETA',
    'FERRAGENS VITRINE', 'MASTERSON', 'CHATUBA'
]
# =================================================

print(">>> 1. Iniciando Inteligência V6 (Com Preço Psicológico .90/.95/.99)...")


def limpar_ean(valor):
    if pd.isna(valor) or valor == '': return None
    return str(valor).replace('.0', '').strip()


def aplicar_preco_psicologico(preco):
    """
    Ajusta o preço para terminar em .90, .95 ou .99
    Escolhendo a opção com a menor distância do valor original.
    """
    if pd.isna(preco) or preco == 0: return preco

    parte_inteira = int(preco)

    # Cria as opções "mágicas" ao redor do preço
    opcoes = [
        parte_inteira - 1 + 0.99,  # Ex: Se é 14.10, tenta 13.99
        parte_inteira + 0.90,  # Ex: 14.90
        parte_inteira + 0.95,  # Ex: 14.95
        parte_inteira + 0.99  # Ex: 14.99
    ]

    # Escolhe a opção que tem a menor diferença absoluta para o preço original
    # Ex: Se o preço é 14.88, o mais perto é 14.90 (dif 0.02)
    # Ex: Se o preço é 14.10, o mais perto é 13.99 (dif 0.11)
    melhor_opcao = min(opcoes, key=lambda x: abs(x - preco))

    return melhor_opcao


def aplicar_trava_com_status(preco_sugerido, preco_atual, margem_percentual):
    if pd.isna(preco_sugerido) or pd.isna(preco_atual) or preco_atual == 0:
        return round(preco_atual, 2), "Sem alteração"

    teto = preco_atual * (1 + margem_percentual)
    piso = preco_atual * (1 - margem_percentual)
    porcentagem_txt = int(margem_percentual * 100)

    valor_travado = preco_sugerido
    status_base = ""

    if preco_sugerido > teto:
        valor_travado = teto
        status_base = f"TRAVA TETO (+{porcentagem_txt}%)"
    elif preco_sugerido < piso:
        valor_travado = piso
        status_base = f"TRAVA PISO (-{porcentagem_txt}%)"
    else:
        if round(preco_sugerido, 2) != round(preco_atual, 2):
            diff = ((preco_sugerido / preco_atual) - 1) * 100
            valor_travado = preco_sugerido
            status_base = f"Livre ({diff:.1f}%)"
        else:
            valor_travado = preco_atual
            status_base = "Mantido"

    return valor_travado, status_base


# --- CARREGAR DADOS ---
print(">>> Lendo planilhas...")
try:
    try:
        df_construmil = pd.read_csv(ARQUIVO_CURVA_A, encoding='latin1', sep=None, engine='python')
    except:
        df_construmil = pd.read_excel(ARQUIVO_CURVA_A)

    if 'Código Acesso' in df_construmil.columns:
        df_construmil['EAN_BUSCA'] = df_construmil['Código Acesso'].apply(limpar_ean)
    else:
        df_construmil['EAN_BUSCA'] = df_construmil.iloc[:, 4].apply(limpar_ean)

    if os.path.isdir(CAMINHO_INFOPRICE):
        arquivos = glob.glob(os.path.join(CAMINHO_INFOPRICE, "*.xlsx"))
        arquivo_final = max(arquivos, key=os.path.getctime)
    else:
        arquivo_final = CAMINHO_INFOPRICE

    try:
        df_infoprice = pd.read_csv(arquivo_final, encoding='latin1', sep=None, engine='python')
    except:
        df_infoprice = pd.read_excel(arquivo_final)

    cols_desejadas = ["Cidade", "Rede", "Identificador Produto", "Preço Pago"]
    cols_existentes = [c for c in cols_desejadas if c in df_infoprice.columns]
    df_infoprice = df_infoprice[cols_existentes]
    df_infoprice['EAN_BUSCA'] = df_infoprice['Identificador Produto'].apply(limpar_ean)

    if df_infoprice['Preço Pago'].dtype == object:
        df_infoprice['Preço Pago'] = df_infoprice['Preço Pago'].astype(str).str.replace(',', '.').astype(float)

    df_infoprice['Cidade'] = df_infoprice['Cidade'].astype(str).str.upper()
    df_infoprice['Rede'] = df_infoprice['Rede'].astype(str).str.upper()

    mask_ignorar = df_infoprice['Rede'].apply(lambda x: any(ignorar in x for ignorar in REDES_IGNORAR))
    df_infoprice = df_infoprice[~mask_ignorar]

    print(f"✓ Dados carregados.")

except Exception as e:
    print(f"❌ Erro leitura: {e}")
    exit()

# --- CÁLCULO ---
print(">>> Calculando preços (Técnico -> Trava -> Psicológico)...")

res_miguel, status_miguel, regra_miguel = [], [], []
res_petropolis, status_petropolis, regra_petropolis = [], [], []
res_conveniencia, status_conveniencia = [], []
obs_list = []

for index, row in df_construmil.iterrows():
    ean = row['EAN_BUSCA']
    preco_atual = row['Preço Vda Unitário']

    df_match = df_infoprice[df_infoprice['EAN_BUSCA'] == ean]

    # Busca Preços
    df_rio = df_match[(df_match['Cidade'] == 'RIO DE JANEIRO') & (df_match['Rede'].isin(REDES_RIO_AGRESSIVAS))]
    p_rio = df_rio['Preço Pago'].min() if not df_rio.empty else None

    df_local_miguel = df_match[df_match['Cidade'].isin(CIDADES_MIGUEL)]
    df_redes_miguel = df_local_miguel[df_local_miguel['Rede'].isin(REDES_IMPORTANTES_MIGUEL)]
    if df_redes_miguel.empty: df_redes_miguel = df_local_miguel
    p_miguel = df_redes_miguel['Preço Pago'].min() if not df_redes_miguel.empty else None

    df_local_petro = df_match[df_match['Cidade'].isin(CIDADES_PETROPOLIS)]
    df_redes_petro = df_local_petro[df_local_petro['Rede'].isin(REDES_IMPORTANTES_PETROPOLIS)]
    if df_redes_petro.empty: df_redes_petro = df_local_petro
    p_petro = df_redes_petro['Preço Pago'].min() if not df_redes_petro.empty else None

    # === MIGUEL (CASCATA) ===
    sugestao_m = preco_atual
    margem_m = 0.05
    regra_m = "Manter"

    if p_miguel:
        sugestao_m, margem_m, regra_m = p_miguel, 0.05, "1. Local (5%)"
    elif p_petro:
        sugestao_m, margem_m, regra_m = p_petro, 0.05, "2. Petrópolis (5%)"
    elif p_rio:
        sugestao_m, margem_m, regra_m = p_rio, 0.02, "3. Rio (2%)"
    else:
        regra_m = "Sem Ref"

    val_m, stat_m = aplicar_trava_com_status(sugestao_m, preco_atual, margem_m)
    # APLICA PSICOLÓGICO
    val_m_psico = aplicar_preco_psicologico(val_m)
    if val_m_psico != val_m: stat_m += " [Psico]"

    # === PETRÓPOLIS (COMPETIÇÃO) ===
    sugestao_p = preco_atual
    margem_p = 0.05
    regra_p = "Manter"

    if p_rio and p_petro:
        if p_rio < p_petro:
            sugestao_p, margem_p, regra_p = p_rio, 0.02, "Rio < Pet (2%)"
        else:
            sugestao_p, margem_p, regra_p = p_petro, 0.05, "Pet < Rio (5%)"
    elif p_rio:
        sugestao_p, margem_p, regra_p = p_rio, 0.02, "Só Rio (2%)"
    elif p_petro:
        sugestao_p, margem_p, regra_p = p_petro, 0.05, "Só Pet (5%)"
    else:
        sugestao_p = val_m  # Fallback pega o valor JÁ calculado de Miguel (que pode ter psico)
        regra_p = "Fallback Miguel"

    val_p, stat_p = aplicar_trava_com_status(sugestao_p, preco_atual, margem_p)
    # APLICA PSICOLÓGICO
    val_p_psico = aplicar_preco_psicologico(val_p)
    if val_p_psico != val_p: stat_p += " [Psico]"

    # === CONVENIÊNCIA ===
    # Preço Atual + 5% -> Travado -> Psicológico
    sugestao_c = preco_atual * 1.05
    val_c, stat_c = aplicar_trava_com_status(sugestao_c, preco_atual, 0.05)
    val_c_psico = aplicar_preco_psicologico(val_c)
    if val_c_psico != val_c: stat_c += " [Psico]"

    # Salvando
    res_miguel.append(val_m_psico)
    status_miguel.append(stat_m)
    regra_miguel.append(regra_m)

    res_petropolis.append(val_p_psico)
    status_petropolis.append(stat_p)
    regra_petropolis.append(regra_p)

    res_conveniencia.append(val_c_psico)
    status_conveniencia.append(stat_c)

    obs_list.append(f"Rio:{p_rio} | Mig:{p_miguel} | Pet:{p_petro}")

# --- SALVAR ---
print(">>> Salvando planilha final...")
df_construmil['Cluster Construmil e Região Miguel'] = res_miguel
df_construmil['Status_Miguel'] = status_miguel
df_construmil['Regra_Miguel'] = regra_miguel

df_construmil['Cluster Pretópolis'] = res_petropolis
df_construmil['Status_Petropolis'] = status_petropolis
df_construmil['Regra_Petropolis'] = regra_petropolis

df_construmil['Cluster Conveniência'] = res_conveniencia
df_construmil['Status_Conveniencia'] = status_conveniencia

df_construmil['Origem_Dados'] = obs_list

pasta_saida = os.path.dirname(ARQUIVO_CURVA_A)
nome_final = os.path.join(pasta_saida, "Curva_A_PRECIFICADA_V6.xlsx")

df_construmil.to_excel(nome_final, index=False)
print(f"✅ Sucesso! Arquivo com marketing aplicado: {nome_final}")