import pandas as pd
import openpyxl
import numpy as np
from datetime import date, datetime, time as dtime
import time
import seaborn as sns
import matplotlib.pyplot as plt
import re
import math
import pytz
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────
# CONFIGURAÇÃO: ajuste conforme seu ambiente
# ─────────────────────────────────────────

USE_COLAB = False  # Mude para True se estiver rodando no Google Colab

if USE_COLAB:
    from google.colab import drive
    drive.mount('/content/drive')
    BASE_DIR = Path("/content/drive/MyDrive/seu_projeto")
else:
    BASE_DIR = Path("./data")  # pasta local com os arquivos

OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ─────────────────────────────────────────
# VARIÁVEIS GLOBAIS
# ─────────────────────────────────────────

# Faixas de login por fornecedor (ajuste conforme o projeto)
intervalos_logins = [100, 399]
fornecedores = ['Fornecedor A']

# Timezone
tz_brasilia = pytz.timezone('America/Sao_Paulo')
agora = datetime.now(tz_brasilia)
data_hoje = agora.date()
data_hoje_str = agora.strftime('%d/%m/%Y')
data_diames = agora.strftime('%d%m')

# Mapeamento de respostas padrão
codigo_respostas = {1: 'Sim', 2: 'Não', 3: 'Não Sei', 4: 'Recusou'}

# Colunas de GPS esperadas
colunas_gps = [
    'LOC_START_LAT', 'LOC_START_LONG', 'LOC_MAIN_LAT', 'LOC_MAIN_LONG',
    'HH_LOC_LAT', 'HH_LOC_LONG', 'LOC_MAIN_END_LAT', 'LOC_MAIN_END_LONG'
]

# Ponto médio das faixas de renda (variável WP19188)
wp19188_midpoint = {
    0:  0.0,
    1:  55.0,
    2:  180.5,
    3:  500.5,
    4:  875.5,
    5:  1250.5,
    6:  1750.5,
    7:  2500.5,
    8:  3500.5,
    9:  5500.5,
    10: 8500.5,
    11: 12500.5,
    12: 20000.0,
}

# Códigos de resultado de visita
codigo_visita_desc = {
    1:  "Entrevista completa",
    2:  "Entrevista interrompida",
    3:  "Recusa do respondente ou domicílio",
    4:  "NEC",
    5:  "Respondente ausente pelo restante do campo",
    6:  "Respondente temporariamente ausente",
    7:  "Acesso negado",
    8:  "Doente/Hospitalizado/Deficiência mental",
    9:  "Barreira linguística",
    10: "Nenhum morador elegível no domicílio",
    11: "Qualquer outro motivo",
}

CODIGOS_AGENDAMENTO = {2, 4, 6}
CODIGOS_ENCERRA = {1, 3, 5, 7, 8, 9, 10, 11}

# ─────────────────────────────────────────
# CAMINHOS DOS ARQUIVOS
# ─────────────────────────────────────────

opconsole_path = BASE_DIR / "OPConsole.xlsx"
basegeral_path = BASE_DIR / "BaseGeral.csv"
base_path      = BASE_DIR / f"Base{data_diames}.csv"

# ─────────────────────────────────────────
# LEITURA DOS ARQUIVOS
# ─────────────────────────────────────────

BaseGeral = pd.read_csv(basegeral_path, low_memory=False)
Base      = pd.read_csv(base_path, low_memory=False)
OPConsole = pd.read_excel(opconsole_path, engine="openpyxl")

# Valida coluna-chave
for df, nome in [(BaseGeral, "BaseGeral"), (Base, "Base"), (OPConsole, "OPConsole")]:
    if "DeviceIndex" not in df.columns:
        raise KeyError(f"A coluna 'DeviceIndex' não existe em {nome}.")

# Padroniza chave de merge
for df in [BaseGeral, Base, OPConsole]:
    df["DeviceIndex"] = df["DeviceIndex"].astype(str).str.strip()

# Merge com OPConsole
if not BaseGeral.empty:
    BaseGeral = BaseGeral.merge(OPConsole, on="DeviceIndex", how="left", suffixes=("", "_OPConsole"))

if not Base.empty:
    Base = Base.merge(OPConsole, on="DeviceIndex", how="left", suffixes=("", "_OPConsole"))


# ─────────────────────────────────────────
# FUNÇÕES DE DESCRIÇÃO DOS PROBLEMAS
# ─────────────────────────────────────────

def membros_desc(x, y):
    return f"Divergência no número de membros do domicílio: {x} na tabela Kish e {y} no questionário (WP8991)."

def idade_desc(x, y):
    return f"Divergência de idade: {x} anos na tabela Kish e {y} anos no questionário."

def sexo_desc(x):
    if x == 1:
        return "Divergência de sexo: Masculino na tabela Kish e Feminino no questionário."
    if x == 2:
        return "Divergência de sexo: Feminino na tabela Kish e Masculino no questionário."

def escolaridade_desc(x, y):
    niveis = {
        0: "sem escolaridade",
        1: "Ensino Fundamental 1 Incompleto (até 5 anos)",
        2: "Ensino Fundamental 1 Completo / EF2 Incompleto (4–9 anos)",
        3: "Ensino Fundamental 2 Completo / EM Incompleto (8–12 anos)",
        4: "Ensino Médio Completo (10–13 anos)",
        5: "Ensino Superior Incompleto (até 15 anos)",
        6: "Ensino Superior Completo (16+ anos)",
    }
    nivel_str = niveis.get(y, f"nível {y}")
    return f"Divergência de escolaridade: declarou {nivel_str}, mas anos de estudo registrado = {x}."

def renda_desc(x):
    return f"Renda declarada fora do intervalo esperado: R${round(x, 2)} mensal."

def gasto_alimentacao_desc(gasto, renda, fonte_renda="WP7133"):
    gasto_txt = f"R${round(gasto, 2)}" if pd.notna(gasto) else "NA"
    renda_txt = f"R${round(renda, 2)}" if pd.notna(renda) else "NA"
    return (
        f"Inconsistência entre gasto com alimentação (SPEND_W_F) e renda ({fonte_renda}). "
        f"Gasto: {gasto_txt} | Renda: {renda_txt}."
    )

def menor_desc(x):
    return f"Entrevista realizada com respondente de {x} anos sem autorização registrada."

def duracao_desc(x):
    return f"Duração abaixo do mínimo esperado: {round(x, 2)} minutos."

def duracao_longa_desc(x):
    return f"Duração acima de 1h10: {round(x, 2)} minutos."

def Membros_menos15anos_desc(x):
    return f"Número elevado de moradores com menos de 15 anos: {x}. Verificar variável WP1230."

def parcial_desc(x):
    return f"Alta velocidade de resposta: {x}% das perguntas respondidas em menos de 3 segundos."

def visitas_nec_desc(data_visita, qtd_nec):
    return (
        f"3 ou mais visitas NEC (ninguém em casa) no mesmo dia ({data_visita}) "
        f"para o mesmo domicílio. Total: {qtd_nec} visitas."
    )

def intervalo_tentativas_desc(idx_a, idx_b, minutos, code_a, code_b):
    desc_a = codigo_visita_desc.get(code_a, f"Código {code_a}")
    desc_b = codigo_visita_desc.get(code_b, f"Código {code_b}")
    return (
        f"Intervalo insuficiente entre tentativas no mesmo domicílio. "
        f"Tentativas {idx_a} ({desc_a}) e {idx_b} ({desc_b}) realizadas "
        f"com {round(minutos, 1)} minutos de intervalo (mínimo: 120 min)."
    )

def alta_primeira_tentativa_psu_desc(psu, pct):
    return (
        f"PSU '{psu}' concluída com {pct:.1f}% das entrevistas na primeira tentativa "
        f"(acima de 50%). Comportamento atípico."
    )

def entrevistas_dia_desc(srvyr, data, count):
    return (
        f"Entrevistador '{srvyr}' completou {count} entrevistas no dia {data} "
        f"(acima do limite de 10)."
    )

def baixa_audio_psu_desc(psu, pct_audio):
    return (
        f"PSU '{psu}' concluída com apenas {pct_audio:.1f}% de autorização de áudio "
        f"(abaixo de 50%)."
    )

def horario_noturno_desc(vend_time):
    return (
        f"Entrevista finalizada em horário noturno (entre 21h e 6h). "
        f"Horário de término: {vend_time.strftime('%H:%M:%S')}."
    )


# ─────────────────────────────────────────
# FUNÇÕES UTILITÁRIAS
# ─────────────────────────────────────────

def to_number_or_nan(v, missing_codes=(98, 99)):
    if pd.isna(v):
        return np.nan
    try:
        s = str(v).strip().replace(",", ".")
        if s == "":
            return np.nan
        x = float(s)
        return np.nan if x in missing_codes else x
    except:
        return np.nan

def salvar_excel(df, nome_arquivo):
    caminho = OUTPUT_DIR / f"{nome_arquivo}.xlsx"
    df.to_excel(caminho, index=False)
    print(f"Arquivo salvo: {caminho}")

def confirmar(df, esperado):
    if len(df) == esperado:
        print("Todas as inconsistências foram registradas corretamente.")
    else:
        print(f"Atenção: esperado {esperado} registros, encontrado {len(df)}.")

def transformar_float(x):
    try:
        return float(str(x).replace(",", "."))
    except:
        return 0.0

def distancia2d(x1, y1, x2, y2):
    return math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)

def _parse_date_safe(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    if s in ("", "nan", "none"):
        return None
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return None if pd.isna(dt) else dt.date()

def _parse_int_safe(v):
    if pd.isna(v):
        return None
    try:
        return int(float(str(v).strip().replace(",", ".")))
    except:
        return None

def _parse_datetime_visit(row, i):
    d = _parse_date_safe(row.get(f"I_{i}_VisitDate"))
    if d is None:
        return None
    t_raw = row.get(f"I_{i}_VisitTime")
    if pd.isna(t_raw):
        return None
    t_str = str(t_raw).strip()
    if t_str in ("", "nan", "none"):
        return None
    try:
        parts = t_str.split(":")
        h, m = int(parts[0]), int(parts[1])
        s = int(parts[2]) if len(parts) > 2 else 0
        return datetime.combine(d, dtime(h, m, s))
    except:
        return None

def as_int_nullable(series):
    s = series.astype("string").str.strip()
    s = s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    return pd.to_numeric(s, errors="coerce").astype("Int64")

def as_float(series):
    s = series.astype("string").str.strip()
    s = s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    return pd.to_numeric(s, errors="coerce")

def as_str_clean(series):
    s = series.astype("string").str.strip()
    return s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})

def first_or_na(x):
    if isinstance(x, (list, tuple, np.ndarray)) and len(x) > 0:
        return x[0]
    return pd.NA

def join_list(x):
    if isinstance(x, (list, tuple, np.ndarray)):
        return "".join([str(i) for i in x])
    return "Não Encontrado"


# ─────────────────────────────────────────
# FUNÇÕES DE DETECÇÃO DE PROBLEMAS
# ─────────────────────────────────────────

def get_nec_3plus_same_day_info(row, max_visits=150):
    """Detecta 3 ou mais visitas NEC no mesmo dia para o mesmo domicílio."""
    nec_counts = {}
    for dom in range(max_visits // 3):
        i_start = dom * 3 + 1
        for i in range(i_start, i_start + 3):
            d = _parse_date_safe(row.get(f"I_{i}_VisitDate"))
            if d is None:
                continue
            if _parse_int_safe(row.get(f"I_{i}_Code")) == 4:
                key = (dom, d)
                nec_counts[key] = nec_counts.get(key, 0) + 1
    for (_, d), cnt in nec_counts.items():
        if cnt >= 3:
            return (True, d.strftime("%d/%m/%Y"), cnt)
    return (False, None, 0)

def get_intervalo_tentativas_info(row, max_visits=150):
    """Detecta tentativas consecutivas no mesmo domicílio com menos de 2h de intervalo."""
    for dom in range(max_visits // 3):
        i_start = dom * 3 + 1
        visitas = []
        for i in range(i_start, i_start + 3):
            dt = _parse_datetime_visit(row, i)
            code = _parse_int_safe(row.get(f"I_{i}_Code"))
            if dt and code:
                visitas.append((i, dt, code))
        visitas.sort(key=lambda x: x[1])
        for k in range(len(visitas) - 1):
            idx_a, dt_a, code_a = visitas[k]
            idx_b, dt_b, code_b = visitas[k + 1]
            if code_a in CODIGOS_AGENDAMENTO or code_b in CODIGOS_AGENDAMENTO:
                continue
            diff_min = (dt_b - dt_a).total_seconds() / 60.0
            if diff_min < 120:
                return (True, idx_a, idx_b, diff_min, code_a, code_b)
    return (False, None, None, None, None, None)

def check_horario_noturno(vend_datetime):
    """Verifica se a entrevista foi finalizada entre 21h e 6h."""
    if pd.isna(vend_datetime):
        return False
    hora = vend_datetime.hour
    return hora >= 21 or hora < 6


# ─────────────────────────────────────────
# FUNÇÕES DE EXTRAÇÃO DE TENTATIVAS
# ─────────────────────────────────────────

def extrair_tentativas(df, col_srvyr="Srvyr_original", max_visits=150):
    """Extrai todas as tentativas de visita com indicadores de horário."""
    linhas = []
    for _, row in df.iterrows():
        for i in range(1, max_visits + 1):
            dt = _parse_datetime_visit(row, i)
            if dt is None:
                continue
            linhas.append({
                "Srvyr":       row.get(col_srvyr, pd.NA),
                "PSU_tratada": row.get("PSU_tratada", pd.NA),
                "SbjNum":      row.get("SbjNum", pd.NA),
                "VisitIndex":  i,
                "VisitDateTime": dt,
                "Is_FDS":      dt.weekday() in (5, 6),
                "Is_apos17":   dt.hour >= 17,
            })
    return pd.DataFrame(linhas)

def extrair_tentativas_com_outros(df, col_srvyr="Srvyr_original", max_visits=150):
    """Extrai tentativas incluindo o motivo 'Outros' (código 11)."""
    linhas = []
    for _, row in df.iterrows():
        for i in range(1, max_visits + 1):
            dt   = _parse_datetime_visit(row, i)
            code = _parse_int_safe(row.get(f"I_{i}_Code"))
            if dt is None and code is None:
                continue
            linhas.append({
                "Srvyr":           row.get(col_srvyr, pd.NA),
                "PSU_tratada":     row.get("PSU_tratada", pd.NA),
                "SbjNum":          row.get("SbjNum", pd.NA),
                "DeviceIndex":     row.get("DeviceIndex", pd.NA),
                "VisitIndex":      i,
                "VisitDateTime":   dt,
                "VisitCode":       code,
                "VisitCode_Desc":  codigo_visita_desc.get(code, "Desconhecido"),
                "Other_Reason":    row.get(f"I_{i}_Other_Reason", pd.NA) if code == 11 else pd.NA,
            })
    return pd.DataFrame(linhas)

def calcular_relatorio_horarios(tentativas_df):
    """Gera resumo global e por entrevistador de tentativas após 17h ou fins de semana."""
    cols = ["Tentativas_total", "Tentativas_17h_FDS", "Percentual_17h_FDS(%)", "Abaixo_de_20%"]
    if len(tentativas_df) == 0:
        return pd.DataFrame([{c: pd.NA for c in cols}]), pd.DataFrame(columns=["Srvyr"] + cols)

    df = tentativas_df.copy()
    df["Flag_17_FDS"] = df["Is_FDS"] | df["Is_apos17"]

    total = len(df)
    total_flag = int(df["Flag_17_FDS"].sum())
    pct = round(total_flag / total * 100, 1) if total > 0 else 0.0

    resumo_global = pd.DataFrame([{
        "Tentativas_total":    total,
        "Tentativas_17h_FDS":  total_flag,
        "Percentual_17h_FDS(%)": pct,
        "Abaixo_de_20%":       "Sim" if pct < 20 else "Não",
    }])

    agg = (
        df.groupby("Srvyr")["Flag_17_FDS"]
        .agg(Tentativas_total="count", Tentativas_17h_FDS="sum")
        .reset_index()
    )
    agg["Percentual_17h_FDS(%)"] = (agg["Tentativas_17h_FDS"] / agg["Tentativas_total"] * 100).round(1)
    agg["Abaixo_de_20%"] = np.where(agg["Percentual_17h_FDS(%)"] < 20, "Sim", "Não")
    return resumo_global, agg


# ─────────────────────────────────────────
# FUNÇÕES DE CHECAGEM POR LINHA
# ─────────────────────────────────────────

def check_escolaridade(row):
    """Verifica consistência entre anos de estudo e nível de escolaridade declarado."""
    if row['Years_educ'] < 98:
        if (row['Level_educ'] == 0 and row['Years_educ'] != 0) or \
           (row['Level_educ'] == 1 and not (0 <= row['Years_educ'] <= 5)) or \
           (row['Level_educ'] == 2 and not (4 <= row['Years_educ'] <= 9)) or \
           (row['Level_educ'] == 3 and not (8 <= row['Years_educ'] <= 12)) or \
           (row['Level_educ'] == 4 and not (10 <= row['Years_educ'] <= 13)) or \
           (row['Level_educ'] == 5 and not (12 <= row['Years_educ'] <= 15)) or \
           (row['Level_educ'] == 6 and row['Years_educ'] < 15):
            return True
    return False

def check_membros(row):
    """Verifica divergência entre membros declarados na Kish e no questionário."""
    kish = row.get('Kish_membros')
    ent  = row.get('Ent_membros')
    if pd.isna(kish) or pd.isna(ent):
        return False
    try:
        return int(float(str(kish).replace(",", "."))) != int(float(str(ent).replace(",", ".")))
    except:
        return kish != ent


# ─────────────────────────────────────────
# PROCESSAMENTO DA BASE
# ─────────────────────────────────────────

def process_dataframe_base(df, data_hoje):
    """Seleciona, renomeia e enriquece colunas da base de campo."""
    df = df.copy()

    visit_cols = []
    for i in range(1, 151):
        visit_cols.extend([f"I_{i}_VisitDate", f"I_{i}_VisitTime", f"I_{i}_Code", f"I_{i}_Other_Reason"])
    visit_cols = [c for c in visit_cols if c in df.columns]

    # ─────────────────────────────────────────────────────────────────
    # MAPEAMENTO DE VARIÁVEIS DO PROJETO
    # Adapte os nomes à esquerda para os nomes reais das colunas
    # na sua base de dados. Os nomes à direita são os nomes internos
    # usados pelo restante do código e não devem ser alterados.
    # ─────────────────────────────────────────────────────────────────
    VAR = {
        # Identificação e controle
        'col_entrevistador':          'VAR_ENTREVISTADOR',       # ID numérico do entrevistador
        'col_autorizou_recontato':    'VAR_AUTORIZOU_RECONTATO', # Autorizou ser recontactado? (1=Sim, 2=Não...)
        'col_duracao':                'VAR_DURACAO',             # Duração da entrevista em minutos
        'col_audio_gravado':          'VAR_AUDIO_GRAVADO',       # Autorizou gravação de áudio? (1=Sim, 2=Não...)
        'col_psu':                    'VAR_PSU',                 # Código do cluster / PSU
        'col_versao_questionario':    'VAR_VERSAO_QUEST',        # Versão do questionário aplicado
        'col_pais':                   'VAR_PAIS',                # País (quando aplicável)

        # Idade
        'col_idade_selecao':          'VAR_IDADE_SELECAO',       # Idade registrada na tabela de seleção (ex: Kish)
        'col_idade_questionario':     'VAR_IDADE_QUEST',         # Idade declarada no questionário
        'col_idade_questionario_dup': 'VAR_IDADE_QUEST_DUP',     # Campo duplo de confirmação de idade

        # Autorização de menor
        'col_autorizacao_menor':      'VAR_AUTOR_MENOR',         # Autorizou entrevista com menor? (1=Sim)

        # Sexo
        'col_sexo_selecao':           'VAR_SEXO_SELECAO',        # Sexo registrado na tabela de seleção
        'col_sexo_questionario':      'VAR_SEXO_QUEST',          # Sexo declarado no questionário

        # Membros do domicílio
        'col_membros_selecao':        'VAR_MEMBROS_SELECAO',     # Nº de membros na tabela de seleção
        'col_membros_questionario':   'VAR_MEMBROS_QUEST',       # Nº de membros declarado no questionário
        'col_membros_menos15':        'VAR_MEMBROS_MENOS15',     # Nº de moradores com menos de 15 anos

        # Escolaridade
        'col_anos_estudo':            'VAR_ANOS_ESTUDO',         # Anos de estudo completados
        'col_nivel_escolaridade':     'VAR_NIVEL_ESCOL',         # Nível de escolaridade (código categórico)

        # Renda e gasto
        'col_renda':                  'VAR_RENDA',               # Renda mensal declarada (valor)
        'col_renda_faixa':            'VAR_RENDA_FAIXA',         # Renda por faixa (código categórico)
        'col_gasto_alimentacao':      'VAR_GASTO_ALIM',          # Gasto mensal com alimentação

        # Qualidade de resposta
        'col_pct_respostas_rapidas':  'VAR_PCT_RAPIDAS',         # % de respostas em menos de 3 segundos
        'col_tentativas_completa':    'VAR_TENTATIVAS_COMPL',    # Nº de tentativas até completar a entrevista

        # GPS e localização
        'col_lat_referencia':         'VAR_LAT',                 # Latitude de referência fornecida
        'col_long_referencia':        'VAR_LONG',                # Longitude de referência fornecida

        # Identificação do respondente
        'col_nome':                   'VAR_NOME',                # Nome do respondente
        'col_telefone':               'VAR_TEL',                 # Telefone do respondente
    }

    # Lista de colunas brutas esperadas na base (usando os nomes reais do projeto)
    base_cols = [
        'SbjNum', 'DeviceIndex', 'Upload', 'VEnd', 'Status', 'FlagsText',
        'LOC_START_LA', 'LOC_START_LO', 'LOC_MAIN_LA', 'LOC_MAIN_LO',
        'LOC_WP12_LA', 'LOC_WP12_LO', 'LOC_MAIN_END_LA', 'LOC_MAIN_END_LO',
        'DIST_SP', 'LocName',
        VAR['col_entrevistador'],
        VAR['col_autorizou_recontato'],
        VAR['col_duracao'],
        VAR['col_audio_gravado'],
        VAR['col_psu'],
        VAR['col_versao_questionario'],
        VAR['col_pais'],
        VAR['col_idade_selecao'],
        VAR['col_idade_questionario'],
        VAR['col_idade_questionario_dup'],
        VAR['col_autorizacao_menor'],
        VAR['col_sexo_selecao'],
        VAR['col_sexo_questionario'],
        VAR['col_membros_selecao'],
        VAR['col_membros_questionario'],
        VAR['col_membros_menos15'],
        VAR['col_anos_estudo'],
        VAR['col_nivel_escolaridade'],
        VAR['col_renda'],
        VAR['col_renda_faixa'],
        VAR['col_gasto_alimentacao'],
        VAR['col_pct_respostas_rapidas'],
        VAR['col_tentativas_completa'],
        VAR['col_lat_referencia'],
        VAR['col_long_referencia'],
        VAR['col_nome'],
        VAR['col_telefone'],
    ]
    df = df[[c for c in base_cols if c in df.columns] + visit_cols]

    # Renomeia variáveis do projeto para nomes internos padronizados
    rename_map = {
        VAR['col_entrevistador']:          'Srvyr',
        VAR['col_autorizou_recontato']:    'Autorizou_Recontato',
        VAR['col_duracao']:                'Duracao',
        VAR['col_audio_gravado']:          'Audio_Gravado',
        VAR['col_psu']:                    'PSU',
        VAR['col_versao_questionario']:    'Versão do questionário',
        VAR['col_pais']:                   'Country',
        VAR['col_idade_selecao']:          'Kish_age',
        VAR['col_idade_questionario']:     'Ent_age',
        VAR['col_idade_questionario_dup']: 'Ent_age_DUB',
        VAR['col_autorizacao_menor']:      'Autor_menor',
        VAR['col_sexo_selecao']:           'Kish_gender',
        VAR['col_sexo_questionario']:      'Ent_Gender',
        VAR['col_membros_selecao']:        'Kish_membros',
        VAR['col_membros_questionario']:   'Ent_membros',
        VAR['col_membros_menos15']:        'Membros_menos15anos',
        VAR['col_anos_estudo']:            'Years_educ',
        VAR['col_nivel_escolaridade']:     'Level_educ',
        VAR['col_renda']:                  'Renda',
        VAR['col_renda_faixa']:            'WP19188',
        VAR['col_gasto_alimentacao']:      'Gasto_alimentacao',
        VAR['col_pct_respostas_rapidas']:  'Porcentagem_menos3s',
        VAR['col_tentativas_completa']:    'Tentativas_por_completa',
        VAR['col_lat_referencia']:         'LAT_FORNECIDO',
        VAR['col_long_referencia']:        'LONG_FORNECIDO',
        VAR['col_nome']:                   'Nome',
        VAR['col_telefone']:               'Telefone',
        'LOC_START_LA':   'LOC_START_LAT',
        'LOC_START_LO':   'LOC_START_LONG',
        'LOC_MAIN_LA':    'LOC_MAIN_LAT',
        'LOC_MAIN_LO':    'LOC_MAIN_LONG',
        'LOC_WP12_LA':    'HH_LOC_LAT',
        'LOC_WP12_LO':    'HH_LOC_LONG',
        'LOC_MAIN_END_LA':'LOC_MAIN_END_LAT',
        'LOC_MAIN_END_LO':'LOC_MAIN_END_LONG',
    }
    df = df.rename(columns=rename_map)

    if 'Srvyr' in df.columns:
        df['Srvyr_original'] = df['Srvyr']
        df['Srvyr'] = df['Srvyr'].astype(str).apply(lambda x: re.sub('[^0-9]', '', x))
        df['Srvyr'] = pd.to_numeric(df['Srvyr'], errors='coerce').astype('Int64')
    else:
        df['Srvyr_original'] = pd.NA
        df['Srvyr'] = pd.NA

    if 'Srvyr' in df.columns and pd.api.types.is_numeric_dtype(df['Srvyr']):
        df.insert(4, "Fornecedor", pd.cut(
            x=df['Srvyr'], bins=intervalos_logins,
            labels=fornecedores, include_lowest=True, right=True
        ), allow_duplicates=False)
    else:
        df.insert(4, "Fornecedor", pd.NA, allow_duplicates=False)

    for col, mapa in [('Autorizou_Recontato', codigo_respostas), ('Audio_Gravado', codigo_respostas)]:
        if col in df.columns:
            df[col] = df[col].map(mapa, na_action='ignore')

    df.insert(5, "Data_envio_volta", data_hoje)
    df['Endereco'] = None
    df['PSU_tratada'] = df['PSU'].astype(str).str[:-3] if 'PSU' in df.columns else pd.NA
    df['Flagsbyscript'] = df['FlagsText'].notnull() if 'FlagsText' in df.columns else False

    for col in colunas_gps + ['LAT_FORNECIDO', 'LONG_FORNECIDO']:
        if col not in df.columns:
            df[col] = np.nan

    df['LAT_FORNECIDO']  = df['LAT_FORNECIDO'].apply(transformar_float)
    df['LONG_FORNECIDO'] = df['LONG_FORNECIDO'].apply(transformar_float)

    for sufixo, lat2, lon2 in [
        ('LOC_START',       'LOC_START_LAT',    'LOC_START_LONG'),
        ('LOC_MAIN',        'LOC_MAIN_LAT',     'LOC_MAIN_LONG'),
        ('HH_LOC',          'HH_LOC_LAT',       'HH_LOC_LONG'),
        ('LOC_MAIN_END',    'LOC_MAIN_END_LAT', 'LOC_MAIN_END_LONG'),
    ]:
        df[f'DIST_{sufixo}'] = df.apply(
            lambda x: distancia2d(x['LAT_FORNECIDO'], x['LONG_FORNECIDO'], x[lat2], x[lon2]), axis=1
        )

    df['DIST_LOC_MAIN_LOC_MAIN_END'] = df.apply(
        lambda x: distancia2d(x['LOC_MAIN_LAT'], x['LOC_MAIN_LONG'], x['LOC_MAIN_END_LAT'], x['LOC_MAIN_END_LONG']), axis=1
    )
    df['Possui_GPS'] = ~df[colunas_gps].isnull().any(axis=1)

    return df.reset_index(drop=True)

def selecionar_colunas_geral(df):
    """Seleciona colunas essenciais para lookups e relatórios."""
    cols = ["Status", "DeviceIndex", "SbjNum", "Srvyr_original", "Fornecedor", "PSU_tratada", "VEnd"]
    return df.loc[:, [c for c in cols if c in df.columns]].copy()


# ─────────────────────────────────────────
# PROCESSAMENTO PRINCIPAL
# ─────────────────────────────────────────

Base = process_dataframe_base(Base, data_hoje)
Base = Base.query("Status not in ['Canceled', 'Expired']").reset_index(drop=True)

Base['Renda_num']      = Base['Renda'].apply(to_number_or_nan)
Base['Gasto_alim_num'] = Base['Gasto_alimentacao'].apply(to_number_or_nan)

if "WP19188" in Base.columns:
    Base["WP19188_code"] = Base["WP19188"].apply(lambda v: to_number_or_nan(v, missing_codes=(98, 99)))
    Base["WP19188_num"]  = Base["WP19188_code"].apply(
        lambda v: wp19188_midpoint.get(int(v), np.nan) if pd.notna(v) else np.nan
    )
else:
    Base["WP19188_num"] = np.nan

Base["Renda_efetiva"] = np.where(
    Base["Renda_num"].notna(), Base["Renda_num"],
    np.where(Base["WP19188_num"].notna(), Base["WP19188_num"], np.nan)
)
Base["Fonte_renda"] = np.where(
    Base["Renda_num"].notna(), "WP7133",
    np.where(Base["WP19188_num"].notna(), "WP19188 (faixa estimada)", "Não informada")
)

BaseGeral_processed = process_dataframe_base(BaseGeral, data_hoje)
BaseGeral_full      = BaseGeral_processed.query("Status not in ['Canceled', 'Expired']").reset_index(drop=True)

Limite_Duracao = Base['Duracao'].mean() - (1.5 * Base['Duracao'].std())

print(f"Base do dia {data_hoje}: {Base.shape[0]} entrevistas")
print(f"Duração média: {round(Base['Duracao'].mean(), 2)} minutos")
print(f"Limite mínimo de duração: {round(Limite_Duracao, 2)} minutos")


# ─────────────────────────────────────────
# IDENTIFICAÇÃO DE INCONSISTÊNCIAS
# ─────────────────────────────────────────

escolaridade_mask  = Base.apply(check_escolaridade, axis=1)
escolaridade_lista = Base.index[escolaridade_mask].tolist()

membros_mask  = Base.apply(check_membros, axis=1)
membros_lista = Base.index[membros_mask].tolist()

idade_lista  = Base.index[abs(Base['Kish_age'] - Base['Ent_age']) > 2].tolist()
sexo_lista   = Base.index[Base['Kish_gender'] != Base['Ent_Gender']].tolist()
menor_lista  = Base.index[(Base['Kish_age'] < 18) & (Base['Autor_menor'] != 1)].tolist()

renda_lista = Base.index[
    Base['Renda'].notna() &
    ((Base['Renda'] < 500) | (Base['Renda'] > 50000)) &
    ~Base['Renda'].isin([98, 99])
].tolist()

duracao_lista       = Base.index[Base['Duracao'] < 19].tolist()
duracao_longa_lista = Base.index[Base['Duracao'] > 70].tolist()
flags_lista         = Base.index[Base['Flagsbyscript'] == True].tolist()
parcial_lista       = Base.index[Base['Porcentagem_menos3s'] >= 20].tolist()
membros15_lista     = Base.index[Base['Membros_menos15anos'] > 6].tolist()

gasto_lista = Base.index[
    (Base['Gasto_alim_num'].notna() & Base['Renda_efetiva'].notna() & (Base['Gasto_alim_num'] > Base['Renda_efetiva'])) |
    (Base['Gasto_alim_num'].notna() & Base['Renda_efetiva'].notna() & (Base['Gasto_alim_num'] == 0) & (Base['Renda_efetiva'] > 0)) |
    (Base['Gasto_alim_num'].notna() & (Base['Gasto_alim_num'] > 0) & Base['Renda_efetiva'].isna())
].tolist()

# Visitas NEC
nec_lista, nec_info = [], {}
for idx, row in Base.iterrows():
    flag, data_v, qtd = get_nec_3plus_same_day_info(row)
    if flag:
        nec_lista.append(idx)
        nec_info[idx] = (data_v, qtd)

# Intervalo entre tentativas
intervalo_lista, intervalo_info = [], {}
for idx, row in Base.iterrows():
    flag, ia, ib, mins, ca, cb = get_intervalo_tentativas_info(row)
    if flag:
        intervalo_lista.append(idx)
        intervalo_info[idx] = (ia, ib, mins, ca, cb)

# Alta taxa de primeira tentativa por PSU
alta_psu_lista, alta_psu_info = [], {}
completed = Base.query("Status == 'Completed'").copy()
if not completed.empty:
    total_psu = completed.groupby('PSU_tratada').size().rename('Total')
    first_psu = completed.query("Tentativas_por_completa == 1").groupby('PSU_tratada').size().rename('First')
    stats = pd.merge(total_psu, first_psu, left_index=True, right_index=True, how='left').fillna(0)
    stats['Pct'] = (stats['First'] / stats['Total'] * 100).round(1)
    for psu, row in stats[(stats['Total'] == 10) & (stats['Pct'] > 50)].iterrows():
        for idx in Base.index[Base['PSU_tratada'] == psu]:
            if idx not in alta_psu_lista:
                alta_psu_lista.append(idx)
                alta_psu_info[idx] = (psu, row['Pct'])

# 10+ entrevistas por dia por entrevistador
dia_lista, dia_info = [], {}
if 'VEnd' in Base.columns and 'Srvyr_original' in Base.columns:
    tmp = Base.copy()
    tmp['VEnd_date'] = pd.to_datetime(tmp['VEnd'], errors='coerce').dt.date
    tmp = tmp.dropna(subset=['VEnd_date', 'Srvyr_original'])
    if not tmp.empty:
        contagem = tmp.groupby(['Srvyr_original', 'VEnd_date']).size().reset_index(name='Count')
        for _, r in contagem[contagem['Count'] >= 10].iterrows():
            for idx in Base.index[
                (pd.to_datetime(Base['VEnd'], errors='coerce').dt.date == r['VEnd_date']) &
                (Base['Srvyr_original'] == r['Srvyr_original'])
            ]:
                if idx not in dia_lista:
                    dia_lista.append(idx)
                    dia_info[idx] = (r['Srvyr_original'], r['VEnd_date'].strftime('%d/%m/%Y'), r['Count'])

# Baixa autorização de áudio por PSU
audio_lista, audio_info = [], {}
comp_psu = Base.query("Status == 'Completed' and PSU_tratada.notna()").copy()
if not comp_psu.empty:
    for psu, group in comp_psu.groupby('PSU_tratada'):
        if len(group) == 10:
            pct = group.query("Audio_Gravado == 'Sim'").shape[0] / len(group) * 100
            if pct < 50:
                for idx in group.index:
                    if idx not in audio_lista:
                        audio_lista.append(idx)
                        audio_info[idx] = (psu, pct)

# Entrevistas em horário noturno
noturno_lista, noturno_info = [], {}
if 'VEnd' in Base.columns:
    Base['VEnd_datetime'] = pd.to_datetime(Base['VEnd'], errors='coerce')
    for idx, row in Base.query("Status == 'Completed'").iterrows():
        if check_horario_noturno(row['VEnd_datetime']):
            noturno_lista.append(idx)
            noturno_info[idx] = row['VEnd_datetime']


# ─────────────────────────────────────────
# CONTAGEM E RESUMO
# ─────────────────────────────────────────

categorias = [
    ("Membros do domicílio",             membros_lista),
    ("Idade",                            idade_lista),
    ("Sexo",                             sexo_lista),
    ("Escolaridade",                     escolaridade_lista),
    ("Menores sem autorização",          menor_lista),
    ("Renda fora do intervalo",          renda_lista),
    ("Duração abaixo do mínimo",         duracao_lista),
    ("Duração acima de 1h10",            duracao_longa_lista),
    ("Flags do sistema",                 flags_lista),
    ("Velocidade de resposta (parcial)", parcial_lista),
    ("Moradores < 15 anos",             membros15_lista),
    ("Gasto alimentação vs renda",       gasto_lista),
    ("3+ visitas NEC no mesmo dia",      nec_lista),
    ("Intervalo entre tentativas < 2h",  intervalo_lista),
    ("Alta taxa 1ª tentativa por PSU",   alta_psu_lista),
    ("10+ entrevistas no mesmo dia",     dia_lista),
    ("Baixa autorização de áudio",       audio_lista),
    ("Entrevistas em horário noturno",   noturno_lista),
]

total_inconsistencias = sum(len(lst) for _, lst in categorias)
print(f"\nTotal de inconsistências encontradas: {total_inconsistencias}\n")
for nome, lst in categorias:
    print(f"{nome}: {len(lst)}")


# ─────────────────────────────────────────
# CONSTRUÇÃO DOS DataFrames DE PROBLEMAS
# ─────────────────────────────────────────

def build_df(lista, desc_func):
    """Constrói DataFrame de inconsistências com coluna 'Problema'."""
    frames = []
    for i in lista:
        linha = Base.iloc[[i]].copy()
        linha['Problema'] = desc_func(Base.iloc[i])
        frames.append(linha)
    return pd.concat(frames) if frames else pd.DataFrame()

membros_df      = build_df(membros_lista,    lambda r: membros_desc(r['Kish_membros'], r['Ent_membros']))
idade_df        = build_df(idade_lista,      lambda r: idade_desc(r['Kish_age'], r['Ent_age']))
sexo_df         = build_df(sexo_lista,       lambda r: sexo_desc(r['Kish_gender']))
escolaridade_df = build_df(escolaridade_lista, lambda r: escolaridade_desc(r['Years_educ'], r['Level_educ']))
renda_df        = build_df(renda_lista,      lambda r: renda_desc(r['Renda']))
menor_df        = build_df(menor_lista,      lambda r: menor_desc(r['Kish_age']))
duracao_df      = build_df(duracao_lista,    lambda r: duracao_desc(r['Duracao']))
duracao_lg_df   = build_df(duracao_longa_lista, lambda r: duracao_longa_desc(r['Duracao']))
membros15_df    = build_df(membros15_lista,  lambda r: Membros_menos15anos_desc(r['Membros_menos15anos']))
parcial_df      = build_df(parcial_lista,    lambda r: parcial_desc(r['Porcentagem_menos3s']))
gasto_df        = build_df(gasto_lista,      lambda r: gasto_alimentacao_desc(r['Gasto_alim_num'], r['Renda_efetiva'], r['Fonte_renda']))

# Flags do sistema
flags_df = pd.DataFrame()
if flags_lista:
    flags_df = pd.concat([Base.iloc[[i]] for i in flags_lista]).copy()
    flags_df['Problema'] = flags_df['FlagsText']

# Inconsistências com contexto específico
def build_df_info(lista, info_dict, desc_func):
    frames = []
    for idx in lista:
        linha = Base.iloc[[idx]].copy()
        linha['Problema'] = desc_func(*info_dict[idx])
        frames.append(linha)
    return pd.concat(frames) if frames else pd.DataFrame()

nec_df        = build_df_info(nec_lista,       nec_info,       visitas_nec_desc)
intervalo_df  = build_df_info(intervalo_lista,  intervalo_info, intervalo_tentativas_desc)
alta_psu_df   = build_df_info(alta_psu_lista,   alta_psu_info,  alta_primeira_tentativa_psu_desc)
dia_df        = build_df_info(dia_lista,         dia_info,       entrevistas_dia_desc)
audio_df      = build_df_info(audio_lista,       audio_info,     baixa_audio_psu_desc)
noturno_df    = build_df_info(noturno_lista,     noturno_info,   horario_noturno_desc)


# ─────────────────────────────────────────
# CONSOLIDAÇÃO E SALVAMENTO
# ─────────────────────────────────────────

todas = pd.concat([
    membros_df, idade_df, sexo_df, escolaridade_df, renda_df,
    menor_df, duracao_df, duracao_lg_df, flags_df, membros15_df,
    parcial_df, gasto_df, nec_df, intervalo_df, alta_psu_df,
    dia_df, audio_df, noturno_df,
], ignore_index=True, sort=False)

confirmar(todas, total_inconsistencias)

if len(todas) > 0:
    colunas_voltas = [
        'SbjNum', 'DeviceIndex', 'Upload', 'Srvyr', 'Fornecedor',
        'Data_envio_volta', 'Status', 'Autorizou_Recontato', 'Audio_Gravado',
        'PSU', 'Nome', 'Telefone', 'Problema'
    ]
    BaseVoltas = todas[[c for c in colunas_voltas if c in todas.columns]].copy()
    BaseVoltas['Tipo de volta']  = None
    BaseVoltas['Justificativa']  = None
    BaseVoltas['Resposta']       = None
    salvar_excel(BaseVoltas, f"Voltas{data_diames}")
else:
    print(f"Nenhuma inconsistência encontrada na base do dia {data_hoje}.")


# ─────────────────────────────────────────
# BASE GERAL E RELATÓRIO DE HORÁRIOS
# ─────────────────────────────────────────

BaseGeral_proc = selecionar_colunas_geral(BaseGeral_processed)
BaseCanceladas = (
    BaseGeral_processed.query("Status == 'Canceled'").copy()
    if "Status" in BaseGeral_processed.columns
    else pd.DataFrame()
)

tentativas_geral = extrair_tentativas(BaseGeral_processed, col_srvyr="Srvyr_original")
relatorio_global, relatorio_srvyr = calcular_relatorio_horarios(tentativas_geral)

print("\n--- Resumo Global de Horários ---")
print(relatorio_global.to_string(index=False))
print("\n--- Por Entrevistador ---")
print(relatorio_srvyr.to_string(index=False))

srvyr_baixo = relatorio_srvyr[relatorio_srvyr["Abaixo_de_20%"] == "Sim"]["Srvyr"].tolist()
if srvyr_baixo:
    print(f"\nEntrevistadores abaixo de 20% de tentativas após 17h ou FDS: {srvyr_baixo}")
else:
    print("\nTodos os entrevistadores atingiram o mínimo de 20%.")


# ─────────────────────────────────────────
# DICIONÁRIOS DE LOOKUP
# ─────────────────────────────────────────

DeviceIndex_Srvyr = (
    BaseGeral_proc
    .loc[BaseGeral_proc["DeviceIndex"].notna(), ["DeviceIndex", "Srvyr_original"]]
    .dropna(subset=["Srvyr_original"])
    .assign(Srvyr_original=lambda d: as_str_clean(d["Srvyr_original"]))
    .dropna(subset=["Srvyr_original"])
    .groupby("DeviceIndex")["Srvyr_original"]
    .apply(lambda s: sorted(pd.unique(s.dropna()).tolist()))
    .to_dict()
) if {"DeviceIndex", "Srvyr_original"}.issubset(BaseGeral_proc.columns) else {}

SbjNum_DeviceIndex = (
    BaseGeral_proc
    .loc[BaseGeral_proc["SbjNum"].notna(), ["SbjNum", "DeviceIndex"]]
    .dropna(subset=["DeviceIndex"])
    .groupby("SbjNum")["DeviceIndex"]
    .apply(lambda s: sorted(pd.unique(as_str_clean(s).dropna()).tolist()))
    .to_dict()
) if {"SbjNum", "DeviceIndex"}.issubset(BaseGeral_proc.columns) else {}

srvyr_fornecedor = (
    BaseGeral_proc[["Srvyr_original", "Fornecedor"]]
    .assign(
        Srvyr_original=lambda d: as_str_clean(d["Srvyr_original"]),
        Fornecedor=lambda d: as_str_clean(d["Fornecedor"]),
    )
    .dropna(subset=["Srvyr_original", "Fornecedor"])
    .drop_duplicates()
    .reset_index(drop=True)
) if {"Srvyr_original", "Fornecedor"}.issubset(BaseGeral_proc.columns) else pd.DataFrame()


# ─────────────────────────────────────────
# SALVAMENTO FINAL
# ─────────────────────────────────────────

colunas_basegeral_final = [
    'SbjNum', 'DeviceIndex', 'Upload', 'VEnd', 'Status',
    'Srvyr_original', 'Fornecedor', 'PSU_tratada',
    'Autorizou_Recontato', 'Duracao', 'Audio_Gravado',
    'Kish_age', 'Ent_age', 'Ent_age_DUB', 'Autor_menor',
    'Kish_gender', 'Ent_Gender', 'Kish_membros', 'Ent_membros',
    'LAT_FORNECIDO', 'LONG_FORNECIDO',
    'LOC_START_LAT', 'LOC_START_LONG', 'LOC_MAIN_LAT', 'LOC_MAIN_LONG',
    'HH_LOC_LAT', 'HH_LOC_LONG', 'LOC_MAIN_END_LAT', 'LOC_MAIN_END_LONG',
    'DIST_SP', 'Tentativas_por_completa', 'LocName', 'Nome', 'Telefone'
]

basegeral_para_salvar = BaseGeral_full.reindex(
    columns=[c for c in colunas_basegeral_final if c in BaseGeral_full.columns]
)

def salvar_com_retry(df, caminho, tentativas=3):
    for i in range(tentativas):
        try:
            df.to_excel(caminho, index=False)
            print(f"Salvo: {caminho}")
            return
        except (ConnectionAbortedError, OSError) as e:
            print(f"Tentativa {i+1} falhou: {e}. Aguardando 5s...")
            time.sleep(5)
    print(f"Erro: não foi possível salvar {caminho}.")

for df, nome in [
    (basegeral_para_salvar, "BaseGeral"),
    (BaseCanceladas,        "BaseCanceladas"),
]:
    if df is not None and len(df) > 0:
        salvar_com_retry(df, OUTPUT_DIR / f"{nome}.xlsx")
    else:
        print(f"{nome}: vazia, não salva.")

# Relatório de supervisão
caminho_sup = OUTPUT_DIR / "Relatorio_Horarios.xlsx"
with pd.ExcelWriter(caminho_sup, engine="openpyxl") as writer:
    if len(relatorio_global) > 0:
        relatorio_global.to_excel(writer, sheet_name="Global", index=False)
    if len(relatorio_srvyr) > 0:
        relatorio_srvyr.to_excel(writer, sheet_name="Por_Entrevistador", index=False)
print(f"Relatório de horários salvo em: {caminho_sup}")


# ─────────────────────────────────────────
# RELATÓRIO DE RESULTADOS DE TENTATIVAS
# ─────────────────────────────────────────

todas_tentativas = extrair_tentativas_com_outros(BaseGeral_processed, col_srvyr="Srvyr_original")

# Tentativas com motivo "Outros" (código 11)
outros_df = todas_tentativas[
    (todas_tentativas['VisitCode'] == 11) & (todas_tentativas['Other_Reason'].notna())
].copy()

if not outros_df.empty:
    outros_df = outros_df[[
        'Srvyr', 'PSU_tratada', 'SbjNum', 'DeviceIndex',
        'VisitDateTime', 'VisitCode_Desc', 'Other_Reason'
    ]].rename(columns={
        'Srvyr': 'Entrevistador', 'PSU_tratada': 'PSU',
        'VisitDateTime': 'Data_Tentativa', 'VisitCode_Desc': 'Codigo_Visita',
        'Other_Reason': 'Motivo'
    })
    outros_df['Data_Tentativa'] = outros_df['Data_Tentativa'].dt.strftime('%d/%m/%Y %H:%M:%S')

# Frequência de códigos por entrevistador
freq_df = todas_tentativas.dropna(subset=['Srvyr', 'VisitCode']).copy()
freq_df['Srvyr'] = freq_df['Srvyr'].astype(str).replace('nan', 'Não Informado')

if not freq_df.empty:
    freq_pivot = pd.pivot_table(
        freq_df, values='VisitIndex', index='Srvyr',
        columns='VisitCode_Desc', aggfunc='count', fill_value=0
    ).reset_index().rename(columns={'Srvyr': 'Entrevistador'})
    for desc in codigo_visita_desc.values():
        if desc not in freq_pivot.columns:
            freq_pivot[desc] = 0
    cols_ord = ['Entrevistador'] + [codigo_visita_desc[k] for k in sorted(codigo_visita_desc)]
    freq_pivot = freq_pivot.reindex(columns=cols_ord, fill_value=0)

caminho_tent = OUTPUT_DIR / "Resultados_Tentativas.xlsx"
with pd.ExcelWriter(caminho_tent, engine="openpyxl") as writer:
    if not outros_df.empty:
        outros_df.to_excel(writer, sheet_name="Outros", index=False)
    if not freq_df.empty:
        freq_pivot.to_excel(writer, sheet_name="Frequencia", index=False)
print(f"Resultados de tentativas salvos em: {caminho_tent}")
