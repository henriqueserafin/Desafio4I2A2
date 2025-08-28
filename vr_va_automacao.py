#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Automação da Compra de VR/VA

- Consolida e trata bases (.xlsx)
- Aplica exclusões (estagiários, aprendizes, afastados, exterior, diretores)
- Calcula dias e valores conforme sindicato, férias e desligamento
- Gera arquivo final no layout exato:
  Matricula, Admissão, Sindicato do Colaborador, Competência, Dias,
  VALOR DIÁRIO VR, TOTAL, Custo empresa, Desconto profissional, OBS GERAL

Uso:
    python vr_va_automacao.py --competencia 2025-05 --saida VR_FINAL.xlsx

Se 'saida' não for informado, será salvo como: VR_FINAL_YYYYMMDD_HHMMSS.xlsx
"""

import argparse
from pathlib import Path
from datetime import datetime
import pandas as pd
import numpy as np

# Cabeçalho final EXATO desejado (não depende de arquivo modelo)
HEADERS = [
    "Matricula",
    "Admissão",
    "Sindicato do Colaborador",
    "Competência",
    "Dias",
    "VALOR DIÁRIO VR",
    "TOTAL",
    "Custo empresa",
    "Desconto profissional",
    "OBS GERAL",
]

# Fallbacks padrão
DIA_UTEIS_PADRAO = 22
VALOR_DIARIO_FALLBACK = 35.0


def detect_dirs():
    base_dir = Path(__file__).resolve().parent if "__file__" in globals() else Path().resolve()
    dados_dir = base_dir / "Dados"
    uploads_dir = base_dir / "Uploads"
    return base_dir, dados_dir, uploads_dir


def read_excel_any(name: str, base_dir: Path, dados_dir: Path, uploads_dir: Path) -> pd.DataFrame:
    for folder in [base_dir, dados_dir, uploads_dir]:
        path = folder / name
        if path.exists():
            try:
                return pd.read_excel(path)
            except Exception as e:
                print(f"✖ Erro ao ler {path.name}: {e}")
                return pd.DataFrame()
    print(f"✖ Arquivo não encontrado: {name}")
    return pd.DataFrame()


def standardize_matricula(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    for col in list(df.columns):
        low = col.lower()
        if "matric" in low or "cadastro" in low:
            if col != "MATRICULA":
                df.rename(columns={col: "MATRICULA"}, inplace=True)
    if "MATRICULA" in df.columns:
        df["MATRICULA"] = pd.to_numeric(df["MATRICULA"], errors="coerce").astype("Int64")
    return df


def standardize_sindicato(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    candidates = [c for c in df.columns if "sind" in c.lower()]
    if candidates and candidates[0] != "Sindicato":
        df.rename(columns={candidates[0]: "Sindicato"}, inplace=True)
    return df


def build_exclusion_set(ativos: pd.DataFrame,
                        estagio: pd.DataFrame,
                        aprendiz: pd.DataFrame,
                        afast: pd.DataFrame,
                        exterior: pd.DataFrame) -> set:
    excluir = set()
    # Estagiários, aprendizes, afastados
    for df in [estagio, aprendiz, afast]:
        if not df.empty and "MATRICULA" in df.columns:
            excluir.update(df["MATRICULA"].dropna().unique())
    # Exterior (coluna 'Cadastro' na base)
    if not exterior.empty:
        if "MATRICULA" not in exterior.columns and "Cadastro" in exterior.columns:
            exterior = exterior.rename(columns={"Cadastro": "MATRICULA"})
        if "MATRICULA" in exterior.columns:
            excluir.update(pd.to_numeric(exterior["MATRICULA"], errors="coerce").dropna().astype("Int64").unique())

    # Diretores (cargo contém 'DIRETOR')
    if not ativos.empty and "TITULO DO CARGO" in ativos.columns:
        mask = ativos["TITULO DO CARGO"].astype(str).str.contains("DIRETOR", case=False, na=False)
        diretores = pd.to_numeric(ativos.loc[mask, "MATRICULA"], errors="coerce").dropna().astype("Int64").unique()
        excluir.update(diretores)

    print(f"🚫 Total de matrículas excluídas: {len(excluir)}")
    return excluir


def map_dias_uteis(dias_uteis_df: pd.DataFrame) -> dict:
    """
    Base 'Basediasuteis.xlsx' costuma ter primeira linha com cabeçalho textual.
    Detecta a coluna de texto (sindicato) e a numérica (dias) de forma leniente.
    """
    out = {}
    if dias_uteis_df.empty:
        return out

    # Escolhe col de sindicato e col de dias
    col_sind = dias_uteis_df.columns[0]
    col_dias = dias_uteis_df.columns[1] if len(dias_uteis_df.columns) > 1 else dias_uteis_df.columns[0]

    for _, row in dias_uteis_df.iterrows():
        s = str(row.get(col_sind, "")).strip()
        d = row.get(col_dias, None)
        if not s or "SINDIC" in s.upper() or "DIAS" in str(d).upper():
            continue
        try:
            d_int = int(d)
        except Exception:
            continue
        out[s] = d_int

    print(f"📅 Mapeamentos de dias úteis: {len(out)}")
    return out


def map_valor_por_estado(sind_val_df: pd.DataFrame) -> dict:
    out = {}
    if sind_val_df.empty:
        return out

    # Detecta colunas "ESTADO" e "VALOR" de forma leniente
    c_estado = next((c for c in sind_val_df.columns if "ESTADO" in c.upper()), sind_val_df.columns[0])
    c_valor = next((c for c in sind_val_df.columns if "VALOR" in c.upper()),
                   sind_val_df.columns[1] if len(sind_val_df.columns) > 1 else sind_val_df.columns[0])

    for _, row in sind_val_df.iterrows():
        est = str(row.get(c_estado, "")).strip()
        val = row.get(c_valor, None)
        if not est:
            continue
        try:
            out[est] = float(val)
        except Exception:
            pass

    print(f"💵 Mapeamentos de valor por estado: {len(out)}")
    return out


def find_dias_uteis(sindicato: str, dias_map: dict) -> int:
    s = str(sindicato)
    for key, v in dias_map.items():
        if key in s:
            return v
    return DIA_UTEIS_PADRAO


def find_valor_diario(sindicato: str, valor_map: dict) -> float:
    s = str(sindicato).upper()
    if "SÃO PAULO" in s or "SAO PAULO" in s or "SP" in s:
        return valor_map.get("São Paulo", 37.5)
    if "RIO DE JANEIRO" in s or "RJ" in s:
        return valor_map.get("Rio de Janeiro", 35.0)
    if "RIO GRANDE DO SUL" in s or "RS" in s:
        return valor_map.get("Rio Grande do Sul", 35.0)
    if "PARANÁ" in s or "PARANA" in s or "PR" in s or "CURITIBA" in s:
        return valor_map.get("Paraná", 35.0)
    return valor_map.get("DEFAULT", VALOR_DIARIO_FALLBACK)


def parse_args():
    ap = argparse.ArgumentParser(description="Automação da Compra de VR/VA")
    ap.add_argument("--competencia", type=str, default="2025-05",
                    help="Competência no formato YYYY-MM (padrão 2025-05)")
    ap.add_argument("--saida", type=str, default="",
                    help="Caminho do XLSX de saída. Se vazio, será gerado VR_FINAL_YYYYMMDD_HHMMSS.xlsx")
    return ap.parse_args()


def main():
    args = parse_args()
    base_dir, dados_dir, uploads_dir = detect_dirs()

    # Competência: primeiro dia do mês solicitado
    try:
        comp = pd.to_datetime(f"{args.competencia}-01")
    except Exception:
        comp = pd.to_datetime("2025-05-01")
    print(f"🗓 Competência: {comp.strftime('%Y-%m-%d')}")

    # Carregar bases
    ativos = read_excel_any("ATIVOS.xlsx", base_dir, dados_dir, uploads_dir)
    ferias = read_excel_any("FERIAS.xlsx", base_dir, dados_dir, uploads_dir)
    desligados = read_excel_any("DESLIGADOS.xlsx", base_dir, dados_dir, uploads_dir)
    admissoes = read_excel_any("ADMISSOABRIL.xlsx", base_dir, dados_dir, uploads_dir)
    sindicato_valor = read_excel_any("Basesindicatoxvalor.xlsx", base_dir, dados_dir, uploads_dir)
    dias_uteis = read_excel_any("Basediasuteis.xlsx", base_dir, dados_dir, uploads_dir)
    afast = read_excel_any("AFASTAMENTOS.xlsx", base_dir, dados_dir, uploads_dir)
    estagio = read_excel_any("ESTAGIO.xlsx", base_dir, dados_dir, uploads_dir)
    aprendiz = read_excel_any("APRENDIZ.xlsx", base_dir, dados_dir, uploads_dir)
    exterior = read_excel_any("EXTERIOR.xlsx", base_dir, dados_dir, uploads_dir)

    # Padronizações
    for df in [ativos, ferias, desligados, admissoes, afast, estagio, aprendiz, exterior]:
        standardize_matricula(df)
    standardize_sindicato(ativos)

    # Exclusões
    excluir = build_exclusion_set(ativos, estagio, aprendiz, afast, exterior)

    # Base inicial
    if ativos.empty or "MATRICULA" not in ativos.columns:
        raise RuntimeError("Base ATIVOS.xlsx vazia ou sem coluna MATRICULA.")

    base = ativos[~ativos["MATRICULA"].isin(excluir)].copy()
    print(f"📊 Base após exclusões: {len(base)} registros")

    # Merge férias
    if not ferias.empty and "DIAS DE FÉRIAS" in ferias.columns:
        base = base.merge(ferias[["MATRICULA", "DIAS DE FÉRIAS"]], on="MATRICULA", how="left")

    # Merge desligados
    cols_desl = []
    if "DATA DEMISSÃO" in desligados.columns:
        cols_desl.append("DATA DEMISSÃO")
    if "COMUNICADO DE DESLIGAMENTO" in desligados.columns:
        cols_desl.append("COMUNICADO DE DESLIGAMENTO")
    if not desligados.empty and cols_desl:
        base = base.merge(desligados[["MATRICULA"] + cols_desl], on="MATRICULA", how="left")

    # Merge admissão
    for col in [c for c in admissoes.columns if "admiss" in c.lower()]:
        admissoes.rename(columns={col: "Admissão"}, inplace=True)
    if not admissoes.empty and "Admissão" in admissoes.columns:
        base = base.merge(admissoes[["MATRICULA", "Admissão"]], on="MATRICULA", how="left")

    # Mapas auxiliares
    dias_map = map_dias_uteis(dias_uteis)
    valor_map = map_valor_por_estado(sindicato_valor)

    # Cálculo linha a linha
    def calcula_linha(row: pd.Series) -> pd.Series:
        dias_base = find_dias_uteis(row.get("Sindicato", ""), dias_map)
        dias = dias_base
        obs = []

        # Férias
        if pd.notna(row.get("DIAS DE FÉRIAS")):
            try:
                dfer = int(row["DIAS DE FÉRIAS"])
                dias -= dfer
                obs.append(f"Férias: -{dfer}")
            except Exception:
                pass

        # Regra de desligamento
        if pd.notna(row.get("DATA DEMISSÃO")):
            try:
                data_desl = pd.to_datetime(row["DATA DEMISSÃO"])
                comunicado = str(row.get("COMUNICADO DE DESLIGAMENTO", "")).strip().upper()
                if data_desl.year == comp.year and data_desl.month == comp.month:
                    if data_desl.day <= 15 and comunicado == "OK":
                        dias = 0
                        obs.append("Desligado até dia 15 - sem benefício")
                    elif data_desl.day > 15:
                        dias_prop = int(dias_base * (data_desl.day / 30))
                        dias = min(dias, dias_prop)
                        obs.append(f"Desligado dia {data_desl.day} - proporcional")
            except Exception:
                pass

        # Opcional: Admissão no mês (proporcional)
        if pd.notna(row.get("Admissão")):
            try:
                data_adm = pd.to_datetime(row["Admissão"])
                if data_adm.year == comp.year and data_adm.month == comp.month:
                    dias_prop_adm = int(dias_base * ((30 - (data_adm.day - 1)) / 30))
                    dias = min(dias, max(0, dias_prop_adm))
                    obs.append(f"Admissão dia {data_adm.day} - proporcional")
            except Exception:
                pass

        dias = max(0, int(dias))
        valor_diario = float(find_valor_diario(row.get("Sindicato", ""), valor_map))
        total = round(dias * valor_diario, 2)
        custo_emp = round(total * 0.80, 2)
        desc_col = round(total * 0.20, 2)

        return pd.Series({
            "Dias": dias,
            "VALOR DIÁRIO VR": valor_diario,
            "TOTAL": total,
            "Custo empresa": custo_emp,
            "Desconto profissional": desc_col,
            "OBS GERAL": "; ".join(obs)
        })

    calc = base.apply(calcula_linha, axis=1)
    base_calc = pd.concat([base, calc], axis=1)

    # DataFrame final conforme layout
    out = pd.DataFrame({
        "Matricula": base_calc["MATRICULA"].astype("Int64"),
        "Admissão": base_calc.get("Admissão", pd.NaT),
        "Sindicato do Colaborador": base_calc["Sindicato"].astype(str),
        "Competência": pd.to_datetime(comp),
        "Dias": base_calc["Dias"].astype(int),
        "VALOR DIÁRIO VR": base_calc["VALOR DIÁRIO VR"].astype(float),
        "TOTAL": base_calc["TOTAL"].astype(float),
        "Custo empresa": base_calc["Custo empresa"].astype(float),
        "Desconto profissional": base_calc["Desconto profissional"].astype(float),
        "OBS GERAL": base_calc["OBS GERAL"].fillna("").astype(str),
    })[HEADERS]

    # Salvar arquivo
    saida = args.saida.strip() or f"VR_FINAL_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    with pd.ExcelWriter(saida, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="VR Mensal")
    print(f"💾 Arquivo gerado: {saida}")
    print(f"👥 Registros: {len(out)}")


if __name__ == "__main__":
    main()