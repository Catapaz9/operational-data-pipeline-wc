import pandas as pd
import re

def normalizar_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        re.sub(r"\s+", " ", str(c)).strip()   
        for c in df.columns
    ]
    return df


def calcular_total_por_wc_code(df: pd.DataFrame, ruta_porcentajes: str) -> pd.DataFrame:
    df = normalizar_cols(df)

    por = pd.read_excel(ruta_porcentajes)
    por = normalizar_cols(por)
    por.columns = [c.upper() for c in por.columns]  

    if "WC CODE" not in por.columns or "MARK UP" not in por.columns:
        raise ValueError(f"PORCENTAJES_MUESTRA.xlsx debe tener columnas 'WC CODE' y 'MARK UP'. Columnas actuales: {list(por.columns)}")

    por["WC CODE"] = pd.to_numeric(por["WC CODE"], errors="coerce")
    por["MARK UP"] = pd.to_numeric(por["MARK UP"], errors="coerce")
    markup_map = dict(zip(por["WC CODE"], por["MARK UP"]))

    if "TOTAL" not in df.columns:
        df["TOTAL"] = ""

    df["WC CODE"] = pd.to_numeric(df["WC CODE"], errors="coerce")
    df["_WC_CODE_FF"] = df["WC CODE"].ffill()

    df["_SUBTOTAL_NUM"] = pd.to_numeric(df["SUBTOTAL"], errors="coerce")

    emp = df["EMPLOYEE"].astype(str)
    mask_totals = emp.str.startswith("Totals for", na=False)

    df.loc[mask_totals, "_MARKUP"] = df.loc[mask_totals, "_WC_CODE_FF"].map(markup_map)

    df.loc[mask_totals, "TOTAL"] = (
        df.loc[mask_totals, "_SUBTOTAL_NUM"] * df.loc[mask_totals, "_MARKUP"]
    ).round(2)

    df.drop(columns=["_WC_CODE_FF", "_SUBTOTAL_NUM", "_MARKUP"], inplace=True, errors="ignore")
    return df


def generar_resumen_por_cliente(df: pd.DataFrame) -> pd.DataFrame:
    df = normalizar_cols(df)
    df = df.copy()
    
    df["CLIENT"] = df["CLIENT"].where(df["CLIENT"].notna()).ffill()
    df["WC CODE"] = pd.to_numeric(df["WC CODE"], errors="coerce").ffill()

    df["EMPLOYEE"] = df["EMPLOYEE"].astype(str)
    mask_totals = df["EMPLOYEE"].str.startswith("Totals for", na=False)
    out = df.loc[mask_totals].copy()

    out["EMPLOYEE"] = out["EMPLOYEE"].str.replace(r"^Totals for\s+", "", regex=True).str.strip()

    columnas = ["CLIENT", "WC CODE", "EMPLOYEE", "REG PAY", "OT PAY", "DT PAY", "SUBTOTAL", "TOTAL"]
    faltan = [c for c in columnas if c not in out.columns]
    if faltan:
        raise ValueError(f"Faltan columnas para generar el resumen final: {faltan}. Columnas actuales: {list(out.columns)}")

    out = out[columnas].copy()

    for c in ["REG PAY", "OT PAY", "DT PAY", "SUBTOTAL", "TOTAL"]:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0)

    
    same_group = out["CLIENT"].eq(out["CLIENT"].shift(1)) & out["WC CODE"].eq(out["WC CODE"].shift(1))
    out.loc[same_group, ["CLIENT", "WC CODE"]] = ""

    return out


if __name__ == "__main__":
    ruta_entrada = "wc_sample_anonimizado.xlsx"
    ruta_porcentajes = "PORCENTAJES_MUESTRA.xlsx"
    ruta_salida = "wc_sample_anonimizado_RESUMEN.xlsx"

    df = pd.read_excel(ruta_entrada)
    df = normalizar_cols(df)

    requeridas = ["CLIENT", "WC CODE", "EMPLOYEE", "REG PAY", "OT PAY", "DT PAY", "SUBTOTAL"]
    faltan = [c for c in requeridas if c not in df.columns]
    if faltan:
        raise ValueError(f"Falta(n) columna(s) requerida(s) en {ruta_entrada}: {faltan}. Columnas actuales: {list(df.columns)}")

    df = calcular_total_por_wc_code(df, ruta_porcentajes)
    resumen_final = generar_resumen_por_cliente(df)

    resumen_final.to_excel(ruta_salida, index=False)
    print(f"Archivo creado: {ruta_salida}")






