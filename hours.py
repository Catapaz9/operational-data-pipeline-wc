import pandas as pd
import numpy as np
import re
from datetime import datetime

def _to_float_from_text(s: str) -> float:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return float("nan")
    s = str(s).strip().replace(",", ".")
    m = re.search(r"(\d+(\.\d+)?)", s)
    return float(m.group(1)) if m else float("nan")

def leer_pay_rates(ruta_divisores: str) -> np.ndarray:
    div = pd.read_excel(ruta_divisores)
    if "PAY RATE" not in div.columns:
        raise ValueError("DIVISORES.xlsx no tiene columna 'PAY RATE'")

    rates = (
        div["PAY RATE"]
        .apply(_to_float_from_text)   # convierte 16,1 -> 16.1
        .dropna()
        .astype(float)
        .unique()
    )
    rates = np.sort(rates)
    if len(rates) == 0:
        raise ValueError("No se encontraron PAY RATE válidos en DIVISORES.xlsx")
    return rates

def generar_resumen_final(ruta_resumen: str, ruta_divisores: str, ruta_salida: str):
    df = pd.read_excel(ruta_resumen)

    required = ["REG PAY", "OT PAY", "DT PAY"]
    faltan = [c for c in required if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan columnas en el resumen: {faltan}")

    # Asegurar numéricos (por si vienen con coma)
    for c in required:
        df[c] = df[c].astype(str).str.replace(",", ".", regex=False)
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    rates = leer_pay_rates(ruta_divisores)

    # Elegir PAY RATE por fila: primer rate que haga REG PAY / PAY RATE <= 40
    rate_req = df["REG PAY"].to_numpy() / 40.0
    idx = np.searchsorted(rates, rate_req, side="left")
    idx = np.clip(idx, 0, len(rates) - 1)
    pay_rate = rates[idx]

    df["PAY RATE"] = np.round(pay_rate, 2)
    df["REG HOURS"] = np.round(df["REG PAY"].to_numpy() / pay_rate, 2)

    # Rates OT/DT: si no hay pago, dejar rate = 0 (como tu imagen)
    df["OT RATE"] = np.where(df["OT PAY"].to_numpy() > 0, np.round(df["PAY RATE"].to_numpy() * 1.5, 2), 0)
    df["DT RATE"] = np.where(df["DT PAY"].to_numpy() > 0, np.round(df["PAY RATE"].to_numpy() * 2.0, 2), 0)

    # Horas OT/DT (división segura)
    df["OT HOURS"] = np.where(df["OT RATE"].to_numpy() > 0, np.round(df["OT PAY"].to_numpy() / df["OT RATE"].to_numpy(), 2), 0)
    df["DT HOURS"] = np.where(df["DT RATE"].to_numpy() > 0, np.round(df["DT PAY"].to_numpy() / df["DT RATE"].to_numpy(), 2), 0)

    # Check de seguridad: REG HOURS debe ser <= 40 sí o sí
    if (df["REG HOURS"] > 40.0001).any():
        malos = df.loc[df["REG HOURS"] > 40.0001, ["EMPLOYEE", "REG PAY", "PAY RATE", "REG HOURS"]].head(10)
        raise ValueError(f"Quedaron filas con REG HOURS > 40. Ejemplos:\n{malos}")

    # Orden final de columnas (igual a tu template)
    orden = [
        "CLIENT", "WC CODE", "EMPLOYEE",
        "REG HOURS", "PAY RATE", "REG PAY",
        "OT HOURS", "OT RATE", "OT PAY",
        "DT HOURS", "DT RATE", "DT PAY",
        "SUBTOTAL", "TOTAL"
    ]
    orden = [c for c in orden if c in df.columns]
    df = df[orden + [c for c in df.columns if c not in orden]]

    df.to_excel(ruta_salida, index=False)
    print(f"✅ Listo: {ruta_salida}")

if __name__ == "__main__":
    ruta_resumen   = r"C:\Users\Cata\OneDrive\Desktop\WC project\wc_sample_anonimizado_RESUMEN.xlsx"
    ruta_divisores = r"C:\Users\Cata\OneDrive\Desktop\WC project\DIVISORES.xlsx"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta_salida = rf"C:\Users\Cata\OneDrive\Desktop\WC project\wc_sample_anonimizado_RESUMEN_FINAL{stamp}.xlsx"
    generar_resumen_final(ruta_resumen, ruta_divisores, ruta_salida)

