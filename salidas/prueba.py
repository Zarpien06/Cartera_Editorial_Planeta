import pandas as pd
import numpy as np
from datetime import datetime
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import train_test_split

# -----------------------------
# CONFIGURACIÓN DE ARCHIVOS
# -----------------------------
provision_file = "provca.csv"
anticipos_file = "anticipos.csv"
TRM_file = "TRM.csv"  # archivo con TRM del mes anterior
output_file = "Formato_Deuda_IA.xlsx"
ultimo_dia_mes = "2026-01-31"

# -----------------------------
# FUNCIONES AUXILIARES
# -----------------------------
def calcular_dias(fecha_vto, cierre):
    fecha_vto = pd.to_datetime(fecha_vto, dayfirst=True)
    cierre = pd.to_datetime(cierre)
    dias_vencidos = max((cierre - fecha_vto).days, 0)
    dias_por_vencer = max((fecha_vto - cierre).days, 0)
    return dias_vencidos, dias_por_vencer

def calcular_dotacion(saldo, dias_vencidos):
    if dias_vencidos >= 180:
        return 100, saldo
    else:
        return 0, 0

# -----------------------------
# 1️⃣ LECTURA DE ARCHIVOS
# -----------------------------
prov = pd.read_csv(provision_file)
antic = pd.read_csv(anticipos_file)
trm = pd.read_csv(TRM_file)  # columnas: Moneda, TRM

# -----------------------------
# 2️⃣ TRANSFORMACIÓN PROVISIÓN
# -----------------------------
prov.rename(columns={
    "PCCDEM": "Codigo",
    "PCCDAC": "Empresa",
    "PCNMCL": "NombreCliente",
    "PCNMCM": "Denominacion",
    "PCFEFA": "FechaFactura",
    "PCFEVE": "FechaVencimiento",
    "PCSALD": "Saldo"
}, inplace=True)

prov["Denominacion"] = prov["Denominacion"].combine_first(prov["NombreCliente"])
prov["FechaFactura"] = pd.to_datetime(prov["FechaFactura"], dayfirst=True)
prov["FechaVencimiento"] = pd.to_datetime(prov["FechaVencimiento"], dayfirst=True)

prov[["DiasVencidos", "DiasPorVencer"]] = prov.apply(
    lambda row: pd.Series(calcular_dias(row["FechaVencimiento"], ultimo_dia_mes)),
    axis=1
)
prov[["PorcDotacion", "ValorDotacion"]] = prov.apply(
    lambda row: pd.Series(calcular_dotacion(row["Saldo"], row["DiasVencidos"])),
    axis=1
)

# -----------------------------
# 3️⃣ TRANSFORMACIÓN ANTICIPOS
# -----------------------------
antic.rename(columns={"VALOR ANTICIPO": "ValorAnticipo"}, inplace=True)
antic["ValorAnticipo"] *= -1

# -----------------------------
# 4️⃣ CREACIÓN MODELO DEUDA
# -----------------------------
modelo = pd.concat([prov, antic], ignore_index=True, sort=False)

# -----------------------------
# 5️⃣ IA: Predicción de riesgo
# -----------------------------
# Creamos datos de ejemplo para entrenar el modelo de riesgo
# En la vida real usarías históricos de mora/incobrable
# Aquí simulamos con los mismos datos
modelo["Incobrable"] = np.where(modelo["DiasVencidos"] >= 180, 1, 0)

# Variables predictoras
X = modelo[["Saldo", "DiasVencidos", "DiasPorVencer", "PorcDotacion"]].fillna(0)
y = modelo["Incobrable"]

# Escalado
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

# Entrenamiento rápido
X_train, X_test, y_train, y_test = train_test_split(X_scaled, y, test_size=0.2, random_state=42)
clf = RandomForestClassifier(n_estimators=100, random_state=42)
clf.fit(X_train, y_train)

# Predicción de riesgo (probabilidad de incobrable)
modelo["Riesgo"] = clf.predict_proba(X_scaled)[:,1]
# Semáforo
modelo["Semaforo"] = pd.cut(modelo["Riesgo"], bins=[-0.01,0.3,0.6,1.0], labels=["Verde","Amarillo","Rojo"])

# -----------------------------
# 6️⃣ Hoja Divisas (convertir con TRM)
# -----------------------------
def convertir_trm(row):
    if "USD" in row.get("Moneda", ""):
        trm_val = trm.loc[trm["Moneda"]=="USD", "TRM"].values[0]
        return row.get("Saldo",0)*trm_val
    elif "EUR" in row.get("Moneda", ""):
        trm_val = trm.loc[trm["Moneda"]=="EUR", "TRM"].values[0]
        return row.get("Saldo",0)*trm_val
    return row.get("Saldo",0)

modelo["SaldoPesos"] = modelo.apply(convertir_trm, axis=1)

# -----------------------------
# 7️⃣ EXPORTAR A EXCEL
# -----------------------------
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    modelo.to_excel(writer, sheet_name="Pesos", index=False)
    
    # Resumen IA
    resumen = modelo.groupby("Semaforo").agg({"Saldo":"sum", "ValorDotacion":"sum"}).reset_index()
    resumen.to_excel(writer, sheet_name="IA_Insights", index=False)

print(f"Formato Deuda generado con IA: {output_file}")
