import pandas as pd
import numpy as np
from ortools.sat.python import cp_model
import math
import os
import time

# ============================
# CONFIGURACIÓN
# ============================
ARCHIVO_EXCEL = "clientes.xlsx"
HOJA = "Datos"
OUTPUT_FILE = "asignacion_optimizada.xlsx"
# Asegúrate de que los nombres coincidan exactamente con la columna 'Gestor_Asignado'
EMPRESAS = ["ESCALL", "FINANCOBRO", "JYC"] 

def optimizar_bloque(df_bloque, empresas):
    """
    Optimiza la asignación de un bloque de clientes utilizando OR-Tools (CP-SAT).
    
    Reglas:
    1. Equidad de Capital y Cantidad de Cuentas entre empresas.
    2. Restricción de Retención: Máximo el 20% de la cartera original puede quedarse
       con el mismo gestor. El resto debe rotar obligatoriamente.
    3. Equidad por Zona geográfica (si el nivel de estrictez lo permite).
    """
    # Agrupación por Documento para manejar clientes únicos
    df_grouped = df_bloque.groupby('Documento', as_index=False).agg({
        'Capital': 'sum',
        'Cosecha': 'first',
        'Zona': 'first',
        'Num_Cuentas': 'first',
        'Gestor_Asignado': 'first'
    })

    n = len(df_grouped)
    m = len(empresas)
    empresa_to_idx = {nombre: i for i, nombre in enumerate(empresas)}

    # Caso borde: Pocos datos para optimizar
    if n < m * 10:
        df_grouped = df_grouped.sort_values('Capital', ascending=False)
        reparto = np.tile(empresas, int(np.ceil(n/m)))[:n]
        df_grouped['Empresa_Asignada'] = reparto
        return df_grouped[['Documento', 'Empresa_Asignada']], "Asignación Directa (Pocos datos)"

    # Pre-cálculo de valores enteros para el solver
    capitals = (df_grouped['Capital'] * 100).astype(int).values
    cuentas = df_grouped['Num_Cuentas'].astype(int).values
    gestores_origen = df_grouped['Gestor_Asignado'].values
    
    # Niveles de relajación de restricciones
    intentos = [
        {'tol': 0.02, 'usa_zona': True,  'desc': 'Estricto (2%)'},
        {'tol': 0.05, 'usa_zona': True,  'desc': 'Medio (5%)'},
        {'tol': 0.05, 'usa_zona': False, 'desc': 'Flexible (5%, Sin Zona)'},
        {'tol': 0.15, 'usa_zona': False, 'desc': 'Rescate (15%, Sin Zona)'}
    ]

    for intento in intentos:
        tol = intento['tol']
        usa_zona = intento['usa_zona']
        model = cp_model.CpModel()
        
        # Variables de decisión: x[empresa][cliente]
        x = {}
        for j in range(m):
            x[j] = [model.NewBoolVar(f'x_{i}_{j}') for i in range(n)]

        # Constraint 1: Cada cliente debe ser asignado a una sola empresa
        for i in range(n):
            model.Add(sum(x[j][i] for j in range(m)) == 1)

        # Constraint 2: Retención Máxima del 20%
        PORCENTAJE_RETENCION = 0.20
        for nombre_empresa in empresas:
            if nombre_empresa in empresa_to_idx:
                idx_empresa = empresa_to_idx[nombre_empresa]
                indices_propios = [i for i, g in enumerate(gestores_origen) if g == nombre_empresa]
                
                if indices_propios:
                    max_permitidos = math.ceil(len(indices_propios) * PORCENTAJE_RETENCION)
                    model.Add(sum(x[idx_empresa][i] for i in indices_propios) <= max_permitidos)

        # Constraint 3: Equidad Global (Capital y Cuentas)
        total_cap = capitals.sum()
        total_cnt = cuentas.sum()
        ideal_cap = total_cap // m
        ideal_cnt = total_cnt // m
        
        for j in range(m):
            model.Add(cp_model.LinearExpr.WeightedSum(x[j], capitals) >= int(ideal_cap * (1 - tol)))
            model.Add(cp_model.LinearExpr.WeightedSum(x[j], capitals) <= int(ideal_cap * (1 + tol)))
            model.Add(cp_model.LinearExpr.WeightedSum(x[j], cuentas) >= int(ideal_cnt * (1 - tol*1.5)))
            model.Add(cp_model.LinearExpr.WeightedSum(x[j], cuentas) <= int(ideal_cnt * (1 + tol*1.5)))

        # Constraint 4: Equidad por Zona
        if usa_zona:
            for zona, idxs in df_grouped.groupby('Zona').groups.items():
                if len(idxs) >= m * 2: 
                    subset_cuentas = cuentas[idxs]
                    ideal_zona = subset_cuentas.sum() / m
                    tol_z = max(tol * 2, 0.10) 
                    for j in range(m):
                        vars_in_zone = [x[j][k] for k in idxs]
                        weights_in_zone = [cuentas[k] for k in idxs]
                        model.Add(cp_model.LinearExpr.WeightedSum(vars_in_zone, weights_in_zone) >= int(ideal_zona * (1 - tol_z)))
                        model.Add(cp_model.LinearExpr.WeightedSum(vars_in_zone, weights_in_zone) <= int(ideal_zona * (1 + tol_z)))

        # Ejecución del Solver
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 30
        solver.parameters.num_search_workers = 8
        status = solver.Solve(model)

        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            asignaciones = [None] * n
            for i in range(n):
                for j in range(m):
                    if solver.Value(x[j][i]) == 1:
                        asignaciones[i] = empresas[j]
                        break
            df_grouped['Empresa_Asignada'] = asignaciones
            return df_grouped[['Documento', 'Empresa_Asignada']], intento['desc']
            
    # Fallback Heurístico (si falla la optimización)
    df_grouped = df_grouped.sort_values('Capital', ascending=False)
    reparto = np.tile(empresas, int(np.ceil(n/m)))[:n]
    df_grouped['Empresa_Asignada'] = reparto
    return df_grouped[['Documento', 'Empresa_Asignada']], "Heurística (Fallo Solver)"

def main():
    start_time = time.time()
    print(">>> INICIANDO PROCESO DE REASIGNACIÓN...")
    
    try:
        df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA)
    except FileNotFoundError:
        print(f"ERROR CRITICO: No se encuentra el archivo '{ARCHIVO_EXCEL}'")
        return

    # Limpieza de datos
    df["Año_Castigo"] = pd.to_numeric(df["Año_Castigo"], errors='coerce').fillna(0)
    df["Capital"] = pd.to_numeric(df["Capital"], errors='coerce').fillna(0)
    df["Documento"] = df["Documento"].astype(str).str.strip()
    
    # Cálculo automático de cuentas por cliente
    if 'Num_Cuentas' not in df.columns:
        print(">>> Calculando columna auxiliar 'Num_Cuentas'...")
        df['Num_Cuentas'] = df.groupby('Documento')['Documento'].transform('count')

    # Separación de cartera inalterable
    mask_inalterable = df["Inalterables"].astype(str).str.upper() == "INALTERABLE"
    df_inalterables = df[mask_inalterable].copy()
    df_inalterables["Nueva_Empresa"] = df_inalterables["Gestor_Asignado"]
    
    df_reasignar = df[~mask_inalterable].copy()
    print(f">>> Registros a procesar: {len(df_reasignar)}")

    resultados = []
    cosechas = df_reasignar["Cosecha"].unique()
    
    # Procesamiento por Lotes (Cosechas)
    for i, cosecha in enumerate(cosechas, 1):
        df_c = df_reasignar[df_reasignar["Cosecha"] == cosecha]
        print(f"[{i}/{len(cosechas)}] Procesando: {str(cosecha):<15} | Clientes: {len(df_c):<6} ...", end=" ")
        
        try:
            mapa, metodo = optimizar_bloque(df_c, EMPRESAS)
            resultados.append(mapa)
            print(f"OK -> {metodo}")
        except Exception as e:
            print(f"ERROR: {e}")
            # Fallback de emergencia simple
            temp = df_c[['Documento']].drop_duplicates()
            temp['Empresa_Asignada'] = np.random.choice(EMPRESAS, len(temp))
            resultados.append(temp)

    # Unificación y Exportación
    print("\n>>> GENERANDO ARCHIVO FINAL...")
    if resultados:
        df_mapa = pd.concat(resultados, ignore_index=True)
        
        df_reasignar = df_reasignar.merge(df_mapa, on='Documento', how='left')
        df_reasignar["Nueva_Empresa"] = df_reasignar["Empresa_Asignada"]
        df_reasignar.drop(columns=['Empresa_Asignada'], inplace=True)
        
        df_final = pd.concat([df_inalterables, df_reasignar], ignore_index=True)
        
        # Relleno de seguridad
        mask_null = df_final["Nueva_Empresa"].isna()
        if mask_null.sum() > 0:
            df_final.loc[mask_null, "Nueva_Empresa"] = np.random.choice(EMPRESAS, mask_null.sum())

        df_final.to_excel(OUTPUT_FILE, index=False)
        print(f"Proceso finalizado en {time.time() - start_time:.2f} seg. Archivo: {OUTPUT_FILE}")
        
        try: os.startfile(OUTPUT_FILE)
        except: pass
    else:
        print("No se generaron resultados.")

if __name__ == "__main__":
    main()