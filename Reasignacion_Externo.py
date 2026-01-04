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
OUTPUT_FILE = "Reasignacion.xlsx"
EMPRESAS = ["ESCALL", "FINANCOBRO", "JYC"] 

def optimizar_bloque(df_bloque, empresas):
    """
    Optimiza la asignación GLOBAL para evitar duplicados y conflictos.
    """
    # 1. Agrupación por Documento (ESTO ES LA CLAVE)
    # Al agrupar aquí, colapsamos las 3 filas de tu cliente en 1 sola para la toma de decisión.
    df_grouped = df_bloque.groupby('Documento', as_index=False).agg({
        'Capital': 'sum',           # Suma la deuda de todas las cosechas
        'Num_Cuentas': 'first',     # Toma el conteo total (ya calculado fuera)
        'Gestor_Asignado': 'first', # Toma el primer gestor (referencial para veto)
        'Zona': 'first'
    })

    n = len(df_grouped)
    m = len(empresas)
    empresa_to_idx = {nombre: i for i, nombre in enumerate(empresas)}

    # Caso borde: Pocos datos
    if n < m * 10:
        df_grouped = df_grouped.sort_values('Capital', ascending=False)
        reparto = np.tile(empresas, int(np.ceil(n/m)))[:n]
        df_grouped['Empresa_Asignada'] = reparto
        return df_grouped[['Documento', 'Empresa_Asignada']], "Asignación Directa (Pocos datos)"

    capitals = (df_grouped['Capital'] * 100).astype(int).values
    cuentas = df_grouped['Num_Cuentas'].astype(int).values
    gestores_origen = df_grouped['Gestor_Asignado'].values
    
    # Intentos de optimización
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
        
        x = {}
        for j in range(m):
            x[j] = [model.NewBoolVar(f'x_{i}_{j}') for i in range(n)]

        # Constraint 1: Un cliente = Una sola empresa
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

        # Constraint 3: Equidad Global
        total_cap = capitals.sum()
        total_cnt = cuentas.sum()
        ideal_cap = total_cap // m
        ideal_cnt = total_cnt // m
        
        for j in range(m):
            model.Add(cp_model.LinearExpr.WeightedSum(x[j], capitals) >= int(ideal_cap * (1 - tol)))
            model.Add(cp_model.LinearExpr.WeightedSum(x[j], capitals) <= int(ideal_cap * (1 + tol)))
            model.Add(cp_model.LinearExpr.WeightedSum(x[j], cuentas) >= int(ideal_cnt * (1 - tol*1.5)))
            model.Add(cp_model.LinearExpr.WeightedSum(x[j], cuentas) <= int(ideal_cnt * (1 + tol*1.5)))

        # Constraint 4: Zona
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

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 45 # Aumentamos un poco el tiempo ya que procesa todo junto
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
            
    # Fallback
    df_grouped = df_grouped.sort_values('Capital', ascending=False)
    reparto = np.tile(empresas, int(np.ceil(n/m)))[:n]
    df_grouped['Empresa_Asignada'] = reparto
    return df_grouped[['Documento', 'Empresa_Asignada']], "Heurística (Fallo Solver)"

def main():
    start_time = time.time()
    print(">>> INICIANDO PROCESO UNIFICADO...")
    
    try:
        df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA)
    except FileNotFoundError:
        print(f"ERROR: No existe '{ARCHIVO_EXCEL}'")
        return

    # Limpieza
    df["Año_Castigo"] = pd.to_numeric(df["Año_Castigo"], errors='coerce').fillna(0)
    df["Capital"] = pd.to_numeric(df["Capital"], errors='coerce').fillna(0)
    df["Documento"] = df["Documento"].astype(str).str.strip()
    
    # 1. CÁLCULO DE CUENTAS (Esto es vital para que Num_Cuentas sea correcto en el global)
    if 'Num_Cuentas' not in df.columns:
        print(">>> Calculando total de cuentas por cliente...")
        df['Num_Cuentas'] = df.groupby('Documento')['Documento'].transform('count')

    # Separar Inalterables
    mask_inalterable = df["Inalterables"].astype(str).str.upper() == "INALTERABLE"
    df_inalterables = df[mask_inalterable].copy()
    df_inalterables["Nueva_Empresa"] = df_inalterables["Gestor_Asignado"]
    
    df_reasignar = df[~mask_inalterable].copy()
    print(f">>> A Reasignar: {len(df_reasignar)} registros (filas de deuda).")

    # ==============================================================================
    # CAMBIO CRÍTICO: NO HAY BUCLE FOR COSECHA. SE PROCESA TODO JUNTO.
    # ==============================================================================
    print(">>> Optimizando cartera completa (esto evita duplicados)...")
    
    try:
        # Enviamos TODO el bloque. El optimizador agrupará internamente por DNI.
        mapa_global, metodo = optimizar_bloque(df_reasignar, EMPRESAS)
        print(f"   EXITO -> Método usado: {metodo}")
        
        # Merge Seguro:
        # df_reasignar tiene 'N' filas (ej: 3 filas para el cliente X)
        # mapa_global tiene '1' fila por cliente (ej: 1 fila para el cliente X)
        # Resultado del merge: 3 filas, todas con la misma empresa.
        df_reasignar = df_reasignar.merge(mapa_global, on='Documento', how='left')
        df_reasignar["Nueva_Empresa"] = df_reasignar["Empresa_Asignada"]
        df_reasignar.drop(columns=['Empresa_Asignada'], inplace=True)
        
    except Exception as e:
        print(f"ERROR CRITICO EN OPTIMIZACION: {e}")
        # Fallback de emergencia
        df_reasignar["Nueva_Empresa"] = "ERROR_ASIGNACION"

    # Unificación final
    print(">>> Generando archivo final...")
    df_final = pd.concat([df_inalterables, df_reasignar], ignore_index=True)
    
    # Relleno de seguridad por si alguno quedó huérfano
    mask_null = df_final["Nueva_Empresa"].isna()
    if mask_null.sum() > 0:
        print(f"AVISO: {mask_null.sum()} registros quedaron sin asignar. Asignando aleatoriamente.")
        df_final.loc[mask_null, "Nueva_Empresa"] = np.random.choice(EMPRESAS, mask_null.sum())

    df_final.to_excel(OUTPUT_FILE, index=False)
    print(f"Terminado en {time.time() - start_time:.2f} seg. Archivo: {OUTPUT_FILE}")
    
    try: os.startfile(OUTPUT_FILE)
    except: pass

if __name__ == "__main__":
    main()