import pandas as pd
import numpy as np
import os
import time

# ============================
# CONFIGURACIÓN
# ============================
ARCHIVO_INPUT = "clientes_asesores.xlsx"
HOJA = "Datos"
ARCHIVO_OUTPUT = "Resultado_Asesores_Internos.xlsx"

# TUS 11 ASESORES (Asegúrate que los nombres sean idénticos a los del Excel)
ASESORES = [
    "ASESOR_1", "ASESOR_2", "ASESOR_3", "ASESOR_4", "ASESOR_5", 
    "ASESOR_6", "ASESOR_7", "ASESOR_8", "ASESOR_9", "ASESOR_10", "ASESOR_11"
]

CUPO_TOP = 10  # Clientes > 10,000
CUPO_MID = 10  # Clientes 1,000 - 10,000

def distribuir_cupos(df_pool, asesores, cupo_max, tipo_lote):
    """
    Asigna clientes a asesores respetando la rotación (no repetir gestor).
    """
    asignaciones = [] # Lista de tuplas (Documento, Asesor_Nuevo)
    
    # Mezclamos aleatoriamente el pool para imparcialidad
    df_pool = df_pool.sample(frac=1, random_state=42).copy()
    
    # Diccionario para controlar cuántos lleva cada asesor
    conteo_asesores = {asesor: 0 for asesor in asesores}
    
    # Lista de clientes disponibles
    clientes_disponibles = df_pool.to_dict('records')
    
    # 1. Ronda de asignación
    # Iteramos cíclicamente sobre los asesores hasta llenar cupos o acabar clientes
    clientes_sin_asignar = []
    
    while clientes_disponibles:
        asignado_en_ronda = False
        
        for asesor in asesores:
            # Si el asesor ya llenó su cupo, saltar
            if conteo_asesores[asesor] >= cupo_max:
                continue
                
            # Buscar un candidato válido para este asesor
            candidato_elegido = None
            
            for i, cliente in enumerate(clientes_disponibles):
                gestor_anterior = str(cliente.get('Gestor_Asignado', '')).strip()
                
                # REGLA DE ORO: NO REPETIR GESTOR
                if gestor_anterior != asesor:
                    candidato_elegido = clientes_disponibles.pop(i) # Sacar de la lista
                    break
            
            if candidato_elegido:
                # Guardar asignación
                asignaciones.append({
                    'Documento': candidato_elegido['Documento'],
                    'Capital': candidato_elegido['Capital'],
                    'Gestor_Anterior': candidato_elegido.get('Gestor_Asignado', ''),
                    'Nuevo_Asesor': asesor,
                    'Cosecha': candidato_elegido['Cosecha'],
                    'Segmento': tipo_lote
                })
                conteo_asesores[asesor] += 1
                asignado_en_ronda = True
        
        # Si pasamos por todos los asesores y nadie pudo agarrar un cliente (bloqueos o cupos llenos)
        if not asignado_en_ronda:
            break
            
    # Lo que sobró se queda sin asignar en esta etapa
    sobrantes = len(clientes_disponibles)
    return pd.DataFrame(asignaciones), sobrantes

def main():
    start_time = time.time()
    print(">>> INICIANDO REASIGNACIÓN DE ASESORES INTERNOS...")
    
    try:
        df = pd.read_excel(ARCHIVO_INPUT, sheet_name=HOJA)
    except FileNotFoundError:
        print(f"ERROR: No existe '{ARCHIVO_INPUT}'")
        return

    # Limpieza
    df.columns = df.columns.str.strip()
    df["Capital"] = pd.to_numeric(df["Capital"], errors='coerce').fillna(0)
    df["Documento"] = df["Documento"].astype(str).str.strip()
    df["MC"] = df["MC"].astype(str).str.strip().str.upper()
    df["Gestor_Asignado"] = df["Gestor_Asignado"].astype(str).str.strip()
    
    # Filtro MC = CD+
    print(">>> Filtrando cartera CD+...")
    df_vip = df[df["MC"] == "CD+"].copy()
    
    # Agrupamos por DNI para que todas las cuentas del mismo cliente vayan juntas
    # Tomamos el Capital Total del cliente para decidir si es TOP o MID
    df_clientes_unicos = df_vip.groupby(['Documento', 'Cosecha', 'Gestor_Asignado'], as_index=False).agg({
        'Capital': 'sum'
    })
    
    resultados_totales = []
    
    cosechas = df_clientes_unicos["Cosecha"].unique()
    
    print(f">>> Procesando {len(cosechas)} cosechas para {len(ASESORES)} asesores.\n")
    
    for cosecha in cosechas:
        print(f"--- Cosecha {cosecha} ---")
        df_c = df_clientes_unicos[df_clientes_unicos["Cosecha"] == cosecha]
        
        # Segmentación
        # Rango 1: > 10,000
        pool_top = df_c[df_c["Capital"] > 10000].copy()
        # Rango 2: 1,000 a 10,000
        pool_mid = df_c[(df_c["Capital"] >= 1000) & (df_c["Capital"] <= 10000)].copy()
        
        print(f"   Candidates TOP (>10k): {len(pool_top)}")
        print(f"   Candidates MID (1k-10k): {len(pool_mid)}")
        
        # Asignación TOP
        res_top, sobras_top = distribuir_cupos(pool_top, ASESORES, CUPO_TOP, "TOP >10k")
        if not res_top.empty: resultados_totales.append(res_top)
        
        # Asignación MID
        res_mid, sobras_mid = distribuir_cupos(pool_mid, ASESORES, CUPO_MID, "MID 1k-10k")
        if not res_mid.empty: resultados_totales.append(res_mid)
        
        print(f"   Asignados: {len(res_top) + len(res_mid)} | Sin cupo/Bloqueados: {sobras_top + sobras_mid}")

    # Exportación
    if resultados_totales:
        df_final = pd.concat(resultados_totales, ignore_index=True)
        
        # Ordenar para presentación
        df_final = df_final.sort_values(by=['Cosecha', 'Nuevo_Asesor', 'Capital'], ascending=[True, True, False])
        
        df_final.to_excel(ARCHIVO_OUTPUT, index=False)
        print(f"\n>>> EXITO. Archivo generado: {ARCHIVO_OUTPUT}")
        print(f">>> Total clientes reasignados: {len(df_final)}")
        
        try: os.startfile(ARCHIVO_OUTPUT)
        except: pass
    else:
        print("\n>>> NO SE GENERARON ASIGNACIONES (Revisa si hay clientes CD+ con capital suficiente).")

if __name__ == "__main__":
    main()