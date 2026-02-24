import pyautogui
import time
import pandas as pd
from utils import forzar_caps_off, f_monto

def ejecutar_cheques(df, total, log_func, progress_func, velocidad):
    """
    Robot de Cheques - Versión Controlada (H: Nombre, I: CUIT, J: Monto).
    Esta versión carga los datos en la grilla para revisión, SIN grabar.
    """
    pyautogui.PAUSE = velocidad
    forzar_caps_off()
    
    log_func("💰 Iniciando Carga de Validación (H: Nombre, I: CUIT, J: Monto)...")
    
    try:
        # --- EXTRACCIÓN DE CABECERA ---
        # Tomamos la entidad de A1
        entidad = str(df.iloc[0, 0]).split('.')[0]
        log_func(f"📌 Entidad: {entidad}")

        # --- NAVEGACIÓN INICIAL ---
        pyautogui.write('2')
        pyautogui.press('enter', presses=6, interval=0.1)
        time.sleep(0.5)
        
        pyautogui.write(entidad); pyautogui.press('enter')
        pyautogui.write('afd'); pyautogui.press('enter')
        pyautogui.write('0'); pyautogui.press('enter')
        time.sleep(1.2)

        # --- BUCLE DE CARGA DE FILAS ---
        for i, row in df.iterrows():
            # Saltamos si la Referencia (Col B) está vacía
            if pd.isna(row[1]): 
                continue 
            
            # Mapeo de Columnas (A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9)
            ref      = str(row[1]).split('.')[0].strip() # Col B
            serie    = str(row[2]).strip()               # Col C
            nro_ch   = str(row[3]).split('.')[0].strip() # Col D
            f_orig   = str(row[4]).strip()               # Col E
            f_depo   = str(row[5]).strip()               # Col F
            banco    = str(row[6]).strip()               # Col G
            nombre_h = str(row[7]).strip() if not pd.isna(row[7]) else "" # Col H
            cuit_i   = str(row[8]).split('.')[0].strip() if not pd.isna(row[8]) else "" # Col I
            monto_j  = f_monto(row[9])                  # Col J

            log_func(f"▶️ Fila {i+1}: Cheque {nro_ch} | Monto: {monto_j}")

            # 1. Escribir Referencia
            pyautogui.write(ref); pyautogui.press('enter')
            time.sleep(0.8) 

            # 2. Datos Internos
            pyautogui.write(serie); pyautogui.press('enter')
            pyautogui.write(nro_ch); pyautogui.press('enter')
            pyautogui.write(f_orig); pyautogui.press('enter')
            pyautogui.write(f_depo); pyautogui.press('enter')
            pyautogui.write(banco); pyautogui.press('enter')
            
            # T de control (Shift + T)
            with pyautogui.hold('shift'): 
                pyautogui.press('t')
            pyautogui.press('enter')
            
            # --- CARGA EN COLUMNAS H, I, J ---
            # Escribir Nombre (H)
            pyautogui.write(nombre_h); pyautogui.press('enter')
            
            # Escribir CUIT (I)
            pyautogui.write(cuit_i)
            
            # 3 Enters para saltar desde CUIT hasta el Monto (J)
            pyautogui.press('enter', presses=3, interval=0.15)
            
            # Escribir Monto (J)
            pyautogui.write(monto_j); pyautogui.press('enter')
            
            time.sleep(0.6)
            progress_func((i + 1) / total)

        # --- CIERRE DE PROCESO (SIN GRABAR) ---
        log_func("🛑 Carga finalizada. Esperando revisión manual.")
        
        # Alerta para detener el flujo y avisarte que revises
        pyautogui.alert(
            "El robot terminó la carga.\n\n"
            "POR FAVOR REVISÁ:\n"
            "- Nombre, CUIT y Monto en la grilla.\n\n"
            "NO se presionó F5 ni se guardó la información.",
            "Revisión de Control"
        )

    except Exception as e:
        log_func(f"❌ Error en proceso: {e}")