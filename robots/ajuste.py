import pyautogui
import time
import pandas as pd
from utils import forzar_caps_off

def ejecutar_ajuste(df, total, log_func, progress_func, velocidad):
    """
    Robot de Ajuste de Stock.
    Clon de lógica de Stock (Cabecera C1, C2, C3 y ráfaga de ítems).
    """
    pyautogui.PAUSE = velocidad
    forzar_caps_off()
    
    # 1. NAVEGACIÓN AL MÓDULO DE AJUSTE (3-6-2)
    log_func("Entrando al módulo de Ajuste (3-6-2)...")
    for k in ['3', '6', '2']: 
        pyautogui.write(k)
        pyautogui.press('enter')
        time.sleep(0.4)
    
    # 2. CARGA DE CABECERA (IGUAL A STOCK: C1, C2, C3)
    try:
        # Extraemos exactamente igual que en el robot de PuTTY
        pedido = str(df.iloc[0, 2]).split('.')[0] if not pd.isna(df.iloc[0, 2]) else ""
        obs    = str(df.iloc[1, 2]) if not pd.isna(df.iloc[1, 2]) else ""
        imp    = str(df.iloc[2, 2]).upper() if not pd.isna(df.iloc[2, 2]) else "AS" # 'AS' como default para Ajuste
        
        log_func(f"📌 Cabecera: Doc {pedido} | Motivo: {obs} | Tipo: {imp}")

        # Secuencia exacta de cabecera
        pyautogui.write(pedido); pyautogui.press('enter'); time.sleep(0.7)
        pyautogui.press('enter') 
        pyautogui.write(obs); pyautogui.press('enter'); time.sleep(0.4)
        pyautogui.press('enter')
        pyautogui.write(imp); pyautogui.press('enter'); time.sleep(1.0)
        
        # Re-ingreso para entrar a la grilla (usando el nro de documento)
        pyautogui.write(pedido); pyautogui.press('enter'); time.sleep(1.2)
        
    except Exception as e:
        log_func(f"❌ Error en cabecera de ajuste: {e}")
        return

    # 3. BUCLE DE ARTÍCULOS (MISMA LÓGICA GANADORA)
    log_func("⏳ Iniciando carga de ítems...")
    
    for i, row in df.iterrows():
        val_sku = row[0]
        if pd.isna(val_sku): continue
        
        sku_str = str(val_sku).strip().lower()
        if sku_str in ['codigo', 'sku', 'articulo', 'código', 'artículo']:
            continue

        try:
            sku = str(val_sku).split('.')[0].strip()
            # Cantidad (columna B). Soporta números negativos como me pasaste (-10, -20)
            cantidad_raw = str(row[1]).strip()
            cantidad = str(int(float(cantidad_raw))) if not pd.isna(row[1]) else "0"
            
            info_d = str(row[3]).strip() if len(row) > 3 and not pd.isna(row[3]) and str(row[3]).lower() != 'nan' else None

            # --- SECUENCIA EN PUTTY ---
            pyautogui.write(sku)
            
            # Los 2 ENTERS del nuevo menú
            pyautogui.press('enter', presses=4, interval=0.1)

            if info_d:
                pyautogui.write('g')
                pyautogui.write(cantidad)
                pyautogui.press('enter')
                time.sleep(0.4)
                pyautogui.write(info_d)
                pyautogui.press('enter')
            else:
                # El sistema toma el negativo (ej: u-10) sin problemas
                pyautogui.write(f"u{cantidad}")
                pyautogui.press('enter')

            time.sleep(0.1) 
            log_func(f"✅ Ajustado: {sku} | Cant: {cantidad}")

        except Exception as row_err:
            log_func(f"❌ Error en fila {i+1}: {row_err}")
            
        progress_func((i + 1) / total)

    # 4. CIERRE Y GUARDADO (MISMA SECUENCIA)
    log_func("💾 Guardando y saliendo...")
    pyautogui.press('f5'); time.sleep(2.5)
    for k in ['end', 'enter', 'end', 'end']:
        pyautogui.press(k)
        time.sleep(0.5)

    log_func("🏁 Ajuste finalizado con éxito.")