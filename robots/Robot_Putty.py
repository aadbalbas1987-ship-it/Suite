import pyautogui
import time
import pandas as pd
from utils import limpiar_sku, forzar_caps_off

def ejecutar_stock(df, total, log_func, progress_func, velocidad):
    """
    Versión Final Optimizada: Bucle continuo de artículos.
    Secuencia: SKU -> 2 ENTER -> u+Cantidad -> 1 ENTER (salto al sig. SKU)
    """
    pyautogui.PAUSE = velocidad
    forzar_caps_off()
    
    # 1. NAVEGACIÓN INICIAL AL MÓDULO (3-6-1)
    log_func("Entrando al módulo de stock (3-6-1)...")
    for k in ['3', '6', '1']:
        pyautogui.write(k)
        pyautogui.press('enter')
        time.sleep(0.4)
    
    # 2. CARGA DE CABECERA (Columna C)
    try:
        pedido = str(df.iloc[0, 2]).split('.')[0] if not pd.isna(df.iloc[0, 2]) else ""
        obs = str(df.iloc[1, 2]) if not pd.isna(df.iloc[1, 2]) else ""
        imp = str(df.iloc[2, 2]).upper() if not pd.isna(df.iloc[2, 2]) else "LP1"
        
        log_func(f"📌 Cabecera: Pedido {pedido} | Obs: {obs}")

        pyautogui.write(pedido); pyautogui.press('enter'); time.sleep(0.7)
        pyautogui.press('enter') 
        pyautogui.write(obs); pyautogui.press('enter'); time.sleep(0.4)
        pyautogui.press('enter')
        pyautogui.write(imp); pyautogui.press('enter'); time.sleep(1.0)
        
        # Re-ingreso para entrar a la grilla de carga
        pyautogui.write(pedido); pyautogui.press('enter'); time.sleep(1.2)
        
    except Exception as e:
        log_func(f"❌ Error en cabecera: {e}")
        return

    # 3. BUCLE DE ARTÍCULOS (LOOP CONTINUO)
    log_func("⏳ Iniciando ráfaga de carga...")
    
    for i, row in df.iterrows():
        val_sku = row[0]
        if pd.isna(val_sku): continue
        
        # Saltamos encabezados
        sku_str = str(val_sku).strip().lower()
        if sku_str in ['codigo', 'sku', 'articulo', 'código', 'artículo']:
            continue
        
        try:
            # Datos: A (SKU), B (Cantidad), D (Modo G)
            sku = str(val_sku).split('.')[0].strip()
            cantidad = str(int(float(str(row[1]).strip()))) if not pd.isna(row[1]) else "0"
            info_d = str(row[3]).strip() if len(row) > 3 and not pd.isna(row[3]) and str(row[3]).lower() != 'nan' else None

            # --- RÁFAGA DENTRO DE LA GRILLA ---
            # Escribir SKU
            pyautogui.write(sku)
            
            # 2 ENTERS para llegar al campo de modo/unidad (según el nuevo menú)
            pyautogui.press('enter', presses=4, interval=0.1)

            if info_d:
                # MODO G (Pesables)
                pyautogui.write('g')
                pyautogui.write(cantidad)
                pyautogui.press('enter')
                time.sleep(0.4)
                pyautogui.write(info_d)
                pyautogui.press('enter') # Este enter lo deja en el SKU de abajo
            else:
                # MODO U (Unidad)
                # Escribimos u + cantidad pegados
                pyautogui.write(f"u{cantidad}")
                # UN SOLO ENTER: Confirma y PuTTY ya salta a la posición del siguiente SKU
                pyautogui.press('enter')
            
            time.sleep(0.1) # Pausa mínima de estabilidad antes del próximo SKU
            log_func(f"✅ Cargado: SKU {sku}")

        except Exception as row_err:
            log_func(f"❌ Error en fila {i+1}: {row_err}")
            
        progress_func((i + 1) / total)

    # 4. FINALIZACIÓN (Solo después de terminar todo el Excel)
    log_func("💾 Todos los artículos cargados. Guardando...")
    pyautogui.press('f5'); time.sleep(2.5)
    
    # Salida al menú principal
    for k in ['end', 'enter', 'end', 'end']:
        pyautogui.press(k)
        time.sleep(0.5)

    log_func("🏁 Proceso finalizado con éxito.")