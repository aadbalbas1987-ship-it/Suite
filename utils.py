import pandas as pd
import ctypes
import pyautogui

def limpiar_sku(v):
    if pd.isna(v): return None
    try:
        return str(int(float(str(v).strip())))
    except: return None

def f_monto(v):
    if pd.isna(v): return "0.00"
    return "{:.2f}".format(float(str(v).replace(',', '.')))

def forzar_caps_off():
    hllDll = ctypes.WinDLL("User32.dll")
    if hllDll.GetKeyState(0x14) & 0x0001:
        pyautogui.press('capslock')