import serial
import time
import pandas as pd
import re
import os

DEFAULT_PROMPT = b'#'
TIMEOUT_READ = 10

def leer_hasta_prompt(ser, prompt=DEFAULT_PROMPT, timeout=TIMEOUT_READ):
    start_time = time.time()
    buffer = b''
    while time.time() - start_time < timeout:
        if ser.in_waiting:
            data = ser.read(ser.in_waiting)
            buffer += data
            if buffer.endswith(prompt):
                return buffer.decode(errors="ignore")
        time.sleep(0.1)
    raise TimeoutError(f"Tiempo de espera agotado ({timeout}s). No se detectó el prompt '{prompt.decode()}'")

def enviar_y_esperar(ser, comando, prompt_esperado=DEFAULT_PROMPT):
    ser.write(f"{comando}\n".encode('ascii'))
    return leer_hasta_prompt(ser, prompt=prompt_esperado)

def obtener_modelo_serie(ser):
    print("Enviando 'show inventory'...")
    if ser.in_waiting:
        ser.read(ser.in_waiting)
    salida = enviar_y_esperar(ser, "show inventory", prompt_esperado=DEFAULT_PROMPT)
    regex_modelo = re.search(r"PID:\s*([\w\-/]+)", salida)
    regex_serie = re.search(r"SN:\s*([\w\d]+)", salida)
    modelo = regex_modelo.group(1).strip() if regex_modelo else None
    serie = regex_serie.group(1).strip() if regex_serie else None
    return modelo, serie

def configurar_dispositivo(ser, nombre, usuario, contrasena, dominio):
    print(f"Iniciando configuración para Hostname: {nombre}...")
    enviar_y_esperar(ser, "configure terminal", prompt_esperado=b'(config)#')
    
    comandos = [
        f"hostname {nombre}",
        f"username {usuario} privilege 15 password 0 {contrasena}",
        f"ip domain-name {dominio}",
    ]
    for cmd in comandos:
        enviar_y_esperar(ser, cmd, prompt_esperado=b'(config)#')

    print("Generando clave RSA...")
    enviar_y_esperar(ser, "crypto key generate rsa", prompt_esperado=b'modulus size')
    ser.write(b"1024\n")
    leer_hasta_prompt(ser, prompt=b'(config)#')

    extra_cmds = [
        "ip ssh version 2",
        "line console 0",
        "login local",
        "line vty 0 4",
        "login local",
        "transport input ssh",
        "transport output ssh",
    ]
    for cmd in extra_cmds:
        if "line " in cmd:
            enviar_y_esperar(ser, cmd, prompt_esperado=b'(config-line)#')
        else:
            enviar_y_esperar(ser, cmd, prompt_esperado=b'(config)#')

    enviar_y_esperar(ser, "end", prompt_esperado=DEFAULT_PROMPT)
    enviar_y_esperar(ser, "write memory", prompt_esperado=DEFAULT_PROMPT)
    print(f"✅ Configuración aplicada y guardada en {nombre}")
    return "Éxito en Configuración y Conexión"


def main():
    ruta_excel = r"C:\\Users\\OmEn\\Documents\\RedesP1.xlsx"
    if not os.path.exists(ruta_excel):
        print(f"❌ Error: No se encontró el archivo {ruta_excel}")
        return
    
    try:
        if ruta_excel.lower().endswith('.csv'):
            df = pd.read_csv(ruta_excel)
        else:
            try:
                df = pd.read_excel(ruta_excel)
            except ImportError:
                df = pd.read_excel(ruta_excel, engine='openpyxl')
    except Exception as e:
        print(f"❌ Error al leer el archivo: {e}")
        return
    
    column_mapping = {'usario': 'usuario', 'contraseña': 'contrasena'}
    df.rename(columns=column_mapping, inplace=True, errors='ignore')
    
    columnas_necesarias = {"modelo", "serie", "puerto", "baudios", "nombre", "usuario", "contrasena", "dominio"}
    if not columnas_necesarias.issubset(df.columns):
        raise ValueError(f"El Excel debe tener las columnas: {columnas_necesarias}. Encontradas: {list(df.columns)}")
    
    fila = df.iloc[0]
    puerto = fila["puerto"]
    baudios = int(str(fila["baudios"]).strip())
    
    try:
        print(f"🔌 Conectando a {puerto} con {baudios} baudios...")
        ser = serial.Serial(puerto, baudios, timeout=TIMEOUT_READ)
        time.sleep(1)
        
        ser.write(b'\n')
        time.sleep(0.5)
        ser.write(b'enable\n')
        
        try:
            leer_hasta_prompt(ser, prompt=DEFAULT_PROMPT, timeout=3)
        except TimeoutError:
            salida_inicial = ser.read_all().decode(errors="ignore")
            if "Password:" in salida_inicial or "password:" in salida_inicial:
                ser.write(f"{fila['contrasena']}\n".encode('ascii'))
                leer_hasta_prompt(ser, prompt=DEFAULT_PROMPT, timeout=TIMEOUT_READ)
        
        modelo_real, serie_real = obtener_modelo_serie(ser)
        print(f"📋 Router detectado -> Modelo: {modelo_real}, Serie: {serie_real}")
        
        match = df[(df["modelo"] == modelo_real) & (df["serie"] == serie_real)]
        
        if not match.empty:
            datos = match.iloc[0]
            print("✅ Coincidencia encontrada en Excel, aplicando configuración...")
            configurar_dispositivo(
                ser,
                datos["nombre"],
                datos["usuario"],
                datos["contrasena"],
                datos["dominio"]
            )
        else:
            print("⚠ El dispositivo detectado no está en el Excel. No se configurará.")
        
        ser.close()
    
    except Exception as e:
        print(f"❌ Error al conectar/configurar: {e}")

if __name__ == "__main__":
    main()