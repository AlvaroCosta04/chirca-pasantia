from pathlib import Path

# Ruta exacta de la carpeta sincronizada en la PC de Pasantías (Windows)
RUTA_BASE = Path("C:/Users/Pasantias/OVC SA/GAMAS - GAMAS")

def mapear_estructura(directorio_raiz):
    if not directorio_raiz.exists():
        print(f"Error: La ruta '{directorio_raiz}' no existe.")
        print("Revisá si OneDrive ya está abierto y la carpeta sincronizada correctamente.")
        return

    print(f"Mapa de carpetas para: {directorio_raiz.name}\n" + "="*50)
    
    # Recorremos buscando carpetas (ignoramos archivos sueltos para no saturar)
    for ruta in sorted(directorio_raiz.rglob('*')):
        if ruta.is_dir():
            # Calculamos la profundidad para hacer la indentación
            profundidad = len(ruta.relative_to(directorio_raiz).parts)
            indentacion = "    " * profundidad
            print(f"{indentacion}📁 {ruta.name}")

if __name__ == "__main__":
    mapear_estructura(RUTA_BASE)
