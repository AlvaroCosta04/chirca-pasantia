from pathlib import Path

# PASO 1: Reemplazá esto por la ruta donde OneDrive sincronizó la carpeta de la empresa.
# Ejemplo en Windows: Path("C:/Users/Sofia/Chirca OVC S.A/Gamas - Documentos")
RUTA_BASE = Path("C:/Users/TuCompanera/Ruta_A_La_Carpeta_Sincronizada")

def mapear_estructura(directorio_raiz):
    if not directorio_raiz.exists():
        print(f"Error: La ruta {directorio_raiz} no existe.")
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
