# Generador Essar

Aplicación de escritorio para generar archivos Excel a partir de una plantilla y archivos en una carpeta de insumos. La plantilla va incluida en el ejecutable; no requiere Python para usarla.

## Uso (ejecutable)

1. Generar el ejecutable (una vez): ver [Construir el ejecutable](#construir-el-ejecutable).
2. Abrir **Generador Essar.exe**.
3. **Carpeta insumos:** elegir la carpeta con los Excel de entrada.
4. **Carpeta salidas:** elegir dónde guardar los resultados.
5. Pulsar **Generar**.

Los archivos se guardan en `salidas/<nombre_archivo_insumo>/` como `ESAR__{col1}__{col2}__{col3}.xlsx`.

## Uso (desarrollo con Python)

```bash
pip install -r requirements.txt
python app.py
```

Misma interfaz; la plantilla está en el repo y se usa por defecto.

## Contenido del repositorio

| Archivo | Descripción |
|---------|-------------|
| `app.py` | Interfaz de la aplicación. |
| `generador.py` | Lógica del generador Excel. |
| `GeneradorEssar.spec` | Configuración de PyInstaller. |
| `build.bat` | Genera el ejecutable (Windows). |
| `requirements.txt` | Dependencias (pandas, openpyxl). |
| `plantilla.xlsx` | Plantilla Excel incluida en el repo; se empaqueta en el .exe. |

El ejecutable (**Generador Essar.exe**) no se sube al repo; se genera localmente con `build.bat`.

## Construir el ejecutable

1. Python 3.10+ instalado.
2. Ejecutar:
   ```bat
   build.bat
   ```
   O manualmente:
   ```bat
   pip install -r requirements.txt
   pip install pyinstaller
   py -3 -m PyInstaller GeneradorEssar.spec --noconfirm --distpath . --workpath build
   ```
3. El .exe se crea en la misma carpeta (incluye la plantilla).

## Requisitos

- **Para usar el .exe:** solo Windows; no hace falta instalar nada.
- **Para desarrollar o construir:** Python 3.10+, pip.
