# Scraper de Registros de Propiedades del Condado de Utah

Este es un script de Python que extrae datos del sitio web de Registros de Propiedades del Condado de Utah. El script busca una lista de apellidos latinos/hispanos, extrae los detalles de la propiedad y guarda los datos en varios formatos (CSV, Excel y PDF).

## Características

- Extrae el nombre del propietario, la dirección, la ciudad, el código postal y el estado de propiedad.
- Busca automáticamente una lista predefinida de apellidos latinos comunes.
- Guarda los datos extraídos en los siguientes formatos:
  - `utah_county_data.csv`
  - `utah_county_data.xlsx`
  - `utah_county_data.pdf`

## Requisitos

Para ejecutar el script, necesitas tener Python 3 y los siguientes paquetes de Python instalados:

- `requests`
- `beautifulsoup4`
- `openpyxl`
- `reportlab`

## Instalación

1.  Clona este repositorio o descarga los archivos.
2.  Instala las dependencias usando pip:

    ```bash
    pip install requests beautifulsoup4 openpyxl reportlab
    ```

## Uso

Para ejecutar el script, simplemente corre el siguiente comando en tu terminal:

```bash
python3 scraper.py
```

El script comenzará a extraer los datos y, una vez completado, encontrarás los archivos generados en el mismo directorio.

### Modificar la lista de apellidos

Si deseas cambiar la lista de apellidos que se buscan, puedes editar la lista `LATINO_LAST_NAMES` directamente en el archivo `scraper.py`.

```python
LATINO_LAST_NAMES = [
    "Garcia",
    "Rodriguez",
    "Martinez",
    "Hernandez",
    "Lopez",
    # ... y más
]
```

## Nota sobre el tiempo de ejecución

El script puede tardar un tiempo considerable en completarse, ya que realiza una gran cantidad de solicitudes web para obtener todos los datos. El número de registros a extraer está actualmente limitado para evitar tiempos de espera en ciertos entornos. Si necesitas extraer una lista muy grande de registros, es posible que debas ejecutar el script en un entorno sin límites de tiempo de ejecución estrictos.
