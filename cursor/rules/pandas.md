# Reglas Pandas

## Objetivo
Hacer analisis claros y reutilizables con Python + Pandas.

## Reglas
- Importar siempre Pandas con `import pandas as pd`.
- Leer el CSV en una variable llamada `dt`.
- Usar nombres de variables claros.
- Evitar codigo repetido: crear funciones reutilizables.
- Ordenar resultados con `sort_values()` cuando aplique.
- Crear nuevas columnas solo si aportan valor al analisis.

## Ejemplo base
```python
import pandas as pd

dt = pd.read_csv("archivo.csv")

def promedio_por_grupo(columna_grupo, columna_valor, ascendente=False):
    return (
        dt.groupby(columna_grupo)[columna_valor]
        .mean()
        .sort_values(ascending=ascendente)
    )
```