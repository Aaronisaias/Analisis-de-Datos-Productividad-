# Reglas de limpieza

## Objetivo
Dejar los datos limpios y listos para analisis o exportacion.

## Reglas
- Usar Pandas para todo el proceso.
- Renombrar columnas si es necesario para que sean claras.
- Tratar valores nulos de forma explicita (`dropna` o imputacion).
- Eliminar columnas innecesarias para el analisis.
- Crear columnas derivadas solo cuando sean utiles.
- Excluir columnas ID en la salida final, salvo que el usuario pida lo contrario.
- Exportar a `Data_analysis.xlsx` solo cuando el usuario lo solicite.