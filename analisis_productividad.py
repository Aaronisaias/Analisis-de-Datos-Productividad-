import pandas as pd


def cargar_datos(path_csv: str = "productividad_empleados.csv") -> pd.DataFrame:
    dt = pd.read_csv(path_csv)
    return dt


def preparar_rangos_edad(dt: pd.DataFrame) -> pd.DataFrame:
    dt = dt.copy()
    bins = [20, 30, 40, 50, 60]
    labels = ["20-29", "30-39", "40-49", "50-59"]
    dt["Rango_Edad"] = pd.cut(dt["Age"], bins=bins, labels=labels, right=False)
    return dt


def promedio_basico(dt: pd.DataFrame) -> dict:
    return {
        "sueno": float(dt["Sleep_Hours"].mean()),
        "horas_trabajo": float(dt["Work_Hours_Per_Day"].mean()),
        "estres": float(dt["Stress_Level"].mean()),
        "pantalla": float(dt["Screen_Time_Hours"].mean()),
        "reuniones": float(dt["Meetings_Per_Day"].mean()),
        "tareas": float(dt["Tasks_Completed"].mean()),
    }


def por_rango_edad(dt: pd.DataFrame) -> pd.DataFrame:
    return (
        dt.groupby("Rango_Edad", observed=True)
        .agg(
            Prom_Estres=("Stress_Level", "mean"),
            Prom_Productividad=("Productivity_Score", "mean"),
            Prom_Sueno=("Sleep_Hours", "mean"),
            Prom_Horas_Trabajo=("Work_Hours_Per_Day", "mean"),
            Prom_Tareas=("Tasks_Completed", "mean"),
            Prom_Reuniones=("Meetings_Per_Day", "mean"),
        )
        .round(2)
    )


def por_area(dt: pd.DataFrame) -> pd.DataFrame:
    return (
        dt.groupby("Department", observed=True)
        .agg(
            Prom_Sueno=("Sleep_Hours", "mean"),
            Prom_Estres=("Stress_Level", "mean"),
            Prom_Pantalla=("Screen_Time_Hours", "mean"),
            Prom_Horas_Trabajo=("Work_Hours_Per_Day", "mean"),
            Prom_Reuniones=("Meetings_Per_Day", "mean"),
            Prom_Tareas=("Tasks_Completed", "mean"),
            Prom_Productividad=("Productivity_Score", "mean"),
        )
        .round(2)
    )


def limpiar_para_excel(dt: pd.DataFrame) -> pd.DataFrame:
    # Renombrar columnas a espanol para el entregable final.
    mapa_columnas = {
        "Department": "Area",
        "Age": "Edad",
        "Work_Hours_Per_Day": "Horas_Trabajo_Dia",
        "Meetings_Per_Day": "Reuniones_Dia",
        "Screen_Time_Hours": "Horas_Pantalla",
        "Sleep_Hours": "Horas_Sueno",
        "Stress_Level": "Nivel_Estres",
        "Tasks_Completed": "Tareas_Completadas",
        "Productivity_Score": "Puntaje_Productividad",
    }

    # Regla de limpieza: excluir ID en salida final.
    columnas_salida = [c for c in mapa_columnas.keys() if c in dt.columns]
    limpio = dt[columnas_salida].copy().rename(columns=mapa_columnas)

    # Asegurar tipos numericos y formato simple.
    columnas_decimales = [
        "Horas_Trabajo_Dia",
        "Horas_Pantalla",
        "Horas_Sueno",
        "Puntaje_Productividad",
    ]
    for col in columnas_decimales:
        limpio[col] = pd.to_numeric(limpio[col], errors="coerce").round(2)

    columnas_enteras = ["Edad", "Reuniones_Dia", "Nivel_Estres", "Tareas_Completadas"]
    for col in columnas_enteras:
        limpio[col] = pd.to_numeric(limpio[col], errors="coerce")

    # Crear columna derivada util para analisis por edad.
    bins = [20, 30, 40, 50, 60]
    labels = ["20-29", "30-39", "40-49", "50-59"]
    limpio["Rango_Edad"] = pd.cut(limpio["Edad"], bins=bins, labels=labels, right=False)

    return limpio


def exportar_excel_limpio(dt: pd.DataFrame, path_xlsx: str = "Data_analysis.xlsx") -> None:
    limpio = limpiar_para_excel(dt)
    with pd.ExcelWriter(path_xlsx, engine="openpyxl") as writer:
        limpio.to_excel(writer, sheet_name="Datos_Limpios", index=False)


def guardar_informe_txt(dt: pd.DataFrame, path_txt: str = "analisis_productividad.txt") -> None:
    dt = preparar_rangos_edad(dt)
    p = promedio_basico(dt)
    edad = por_rango_edad(dt)
    area = por_area(dt)

    max_estres_rango = edad["Prom_Estres"].idxmax()
    max_estres_valor = float(edad.loc[max_estres_rango, "Prom_Estres"])
    max_prod_rango = edad["Prom_Productividad"].idxmax()
    min_prod_rango = edad["Prom_Productividad"].idxmin()

    area_horas = area["Prom_Horas_Trabajo"].idxmax()
    area_reuniones = area["Prom_Reuniones"].idxmax()
    area_tareas = area["Prom_Tareas"].idxmax()
    area_productividad = area["Prom_Productividad"].idxmax()

    # Rango de edad mas cercano a 8 horas de sueno y con productividad alta.
    edad_q7 = edad.copy()
    edad_q7["dist_8h"] = (edad_q7["Prom_Sueno"] - 8).abs()
    edad_q7 = edad_q7.sort_values(["dist_8h", "Prom_Productividad"], ascending=[True, False])
    mejor_rango_q7 = edad_q7.index[0]
    q7 = edad_q7.iloc[0]

    lines = []
    lines.append("ANALISIS DEL DATASET - RESPUESTAS A PREGUNTAS")
    lines.append("============================================")
    lines.append("")
    lines.append("Pregunta 1. Cual es el promedio de sueno, horas de trabajo, estres y tiempo de pantalla?")
    lines.append(f"- Sueno promedio: {p['sueno']:.2f} h")
    lines.append(f"- Horas de trabajo promedio: {p['horas_trabajo']:.2f} h")
    lines.append(f"- Estres promedio: {p['estres']:.2f}")
    lines.append(f"- Tiempo en pantalla promedio: {p['pantalla']:.2f} h")
    lines.append("")

    lines.append("Pregunta 2. Cual es el promedio de reuniones y de tareas completadas?")
    lines.append(f"- Reuniones promedio por dia: {p['reuniones']:.2f}")
    lines.append(f"- Tareas completadas promedio: {p['tareas']:.2f}")
    lines.append("")

    lines.append("Pregunta 3. Por rango de edad, cual es el promedio maximo de estres?")
    lines.append(f"- Rango con mayor estres promedio: {max_estres_rango} ({max_estres_valor:.2f})")
    lines.append("- Detalle de estres por rango de edad:")
    for idx, row in edad["Prom_Estres"].sort_values(ascending=False).items():
        lines.append(f"  - {idx}: {row:.2f}")
    lines.append("")

    lines.append("Pregunta 4. Cual es el rango de edad con mas y menos productividad promedio?")
    lines.append(
        f"- Mayor productividad promedio: {max_prod_rango} ({edad.loc[max_prod_rango, 'Prom_Productividad']:.2f})"
    )
    lines.append(
        f"- Menor productividad promedio: {min_prod_rango} ({edad.loc[min_prod_rango, 'Prom_Productividad']:.2f})"
    )
    lines.append("- Detalle de productividad por rango de edad:")
    for idx, row in edad["Prom_Productividad"].sort_values(ascending=False).items():
        lines.append(f"  - {idx}: {row:.2f}")
    lines.append("")

    lines.append(
        "Pregunta 5. Que area tiene mas horas de trabajo, reuniones, tareas cumplidas y puntaje de productividad?"
    )
    lines.append(f"- Area con mas horas de trabajo promedio: {area_horas} ({area.loc[area_horas, 'Prom_Horas_Trabajo']:.2f})")
    lines.append(f"- Area con mas reuniones promedio: {area_reuniones} ({area.loc[area_reuniones, 'Prom_Reuniones']:.2f})")
    lines.append(f"- Area con mas tareas promedio: {area_tareas} ({area.loc[area_tareas, 'Prom_Tareas']:.2f})")
    lines.append(
        f"- Area con mas productividad promedio: {area_productividad} ({area.loc[area_productividad, 'Prom_Productividad']:.2f})"
    )
    lines.append("")

    lines.append(
        "Pregunta 6. Cual es el promedio por area de sueno, estres, tiempo en pantalla, horas de trabajo y productividad?"
    )
    lines.append("Area,Prom_Sueno,Prom_Estres,Prom_Pantalla,Prom_Horas_Trabajo,Prom_Productividad")
    for dept_name, row in area.iterrows():
        lines.append(
            f"{dept_name},{row['Prom_Sueno']:.2f},{row['Prom_Estres']:.2f},"
            f"{row['Prom_Pantalla']:.2f},{row['Prom_Horas_Trabajo']:.2f},"
            f"{row['Prom_Productividad']:.2f}"
        )
    lines.append("")

    lines.append(
        "Pregunta 7. Por rango de edad, cual tiene promedio de sueno de 8 horas y productividad alta, y cuales son horas trabajadas, tareas y reuniones?"
    )
    lines.append("- En este dataset ningun rango llega a 8.00 h promedio de sueno.")
    lines.append(
        f"- El rango mas cercano a 8 horas y con productividad alta es: {mejor_rango_q7}"
    )
    lines.append(f"  - Sueno promedio: {q7['Prom_Sueno']:.2f} h")
    lines.append(f"  - Productividad promedio: {q7['Prom_Productividad']:.2f}")
    lines.append(f"  - Horas trabajadas promedio: {q7['Prom_Horas_Trabajo']:.2f} h")
    lines.append(f"  - Tareas cumplidas promedio: {q7['Prom_Tareas']:.2f}")
    lines.append(f"  - Reuniones promedio: {q7['Prom_Reuniones']:.2f}")
    lines.append("")

    lines.append("RECOMENDACIONES PARA SOLUCIONAR EL PROBLEMA")
    lines.append("===========================================")
    lines.append("- Revisar la calidad e impacto de tareas, no solo cantidad.")
    lines.append("- Ajustar KPIs a resultados de negocio, no solo esfuerzo.")
    lines.append("- Identificar areas con mayor carga de horas y reuniones para optimizar procesos.")
    lines.append("- Aplicar seguimiento especial al rango 50-59 por estres alto.")
    lines.append("- Mantener monitoreo por area y rango de edad para decisiones de RRHH.")
    lines.append("")

    lines.append("QUE MOSTRAR EN POWER BI")
    lines.append("=======================")
    lines.append("- Tarjetas KPI: sueno, horas, estres, pantalla, reuniones, tareas, productividad.")
    lines.append("- Barras por area: productividad, horas, reuniones y tareas promedio.")
    lines.append("- Barras por rango de edad: estres y productividad promedio.")
    lines.append("- Matriz por area con sueno, estres, pantalla, horas y productividad.")
    lines.append("- Scatter: horas trabajadas vs productividad, color por area.")

    with open(path_txt, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


if __name__ == "__main__":
    datos = cargar_datos("productividad_empleados.csv")
    guardar_informe_txt(datos, "analisis_productividad.txt")
    exportar_excel_limpio(datos, "Data_analysis.xlsx")

