# Generador de los planos de emplazamiento y distribución

Dibuja el PDF A3 (`../Planos_Nave2_medidas_archivo_A3.pdf`) y el DXF (`../Plano_Nave2_medidas_archivo.dxf`).

- `modelo.py`: medidas de la nave, zonas, máquinas, cuadros, recorridos y aparcamiento. **Aquí se cambian los datos.**
- `construir.py`: pasa el modelo a entidades de dibujo.
- `render.py`: dibuja las hojas A3 (cajetín y tablas) y exporta el DXF.

Uso: `python render.py` (requiere matplotlib). Los resultados salen en `out/`.
La geometría de la nave sale de `Nave 2.dwg` (planta en L de 53,5 × 26,5 m + ala de 19 × 38 m).
