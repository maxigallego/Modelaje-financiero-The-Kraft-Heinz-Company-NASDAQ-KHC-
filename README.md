# Financial Model: The Kraft Heinz Company (KHC)
### 3-Statement Integrated Model (FY2026E–FY2030E)

---

## 📌 Enfoque General

El modelo fue construido como un **modelo financiero integrado de tres estados** (*P&L, Balance Sheet y Cash Flow Statement*), con una proyección a 5 años (**FY2026E–FY2030E**) sobre una base histórica de 3 años (**FY2023A–FY2025A**).

*   **Fuente única de datos históricos:** Formulario 10-K de KHC correspondiente al ejercicio fiscal finalizado el 27 de diciembre de 2025, presentado ante la SEC el 12 de febrero de 2026.
*   **Filosofía de diseño:** Se priorizó la consistencia contable, la trazabilidad de los supuestos y la claridad estructural. Cada módulo de *buildup* alimenta directamente los tres estados financieros, y el Balance cierra correctamente en todos los períodos proyectados.

---

## 🗂️ Estructura del Modelo

El libro de Excel está organizado en **11 pestañas** distribuidas de manera lógica y secuencial:

| # | Pestaña | Descripción |
| :-: | :--- | :--- |
| 1 | `Cover` | Carátula, índice de hojas, convenciones de color y supuestos clave. |
| 2 | `Assumptions` | Centralización de todos los *drivers* y supuestos del modelo en una sola hoja. |
| 3 | `Revenue` | *Buildup* de ingresos proyectado por segmento (*North America, International Developed Markets, Emerging Markets*). |
| 4 | `COGS` | *Buildup* de costos de venta por segmento expresados como porcentaje de las ventas. |
| 5 | `OpEx` | Proyección de gastos de administración y ventas (SG&A), excluyendo *impairments*. |
| 6 | `NWC` | Capital de trabajo neto proyectado mediante días operativos (DSO, DIO y DPO). |
| 7 | `CapEx_DA` | Inversión en bienes de uso, depreciación, amortización y *roll-forward* de PPE e intangibles. |
| 8 | `Debt` | *Debt schedule* con el *roll-forward* detallado de la deuda y el gasto por intereses (*interest expense*). |
| 9 | `IS` | Estado de Resultados consolidado (*Income Statement*). |
| 10 | `BS` | Balance General consolidado (*Balance Sheet*). |
| 11 | `CFS` | Estado de Flujos de Efectivo consolidado (*Cash Flow Statement*). |

---

## 📈 Supuestos Clave (*Drivers*)

*   **Revenue:** Crecimiento anual por segmento aplicado sobre FY2025A.
    *   *North America:* Proyecta una recuperación gradual (-2% en FY2026E hasta +2,5% en FY2030E).
    *   *Emerging Markets:* Lidera el crecimiento con mayor dinamismo (+3% a +5%).
*   **COGS:** Proyectado como porcentaje de las ventas por segmento, asumiendo una mejora gradual de márgenes que refleja eficiencias operativas futuras.
*   **SG&A:** Crecimiento anual moderado (1,5%–2,5%). Los *impairments* de *goodwill* e intangibles se proyectaron en cero dado su carácter no recurrente.
*   **NWC:** Proyectado con días estables para cuentas por cobrar y pagar (DSO 33d, DPO 95d) y una leve optimización en inventarios (DIO de 69d → 67d).
*   **CapEx:** Proyectado como % de las ventas (~3,4% en FY2026E-FY2027E), reflejando un ciclo de reinversión moderado consistente con el promedio histórico FY2023A-FY2024A.
*   **Intangibles:** El *roll-forward* refleja únicamente la amortización recurrente (~$250M/año), sin incorporar nuevas adquisiciones, en línea con la etapa actual de reestructuración de portfolio de KHC.
*   **Deuda:** Refinanciación del 100% de los vencimientos programados (*debt rollover*) a una tasa de interés promedio proyectada del 4,6%.
*   **Impuestos:** Tasa efectiva proyectada en 25%, consistente con el *mix* de operaciones en EE. UU. y el exterior.
*   **Dividendos:** DPS mantenido en $1,60/acción, alineado con la política de distribución reciente de la compañía.

> ⚠️ **Nota sobre la Caja (Cash Plug):** El efectivo final se obtiene automáticamente como el *plug* proveniente del CFS. El saldo proyectado en el Balance corresponde estrictamente a *cash and cash equivalents*, excluyendo el *restricted cash* ($329M al cierre de FY2025A), dado que no está libremente disponible para operaciones ni para el servicio de la deuda.

---

## 🛠️ Herramientas Utilizadas

*   **Microsoft Excel:** Construcción del modelo financiero integrado y lógica matemática.
*   **Python:** Utilización de código para el diseño y la estructura visual del archivo.
*   **10-K FY2025 de The Kraft Heinz Company:** Fuente única y primaria de datos históricos (extraída vía SEC EDGAR).

---

## 📚 Fuentes de Referencia

*   [KHC 10-K FY2025 - SEC Filings](https://ir.kraftheinzcompany.com/financials/sec-filings)
*   *Wall Street Prep* – Financial Modeling Fundamentals
*   *Breaking Into Wall Street* – Three Statement Model Structure

---

## ✉️ Contacto

Ante cualquier consulta sobre el modelo, la mecánica de las pestañas o los supuestos adoptados, quedo a entera disposición.

*   **Autor:** Maximiliano Gallego
*   **Email:** maxigallego089@gmail.com
