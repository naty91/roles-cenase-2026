# Portal Roles CENASE — Conciliación Real v5

Corrección principal:
- NO calcula SECAP/IECE como 1% del total general.
- Usa la suma de `Valor CCC` que trae el Consolidado IESS, trabajador por trabajador.
- Aporte individual: suma real de la columna Individual IESS.
- Aporte patronal: suma real de la columna Patronal IESS.
- PLANI esperado = Individual + Patronal + Valor CCC del Consolidado.
- El Reporte de Planillas es la verificación del PAGO REAL.
- Si Consolidado y Planillas no coinciden, la conciliación muestra la diferencia y el pago contable toma el valor realmente pagado.
- DIVPRE, FONDOS, PLTJEM y EXTSALCY se mantienen separados.
- El detalle de planillas conserva número de planilla, fecha de generación, fecha de pago, vencimiento y mes pagado.
- Operativos, Gerentes y Administrativos continúan separados.
- Gratificaciones = Otros Ingresos del Rol.
- El asiento de devengo y el soporte/asiento de pago IESS permanecen separados.
