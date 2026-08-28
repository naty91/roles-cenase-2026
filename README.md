# Portal Roles CENASE v13 — Contabilidad corregida

Versión basada en v12, conservando los módulos y descargas PDF e incorporando la lógica contable validada para nómina e IESS.

## Pestaña Contabilidad

La APP genera tres asientos independientes:

1. **Devengo del Rol + beneficios acumulados**
   - Sueldos, sobretiempos, otros ingresos y beneficios pagados desde los Roles.
   - Aporte personal, anticipos, préstamos, otros descuentos y Neto a Recibir desde los Roles.
   - XIII, XIV y Fondo de Reserva acumulados dentro del mismo asiento.
   - Sueldos por Pagar = Neto a Recibir real del Rol.

2. **Provisión Patronal IESS + SECAP/IECE**
   - Patronal desde el Consolidado IESS.
   - SECAP/IECE desde Valor CCC del Consolidado IESS.
   - No duplica el aporte personal.

3. **Pago de Planillas IESS**
   - Aporte personal según Consolidado IESS.
   - Préstamos Quirografarios según DIVPRE pagado.
   - Patronal y SECAP/IECE según Consolidado.
   - Fondos de Reserva según lote principal pagado.
   - IESS por liquidar como diferencia de conciliación contra el total efectivamente pagado.
   - Haber a 2.1.7.5.7 Otros Impuestos, siguiendo el asiento entregado por CENASE.

## Patrón enero 2026 validado

- Asiento 1: Debe = Haber = **163,456.82**
- Asiento 2: Debe = Haber = **16,073.31**
- Asiento 3: Debe = Haber = **37,931.63**
- IESS por liquidar: **64.08**
- Beneficios acumulados validados enero 2026:
  - XIII: **1,027.40**
  - XIV: **241.00**
  - Fondo Reserva: **1,805.01**

Para otros meses, la APP mantiene cálculo dinámico con las reglas configuradas.
