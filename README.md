# Monitor SO

Herramienta web (GitHub Pages, sin servidor) para dar seguimiento a la llegada de producto de las SOs de la región.

## Uso
1. Abre la página y arrastra los 4 Excel (se detectan solos): **SOs**, **TOs**, **Inventario agencias**, **Inventario Cedis**.
   Puedes actualizar solo uno; los demás se conservan.
2. TOs que llegan por correo y aún no salen en el Excel: pestaña **TOs manuales** → pega la tabla. Se borran solas cuando la TO aparece en el archivo de TOs.
3. Captura el comentario del gerente en cada SO. Se guarda por número de SO y se conserva entre cargas.

## Lógica de clasificación
Las SOs se ordenan por prioridad: **Pagado → Financiamiento → Apartado → Facturación pendiente → Pendiente de revisión → otros → Cerrada**; dentro de cada grupo, la de mayor antigüedad va primero. Cada unidad se asigna una sola vez:

| Paso | Si… | Clasificación |
|---|---|---|
| A | Hay unidad apta en el inventario de la agencia (se excluyen Garantía e Informativo) | Facturación inmediata pendiente |
| B | La SO tiene una TO vinculada (TO Asignada o Nota de una TO manual) con ese artículo | Tránsito – en camino / por embarcar |
| C | Hay una TO a la agencia con ese artículo | Tránsito – en camino (Recepción pendiente / IF enviada) o por embarcar |
| D | El cedis que la surte tiene "Disponible Hoy" | Carga de TO pendiente |
| E | Nada de lo anterior | Escalar a Planeación (dice si el cedis tiene tránsito de entrada) |

Cedis: todas las agencias → **15001 Chalco (CHA)**; **16006 Ecatepec → 15002 Tultepec (TUL)**. Se configura en `CFG` al inicio de `app.js`.

Alertas: departamento distinto al habitual de la agencia, TO asignada que trae otro artículo, TO "Recibida" sin unidad en inventario, inventario no apto, CD de la SO distinto al esperado.

## Datos
Todo se procesa y guarda **en tu navegador** (localStorage). Ningún dato se sube al repositorio. Usa **Respaldo ↓** con regularidad y **Importar respaldo** para cambiar de equipo o navegador.

Librería incluida: SheetJS 0.18.5 (`vendor/`, licencia Apache-2.0).
