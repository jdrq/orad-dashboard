/* CONFIGURACION UNICA DEL AÑO — ORAD Dashboard.
   Es el UNICO lugar donde se cambia el año. Lo leen index.html (en el navegador)
   y los scripts de Python (scripts/config_orad.py).
   Regla: en enero se sube a 2027 SOLO despues de cerrar y confirmar diciembre (≈ 8 de enero).
   Mantener el bloque entre llaves como JSON valido (comillas dobles, sin comas finales). */
window.ORAD_CONFIG = {
  "anio_activo": 2026,
  "primer_anio_historico": 2015
};
window.ORAD_ANIO_NUM = window.ORAD_CONFIG.anio_activo;
window.ORAD_ANIO_PRIMER = window.ORAD_CONFIG.primer_anio_historico;
window.ORAD_ANIO = String(window.ORAD_ANIO_NUM);
