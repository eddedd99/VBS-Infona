:: @echo off
pushd %~dp0
:: cscript "Inserta Total CL_CLIENTES Access.vbs"
:: cscript "Inserta Total CL_CLIENTES_DIR Access.vbs"
:: cscript "Inserta Total TOPAZ.AUX_IMSS_TRAB.vbs"
:: cscript "Inserta Total TOPAZ.TC_IMSS_TRAB.vbs"
:: cscript "Inserta Total ReporteSolicitudes.vbs"
:: cscript "Inserta Total ReporteSolicitudesHist.vbs"
:: cscript "Inserta Total PORTALCT.CL_PRE_AFILIACION_CT.vbs"

:: cscript "Inserta Total Tablas Credere.vbs"
:: cscript "Inserta Total TipoCliente CL_CLIENTES.vbs"
:: cscript "Inserta Total TipoCliente CL_CLIENTES 01 Estatus.vbs"
:: cscript "Inserta Total TipoCliente CL_CLIENTES 02 Estatus.vbs"
:: cscript "Inserta Total TipoCliente CL_CLIENTES x sucursal.vbs"
:: cscript "Inserta Total TipoCliente CL_CLIENTES fechaalta.vbs"
cscript "Inserta Total TipoCliente CL_CLIENTES diario imss.vbs"

:: shutdown /s

:: Para ejecutar de forma asincrona
:: start /b "" cscript "Inserta Total CL_CLIENTES Access.vbs"
:: cscript pfisicas.vbs 190079493a