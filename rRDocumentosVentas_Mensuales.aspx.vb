'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRDocumentosVentas_Mensuales"
'-------------------------------------------------------------------------------------------'
Partial Class rRDocumentosVentas_Mensuales
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("   SELECT      ")
            loComandoSeleccionar.AppendLine("                     CASE WHEN Cotizaciones.Documento   <> '' AND Renglones_Cotizaciones.renglon <= 1 THEN 1 ELSE 0 END	AS Doc_Cot,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_Cotizaciones.Can_Art1          AS Can_Art1_Cot,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_Cotizaciones.Mon_Net           AS Mon_Net_Cot,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Ped,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Ped,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Ped,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Fac,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Fac,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Fac,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_DCli,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_DCli,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_DCli,      ")
            loComandoSeleccionar.AppendLine("                     datepart(year, Cotizaciones.Fec_Ini) AS Año,      ")
            loComandoSeleccionar.AppendLine("                     DATEPART(MONTH, Cotizaciones.Fec_Ini) AS Mes,      ")
            loComandoSeleccionar.AppendLine("                     Cotizaciones.Fec_Ini  AS Fecha      ")
            loComandoSeleccionar.AppendLine("   INTO #tempCotizaciones      ")
            loComandoSeleccionar.AppendLine("   FROM        Cotizaciones      ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Cotizaciones ON Renglones_Cotizaciones.Documento = Cotizaciones.Documento      ")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_Cotizaciones.Cod_Art      ")
            loComandoSeleccionar.AppendLine("   --ORDER BY Cotizaciones.Fec_Ini      ")
            loComandoSeleccionar.AppendLine("   WHERE       ")
            loComandoSeleccionar.AppendLine(" 			Cotizaciones.Fec_Ini            Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Art               Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep               Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Renglones_Cotizaciones.Cod_Alm  Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar               Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Cod_Cli            Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Cod_Ven            Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Status            IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_tip               Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_cla               Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Cod_Suc               Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro10Hasta)
            


            loComandoSeleccionar.AppendLine("   SELECT       ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Cot,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Cot,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)                AS Mon_Net_Cot,      ")
            loComandoSeleccionar.AppendLine("                     CASE WHEN Pedidos.Documento   <> ''  AND Renglones_Pedidos.renglon <= 1  THEN 1 ELSE 0 END AS Doc_Ped,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_Pedidos.Can_Art1         AS Can_Art1_Ped,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_Pedidos.Mon_Net          AS Mon_Net_Ped,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Fac,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Fac,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Fac,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_DCli,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_DCli,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_DCli,      ")
            loComandoSeleccionar.AppendLine("                     datepart(YEAR, Pedidos.Fec_Ini) AS Año,      ")
            loComandoSeleccionar.AppendLine("                     DATEPART(MONTH, Pedidos.Fec_Ini) AS Mes,      ")
            loComandoSeleccionar.AppendLine("                     Pedidos.Fec_Ini  AS Fecha      ")
            loComandoSeleccionar.AppendLine("   INTO #tempPedidos      ")
            loComandoSeleccionar.AppendLine("   FROM Pedidos      ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Pedidos ON Renglones_Pedidos.Documento = Pedidos.Documento      ")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_Pedidos.Cod_Art      ")
            loComandoSeleccionar.AppendLine("   --ORDER BY Pedidos.Fec_Ini      ")
            loComandoSeleccionar.AppendLine("   WHERE       ")
            loComandoSeleccionar.AppendLine(" 			Pedidos.Fec_Ini            Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Art               Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep               Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Renglones_Pedidos.Cod_Alm       Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar               Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Pedidos.Cod_Cli            Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Pedidos.Cod_Ven            Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Pedidos.Status            IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_tip               Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_cla               Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Pedidos.Cod_Suc               Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro10Hasta)


            loComandoSeleccionar.AppendLine("   SELECT       ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Cot,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Cot,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Cot,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Ped,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Ped,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Ped,      ")
            loComandoSeleccionar.AppendLine("                     CASE WHEN Facturas.Documento   <> ''  AND Renglones_Facturas.renglon <= 1 THEN 1 ELSE 0 END AS Doc_Fac,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_Facturas.Can_Art1        AS Can_Art1_Fac,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_Facturas.Mon_Net         AS Mon_Net_Fac,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_DCli,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_DCli,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_DCli,      ")
            loComandoSeleccionar.AppendLine("                     datepart(year, Facturas.Fec_Ini) AS Año,      ")
            loComandoSeleccionar.AppendLine("                     DATEPART(MONTH, Facturas.Fec_Ini) AS Mes,      ")
            loComandoSeleccionar.AppendLine("                     Facturas.Fec_Ini  AS Fecha      ")
            loComandoSeleccionar.AppendLine("   INTO #tempFacturas      ")
            loComandoSeleccionar.AppendLine("   FROM Facturas      ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento      ")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art      ")
            loComandoSeleccionar.AppendLine("   --ORDER BY Facturas.Fec_Ini      ")
            loComandoSeleccionar.AppendLine("   WHERE       ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Fec_Ini            Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Art               Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep               Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Renglones_Facturas.Cod_Alm      Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar               Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Cli            Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Ven            Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Status            IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_tip               Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_cla               Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Suc               Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro10Hasta)


            loComandoSeleccionar.AppendLine("   SELECT       ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Cot,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Cot,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Cot,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Ped,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Ped,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Ped,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Fac,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Fac,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Fac,      ")
            loComandoSeleccionar.AppendLine("                     CASE WHEN Devoluciones_Clientes.Documento   <> ''  AND Renglones_dClientes.renglon <= 1  THEN 1 ELSE 0 END	AS Doc_DCli,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_dClientes.Can_Art1       AS Can_Art1_DCli,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_dClientes.Mon_Net              AS Mon_Net_DCli,      ")
            loComandoSeleccionar.AppendLine("                     datepart(year, Devoluciones_Clientes.Fec_Ini) AS Año,      ")
            loComandoSeleccionar.AppendLine("                     DATEPART(MONTH, Devoluciones_Clientes.Fec_Ini) AS Mes,      ")
            loComandoSeleccionar.AppendLine("                     Devoluciones_Clientes.Fec_Ini  AS Fecha      ")
            loComandoSeleccionar.AppendLine("   INTO #tempDevoluciones      ")
            loComandoSeleccionar.AppendLine("   FROM Devoluciones_Clientes      ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_dClientes ON Renglones_dClientes.Documento = Devoluciones_Clientes.Documento      ")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_dClientes.Cod_Art      ")
            loComandoSeleccionar.AppendLine("   --ORDER BY Devoluciones_Clientes.Fec_Ini      ")
            loComandoSeleccionar.AppendLine("   WHERE       ")
            loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Fec_Ini            Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Art               Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep               Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Renglones_dClientes.Cod_Alm     Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar               Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Cod_Cli            Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Cod_Ven            Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Status            IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_tip               Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_cla               Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Cod_Suc               Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro10Hasta)


            loComandoSeleccionar.AppendLine("   SELECT Doc_Cot, Can_Art1_Cot, Mon_Net_Cot, Doc_Ped, Can_Art1_Ped, Mon_Net_Ped, Doc_Fac, Can_Art1_Fac, Mon_Net_Fac,      ")
            loComandoSeleccionar.AppendLine("   Doc_DCli, Can_Art1_DCli, Mon_Net_DCli, Año, Mes      ")
            loComandoSeleccionar.AppendLine("   FROM #tempCotizaciones      ")

            loComandoSeleccionar.AppendLine("   UNION ALL      ")

            loComandoSeleccionar.AppendLine("   SELECT Doc_Cot, Can_Art1_Cot, Mon_Net_Cot, Doc_Ped, Can_Art1_Ped, Mon_Net_Ped, Doc_Fac, Can_Art1_Fac, Mon_Net_Fac,      ")
            loComandoSeleccionar.AppendLine("   Doc_DCli, Can_Art1_DCli, Mon_Net_DCli, Año, Mes      ")
            loComandoSeleccionar.AppendLine("   FROM #tempPedidos      ")

            loComandoSeleccionar.AppendLine("   UNION ALL      ")

            loComandoSeleccionar.AppendLine("   SELECT Doc_Cot, Can_Art1_Cot, Mon_Net_Cot, Doc_Ped, Can_Art1_Ped, Mon_Net_Ped, Doc_Fac, Can_Art1_Fac, Mon_Net_Fac,      ")
            loComandoSeleccionar.AppendLine("   Doc_DCli, Can_Art1_DCli, Mon_Net_DCli, Año, Mes      ")
            loComandoSeleccionar.AppendLine("   FROM #tempFacturas      ")

            loComandoSeleccionar.AppendLine("   UNION ALL      ")

            loComandoSeleccionar.AppendLine("   SELECT       ")
            loComandoSeleccionar.AppendLine("   Doc_Cot, Can_Art1_Cot, Mon_Net_Cot, Doc_Ped, Can_Art1_Ped, Mon_Net_Ped, Doc_Fac, Can_Art1_Fac, Mon_Net_Fac,      ")
            loComandoSeleccionar.AppendLine("   Doc_DCli, Can_Art1_DCli, Mon_Net_DCli, Año, Mes      ")

            loComandoSeleccionar.AppendLine("   FROM #tempDevoluciones      ")
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento)


            loComandoSeleccionar.AppendLine("   SELECT ")
            loComandoSeleccionar.AppendLine("   		ISNULL(")
            loComandoSeleccionar.AppendLine("   				(")
            loComandoSeleccionar.AppendLine("   					(DATEDIFF(y,")
            loComandoSeleccionar.AppendLine("   									CASE ")
            loComandoSeleccionar.AppendLine("   										WHEN DATEPART(DAY, MIN(Minimo)) >=1  AND DATEPART(DAY, MIN(Minimo)) < 30 THEN CAST(DATEPART(YEAR, MIN(Minimo)) As varchar(4)) + ")
            loComandoSeleccionar.AppendLine("   																														CASE ")
            loComandoSeleccionar.AppendLine("   																																WHEN DATEPART(MONTH, MIN(Minimo)) < 10 THEN '0' + CAST(DATEPART(MONTH, MIN(Minimo)) As varchar(2)) ")
            loComandoSeleccionar.AppendLine("   																																ELSE CAST(DATEPART(MONTH, MIN(Minimo)) As varchar(2)) ")
            loComandoSeleccionar.AppendLine("   																														END + '01'")
            loComandoSeleccionar.AppendLine("   									END,")
            loComandoSeleccionar.AppendLine("   									CASE ")
            loComandoSeleccionar.AppendLine("   										WHEN DATEPART(DAY, MAX(Maximo)) >=1  AND DATEPART(DAY, MAX(Maximo)) < 30 THEN CAST(DATEPART(YEAR, MAX(Maximo)) As varchar(4)) + ")
            loComandoSeleccionar.AppendLine("   																														CASE ")
            loComandoSeleccionar.AppendLine("   																																WHEN DATEPART(MONTH, MAX(Maximo)) < 10 THEN '0' + CAST(DATEPART(MONTH, MAX(Maximo)) As varchar(2)) ")
            loComandoSeleccionar.AppendLine("   																																ELSE CAST(DATEPART(MONTH, MAX(Maximo)) As varchar(2)) ")
            loComandoSeleccionar.AppendLine("   																														END +")
            loComandoSeleccionar.AppendLine("   																														CASE ")
            loComandoSeleccionar.AppendLine("   																																WHEN DATEPART(MONTH, MAX(Maximo)) = 2 THEN '28'")
            loComandoSeleccionar.AppendLine("   																																ELSE '30'")
            loComandoSeleccionar.AppendLine("   																														END")
            loComandoSeleccionar.AppendLine("   									END")
            loComandoSeleccionar.AppendLine("   								) +")
            loComandoSeleccionar.AppendLine("   							CASE ")
            loComandoSeleccionar.AppendLine("   								WHEN  DATEPART(YEAR," & lcParametro0Desde & ") = DATEPART(YEAR," & lcParametro0Hasta & ") THEN ")
            loComandoSeleccionar.AppendLine("   									case ")
            loComandoSeleccionar.AppendLine("   										WHEN DATEPART(MONTH, " & lcParametro0Desde & ") >= 1 THEN 2")
            loComandoSeleccionar.AppendLine("   										else 0")
            loComandoSeleccionar.AppendLine("   									END")
            loComandoSeleccionar.AppendLine("   								ELSE")
            loComandoSeleccionar.AppendLine("   										(2 * DATEDIFF(YEAR, "& lcParametro0Desde &", "& lcParametro0Hasta &"))")
            loComandoSeleccionar.AppendLine("   							END")
            loComandoSeleccionar.AppendLine("   				) /30),0) AS Num_Meses, ")
                        
            
            loComandoSeleccionar.AppendLine("   	    ISNULL(DATEDIFF(y,MIN(Minimo), MAX(Maximo)),0) AS Num_Dias      ")
            loComandoSeleccionar.AppendLine("   FROM (      ")
            loComandoSeleccionar.AppendLine("   					   SELECT MIN(Fecha) AS Minimo, MAX(Fecha) AS Maximo      ")
            loComandoSeleccionar.AppendLine("   					   FROM #tempCotizaciones            ")
            loComandoSeleccionar.AppendLine("   					   UNION             ")
            loComandoSeleccionar.AppendLine("   					   SELECT MIN(Fecha) AS Minimo, MAX(Fecha) AS Maximo      ")
            loComandoSeleccionar.AppendLine("   					   FROM #tempPedidos            ")
            loComandoSeleccionar.AppendLine("   					   UNION       ")
            loComandoSeleccionar.AppendLine("   					   SELECT MIN(Fecha) AS Minimo, MAX(Fecha) AS Maximo      ")
            loComandoSeleccionar.AppendLine("   					   FROM #tempFacturas            ")
            loComandoSeleccionar.AppendLine("   					   UNION       ")
            loComandoSeleccionar.AppendLine("   					   SELECT MIN(Fecha) AS Minimo, MAX(Fecha) AS Maximo      ")
            loComandoSeleccionar.AppendLine("   					   FROM #tempDevoluciones       ")
            loComandoSeleccionar.AppendLine("   ) AS Tiempo      ")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRDocumentosVentas_Mensuales", laDatosReporte)
            If laDatosReporte.Tables(1).Rows(0).Item(0) = 0 Then
                laDatosReporte.Tables(1).Rows(0).Item(0) = 1
            End If

            If laDatosReporte.Tables(1).Rows(0).Item(1) = 0 Then
                laDatosReporte.Tables(1).Rows(0).Item(1) = 1
            End If

            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m1").Text = "CDBL (Sum({curReportes.Doc_Cot})/" & laDatosReporte.Tables(1).Rows(0).Item(0) & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m2").Text = " CDBL(Sum({curReportes.Can_Art1_Cot})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m3").Text = "CDBL (Sum({curReportes.Mon_Net_Cot})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m4").Text = "CDBL (Sum({curReportes.Doc_Ped})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m5").Text = "CDBL (Sum({curReportes.Can_Art1_Ped})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m6").Text = "CDBL (Sum({curReportes.Mon_Net_Ped})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m7").Text = "CDBL (Sum({curReportes.Doc_Fac})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m8").Text = "CDBL (Sum({curReportes.Can_Art1_Fac})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m9").Text = "CDBL (Sum({curReportes.Mon_Net_Fac})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m10").Text = "CDBL (Sum({curReportes.Doc_DCli})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m11").Text = "CDBL (Sum({curReportes.Can_Art1_DCli})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m12").Text = "CDBL (Sum({curReportes.Mon_Net_DCli})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"

            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d1").Text = "ToNumber (Sum({curReportes.Doc_Cot})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d2").Text = "ToNumber (Sum({curReportes.Can_Art1_Cot})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d3").Text = "ToNumber (Sum({curReportes.Mon_Net_Cot})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d4").Text = "ToNumber (Sum({curReportes.Doc_Ped})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d5").Text = "ToNumber (Sum({curReportes.Can_Art1_Ped})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d6").Text = "ToNumber (Sum({curReportes.Mon_Net_Ped})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d7").Text = "ToNumber (Sum({curReportes.Doc_Fac})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d8").Text = "ToNumber (Sum({curReportes.Can_Art1_Fac})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d9").Text = "ToNumber (Sum({curReportes.Mon_Net_Fac})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d10").Text = "ToNumber (Sum({curReportes.Doc_DCli})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d11").Text = "ToNumber (Sum({curReportes.Can_Art1_DCli})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d12").Text = "ToNumber (Sum({curReportes.Mon_Net_DCli})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"

            'loObjetoReporte.DataDefinition.FormulaFields("Promedio_d12").ValueType = CrystalDecisions.Shared.FieldValueType.NumberField

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRDocumentosVentas_Mensuales.ReportSource = loObjetoReporte


        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try
    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Try

            loObjetoReporte.Close()

        Catch loExcepcion As Exception

        End Try

    End Sub
End Class

'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' CMS: 01/07/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 13/07/09: Verificación de registros
'-------------------------------------------------------------------------------------------'
