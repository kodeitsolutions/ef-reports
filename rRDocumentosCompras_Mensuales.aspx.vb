'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRDocumentosCompras_Mensuales"
'-------------------------------------------------------------------------------------------'
Partial Class rRDocumentosCompras_Mensuales
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
            loComandoSeleccionar.AppendLine("                     CASE WHEN Requisiciones.Documento   <> '' AND Renglones_Requisiciones.renglon <= 1 THEN 1 ELSE 0 END	AS Doc_Req,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_Requisiciones.Can_Art1          AS Can_Art1_Req,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_Requisiciones.Mon_Net           AS Mon_Net_Req,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_OC,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_OC,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_OC,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Com,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Com,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Com,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_DPro,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_DPro,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_DPro,      ")
            loComandoSeleccionar.AppendLine("                     datepart(year, Requisiciones.Fec_Ini) AS Año,      ")
            loComandoSeleccionar.AppendLine("                     DATEPART(MONTH, Requisiciones.Fec_Ini) AS Mes,      ")
            loComandoSeleccionar.AppendLine("                     Requisiciones.Fec_Ini  AS Fecha      ")
            loComandoSeleccionar.AppendLine("   INTO #tempRequisiciones      ")
            loComandoSeleccionar.AppendLine("   FROM        Requisiciones      ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Requisiciones ON Renglones_Requisiciones.Documento = Requisiciones.Documento      ")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_Requisiciones.Cod_Art      ")
            loComandoSeleccionar.AppendLine("   --ORDER BY Requisiciones.Fec_Ini      ")
            loComandoSeleccionar.AppendLine("   WHERE       ")
            loComandoSeleccionar.AppendLine(" 			Requisiciones.Fec_Ini            Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Art               Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep               Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Renglones_Requisiciones.Cod_Alm  Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar               Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Requisiciones.Cod_Pro            Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Requisiciones.Cod_Ven            Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Requisiciones.Status            IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_tip               Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_cla               Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Requisiciones.Cod_Suc               Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro10Hasta)
            


            loComandoSeleccionar.AppendLine("   SELECT       ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Req,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Req,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)                AS Mon_Net_Req,      ")
            loComandoSeleccionar.AppendLine("                     CASE WHEN Ordenes_Compras.Documento   <> ''  AND Renglones_oCompras.renglon <= 1  THEN 1 ELSE 0 END AS Doc_OC,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_oCompras.Can_Art1         AS Can_Art1_OC,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_oCompras.Mon_Net          AS Mon_Net_OC,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Com,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Com,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Com,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_DPro,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_DPro,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_DPro,      ")
            loComandoSeleccionar.AppendLine("                     datepart(YEAR, Ordenes_Compras.Fec_Ini) AS Año,      ")
            loComandoSeleccionar.AppendLine("                     DATEPART(MONTH, Ordenes_Compras.Fec_Ini) AS Mes,      ")
            loComandoSeleccionar.AppendLine("                     Ordenes_Compras.Fec_Ini  AS Fecha      ")
            loComandoSeleccionar.AppendLine("   INTO #tempOrdenes_Compras      ")
            loComandoSeleccionar.AppendLine("   FROM Ordenes_Compras      ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_oCompras ON Renglones_oCompras.Documento = Ordenes_Compras.Documento      ")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_oCompras.Cod_Art      ")
            loComandoSeleccionar.AppendLine("   --ORDER BY Ordenes_Compras.Fec_Ini      ")
            loComandoSeleccionar.AppendLine("   WHERE       ")
            loComandoSeleccionar.AppendLine(" 			Ordenes_Compras.Fec_Ini            Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Art               Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep               Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Renglones_oCompras.Cod_Alm       Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar               Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Ordenes_Compras.Cod_Pro            Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Ordenes_Compras.Cod_Ven            Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Ordenes_Compras.Status            IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_tip               Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_cla               Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Ordenes_Compras.Cod_Suc               Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro10Hasta)


            loComandoSeleccionar.AppendLine("   SELECT       ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Req,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Req,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Req,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_OC,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_OC,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_OC,      ")
            loComandoSeleccionar.AppendLine("                     CASE WHEN Compras.Documento   <> ''  AND Renglones_Compras.renglon <= 1 THEN 1 ELSE 0 END AS Doc_Com,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_Compras.Can_Art1        AS Can_Art1_Com,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_Compras.Mon_Net         AS Mon_Net_Com,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_DPro,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_DPro,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_DPro,      ")
            loComandoSeleccionar.AppendLine("                     datepart(year, Compras.Fec_Ini) AS Año,      ")
            loComandoSeleccionar.AppendLine("                     DATEPART(MONTH, Compras.Fec_Ini) AS Mes,      ")
            loComandoSeleccionar.AppendLine("                     Compras.Fec_Ini  AS Fecha      ")
            loComandoSeleccionar.AppendLine("   INTO #tempCompras      ")
            loComandoSeleccionar.AppendLine("   FROM Compras      ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Compras ON Renglones_Compras.Documento = Compras.Documento      ")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_Compras.Cod_Art      ")
            loComandoSeleccionar.AppendLine("   --ORDER BY Compras.Fec_Ini      ")
            loComandoSeleccionar.AppendLine("   WHERE       ")
            loComandoSeleccionar.AppendLine(" 			Compras.Fec_Ini            Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Art               Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep               Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Renglones_Compras.Cod_Alm      Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar               Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Pro            Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Ven            Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Status            IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_tip               Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_cla               Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Suc               Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro10Hasta)


            loComandoSeleccionar.AppendLine("   SELECT       ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Req,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Req,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Req,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_OC,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_OC,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_OC,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0 AS CHAR(10))                AS Doc_Com,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Can_Art1_Com,      ")
            loComandoSeleccionar.AppendLine("                     CAST(0.0 AS DECIMAL)               AS Mon_Net_Com,      ")
            loComandoSeleccionar.AppendLine("                     CASE WHEN Devoluciones_Proveedores.Documento   <> ''  AND Renglones_dProveedores.renglon <= 1  THEN 1 ELSE 0 END	AS Doc_DPro,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_dProveedores.Can_Art1       AS Can_Art1_DPro,      ")
            loComandoSeleccionar.AppendLine("                     Renglones_dProveedores.Mon_Net              AS Mon_Net_DPro,      ")
            loComandoSeleccionar.AppendLine("                     datepart(year, Devoluciones_Proveedores.Fec_Ini) AS Año,      ")
            loComandoSeleccionar.AppendLine("                     DATEPART(MONTH, Devoluciones_Proveedores.Fec_Ini) AS Mes,      ")
            loComandoSeleccionar.AppendLine("                     Devoluciones_Proveedores.Fec_Ini  AS Fecha      ")
            loComandoSeleccionar.AppendLine("   INTO #tempDevoluciones      ")
            loComandoSeleccionar.AppendLine("   FROM Devoluciones_Proveedores      ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_dProveedores ON Renglones_dProveedores.Documento = Devoluciones_Proveedores.Documento      ")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_dProveedores.Cod_Art      ")
            loComandoSeleccionar.AppendLine("   --ORDER BY Devoluciones_Proveedores.Fec_Ini      ")
            loComandoSeleccionar.AppendLine("   WHERE       ")
            loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Fec_Ini            Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Art               Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep               Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Renglones_dProveedores.Cod_Alm     Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar               Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Proveedores.Cod_Pro            Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Proveedores.Cod_Ven            Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Proveedores.Status            IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_tip               Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_cla               Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Proveedores.Cod_Suc               Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro10Hasta)


            loComandoSeleccionar.AppendLine("   SELECT Doc_Req, Can_Art1_Req, Mon_Net_Req, Doc_OC, Can_Art1_OC, Mon_Net_OC, Doc_Com, Can_Art1_Com, Mon_Net_Com,      ")
            loComandoSeleccionar.AppendLine("   Doc_DPro, Can_Art1_DPro, Mon_Net_DPro, Año, Mes      ")
            loComandoSeleccionar.AppendLine("   FROM #tempRequisiciones      ")

            loComandoSeleccionar.AppendLine("   UNION ALL      ")

            loComandoSeleccionar.AppendLine("   SELECT Doc_Req, Can_Art1_Req, Mon_Net_Req, Doc_OC, Can_Art1_OC, Mon_Net_OC, Doc_Com, Can_Art1_Com, Mon_Net_Com,      ")
            loComandoSeleccionar.AppendLine("   Doc_DPro, Can_Art1_DPro, Mon_Net_DPro, Año, Mes      ")
            loComandoSeleccionar.AppendLine("   FROM #tempOrdenes_Compras      ")

            loComandoSeleccionar.AppendLine("   UNION ALL      ")

            loComandoSeleccionar.AppendLine("   SELECT Doc_Req, Can_Art1_Req, Mon_Net_Req, Doc_OC, Can_Art1_OC, Mon_Net_OC, Doc_Com, Can_Art1_Com, Mon_Net_Com,      ")
            loComandoSeleccionar.AppendLine("   Doc_DPro, Can_Art1_DPro, Mon_Net_DPro, Año, Mes      ")
            loComandoSeleccionar.AppendLine("   FROM #tempCompras      ")

            loComandoSeleccionar.AppendLine("   UNION ALL      ")

            loComandoSeleccionar.AppendLine("   SELECT       ")
            loComandoSeleccionar.AppendLine("   Doc_Req, Can_Art1_Req, Mon_Net_Req, Doc_OC, Can_Art1_OC, Mon_Net_OC, Doc_Com, Can_Art1_Com, Mon_Net_Com,      ")
            loComandoSeleccionar.AppendLine("   Doc_DPro, Can_Art1_DPro, Mon_Net_DPro, Año, Mes      ")

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
            loComandoSeleccionar.AppendLine("   					   FROM #tempRequisiciones            ")
            loComandoSeleccionar.AppendLine("   					   UNION             ")
            loComandoSeleccionar.AppendLine("   					   SELECT MIN(Fecha) AS Minimo, MAX(Fecha) AS Maximo      ")
            loComandoSeleccionar.AppendLine("   					   FROM #tempOrdenes_Compras            ")
            loComandoSeleccionar.AppendLine("   					   UNION       ")
            loComandoSeleccionar.AppendLine("   					   SELECT MIN(Fecha) AS Minimo, MAX(Fecha) AS Maximo      ")
            loComandoSeleccionar.AppendLine("   					   FROM #tempCompras            ")
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRDocumentosCompras_Mensuales", laDatosReporte)
            If laDatosReporte.Tables(1).Rows(0).Item(0) = 0 Then
                laDatosReporte.Tables(1).Rows(0).Item(0) = 1
            End If

            If laDatosReporte.Tables(1).Rows(0).Item(1) = 0 Then
                laDatosReporte.Tables(1).Rows(0).Item(1) = 1
            End If

            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m1").Text = "CDBL (Sum({curReportes.Doc_Req})/" & laDatosReporte.Tables(1).Rows(0).Item(0) & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m2").Text = " CDBL(Sum({curReportes.Can_Art1_Req})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m3").Text = "CDBL (Sum({curReportes.Mon_Net_Req})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m4").Text = "CDBL (Sum({curReportes.Doc_OC})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m5").Text = "CDBL (Sum({curReportes.Can_Art1_OC})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m6").Text = "CDBL (Sum({curReportes.Mon_Net_OC})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m7").Text = "CDBL (Sum({curReportes.Doc_Com})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m8").Text = "CDBL (Sum({curReportes.Can_Art1_Com})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m9").Text = "CDBL (Sum({curReportes.Mon_Net_Com})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m10").Text = "CDBL (Sum({curReportes.Doc_DPro})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m11").Text = "CDBL (Sum({curReportes.Can_Art1_DPro})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_m12").Text = "CDBL (Sum({curReportes.Mon_Net_DPro})/" & laDatosReporte.Tables(1).Rows(0).Item(0).ToString & " )"

            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d1").Text = "ToNumber (Sum({curReportes.Doc_Req})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d2").Text = "ToNumber (Sum({curReportes.Can_Art1_Req})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d3").Text = "ToNumber (Sum({curReportes.Mon_Net_Req})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d4").Text = "ToNumber (Sum({curReportes.Doc_OC})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d5").Text = "ToNumber (Sum({curReportes.Can_Art1_OC})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d6").Text = "ToNumber (Sum({curReportes.Mon_Net_OC})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d7").Text = "ToNumber (Sum({curReportes.Doc_Com})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d8").Text = "ToNumber (Sum({curReportes.Can_Art1_Com})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d9").Text = "ToNumber (Sum({curReportes.Mon_Net_Com})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d10").Text = "ToNumber (Sum({curReportes.Doc_DPro})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d11").Text = "ToNumber (Sum({curReportes.Can_Art1_DPro})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"
            loObjetoReporte.DataDefinition.FormulaFields("Promedio_d12").Text = "ToNumber (Sum({curReportes.Mon_Net_DPro})/" & laDatosReporte.Tables(1).Rows(0).Item(1).ToString & " )"

            'loObjetoReporte.DataDefinition.FormulaFields("Promedio_d12").ValueType = CrystalDecisions.Shared.FieldValueType.NumberField

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRDocumentosCompras_Mensuales.ReportSource = loObjetoReporte


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
' CMS: 03/05/10: Programacion inicial
'-------------------------------------------------------------------------------------------'