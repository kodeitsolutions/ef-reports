'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTrazabilidad_CyT_TipoArticulo"
'-------------------------------------------------------------------------------------------'
Partial Class rTrazabilidad_CyT_TipoArticulo 
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0)) 'Articulo
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia) 'Fecha
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2)) 'Proveedor compra
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3)) 'Vendedor compra
			Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4)) 'Departamento
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5)) 'Sección
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6)) 'Marca
			Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7)) 'Almacen compra
			Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
			Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8)) 'Revisión Compra
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
			Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9)) 'Sucursal Compra
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10)) 'Almacen Origen
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11)) 'Almacen Destino
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12)) 'Sucursal Traslado
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13)) 'Revision Traslado
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
			
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            
			Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpOperaciones(   Fecha               DATETIME,")
            loConsulta.AppendLine("                                Articulo_Cod_Tip    CHAR(10),")
            loConsulta.AppendLine("                                Articulo_Codigo     CHAR(30),")
            loConsulta.AppendLine("                                Articulo_Nombre     VARCHAR(100),")
            loConsulta.AppendLine("                                Compra_Fecha        DATETIME,")
            loConsulta.AppendLine("                                Compra_Proveedor    VARCHAR(100),")
            loConsulta.AppendLine("                                Compra_Documento    CHAR(10),")
            loConsulta.AppendLine("                                Compra_Renglon      INT,")
            loConsulta.AppendLine("                                Compra_Almacen      CHAR(10),")
            loConsulta.AppendLine("                                Compra_Unidad       CHAR(10),")
            loConsulta.AppendLine("                                Compra_Cantidad     DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Compra_Precio       DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Compra_Bruto        DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Compra_Impuesto     DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Compra_Neto         DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Traslado_Fecha      DATETIME,")
            loConsulta.AppendLine("                                Traslado_Origen     CHAR(10),")
            loConsulta.AppendLine("                                Traslado_Destino    CHAR(10),")
            loConsulta.AppendLine("                                Traslado_Documento  CHAR(10),")
            loConsulta.AppendLine("                                Traslado_Renglon    INT,")
            loConsulta.AppendLine("                                Traslado_Unidad     CHAR(10),")
            loConsulta.AppendLine("                                Traslado_Cantidad   DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Traslado_Costo      DECIMAL(28, 10)")
            loConsulta.AppendLine("                            );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Busca las compras")
            loConsulta.AppendLine("INSERT INTO #tmpOperaciones( Fecha, Articulo_Cod_Tip, Articulo_Codigo, Articulo_Nombre,")
            loConsulta.AppendLine("                             Compra_Fecha, Compra_Proveedor, Compra_Documento, Compra_Renglon, ")
            loConsulta.AppendLine("                             Compra_Almacen, Compra_Unidad, Compra_Cantidad, Compra_Precio, ")
            loConsulta.AppendLine("                             Compra_Bruto, Compra_Impuesto, Compra_Neto)")
            loConsulta.AppendLine("SELECT      Compras.Fec_Ini                 AS Fecha,")
            loConsulta.AppendLine("            Articulos.Cod_Tip               AS Articulo_Cod_Tip,")
            loConsulta.AppendLine("            Articulos.Cod_Art               AS Articulo_Codigo,")
            loConsulta.AppendLine("            Articulos.Nom_Art               AS Articulo_Nombre,")
            loConsulta.AppendLine("            Compras.Fec_Ini                 AS Compra_Fecha,")
            loConsulta.AppendLine("            Proveedores.Nom_Pro             AS Compra_Proveedor,")
            loConsulta.AppendLine("            Renglones_Compras.Documento     AS Compra_Documento,")
            loConsulta.AppendLine("            Renglones_Compras.Renglon       AS Compra_Renglon,")
            loConsulta.AppendLine("            Renglones_Compras.Cod_Alm       AS Compra_Almacen,")
            loConsulta.AppendLine("            Renglones_Compras.Cod_Uni       AS Compra_Unidad,")
            loConsulta.AppendLine("            Renglones_Compras.Can_Art1      AS Compra_Cantidad,")
            loConsulta.AppendLine("            Renglones_Compras.Precio1       AS Compra_Precio,")
            loConsulta.AppendLine("            Renglones_Compras.Mon_Net       AS Compra_Bruto, ")
            loConsulta.AppendLine("            Renglones_Compras.Mon_Imp1      AS Compra_Impuesto, ")
            loConsulta.AppendLine("            (Renglones_Compras.Mon_Net      ")
            loConsulta.AppendLine("              + Renglones_Compras.Mon_Imp1) AS Compra_Neto")
            loConsulta.AppendLine("FROM        Renglones_Compras")
            loConsulta.AppendLine("    JOIN    Compras ")
            loConsulta.AppendLine("        ON  Compras.Documento = Renglones_Compras.Documento")
            loConsulta.AppendLine("        AND Compras.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loConsulta.AppendLine("    JOIN    Articulos")
            loConsulta.AppendLine("        ON  Articulos.Cod_Art = Renglones_Compras.Cod_Art")
            loConsulta.AppendLine("    JOIN    Proveedores")
            loConsulta.AppendLine("        ON  Proveedores.Cod_Pro = Compras.Cod_Pro")
            loConsulta.AppendLine("WHERE       Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("        AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Compras.Fec_Ini BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("        AND " & lcParametro1Hasta)
            loConsulta.AppendLine("        AND Compras.Cod_Pro BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("        AND " & lcParametro2Hasta)
            loConsulta.AppendLine("        AND Compras.Cod_Ven BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("        AND " & lcParametro3Hasta)
            loConsulta.AppendLine("        AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("        AND " & lcParametro4Hasta)
            loConsulta.AppendLine("        AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("        AND " & lcParametro5Hasta)
            loConsulta.AppendLine("        AND Articulos.Cod_Mar BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("        AND " & lcParametro6Hasta)
            loConsulta.AppendLine("        AND Renglones_Compras.Cod_Alm BETWEEN " & lcParametro7Desde)
            loConsulta.AppendLine("        AND " & lcParametro7Hasta)
            loConsulta.AppendLine("        AND Compras.Cod_Rev BETWEEN " & lcParametro8Desde)
            loConsulta.AppendLine("        AND " & lcParametro8Hasta)
            loConsulta.AppendLine("        AND Compras.Cod_Suc BETWEEN " & lcParametro9Desde)
            loConsulta.AppendLine("        AND " & lcParametro9Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Busca los traslados")
            loConsulta.AppendLine("INSERT INTO #tmpOperaciones( Fecha, Articulo_Cod_Tip, Articulo_Codigo, Articulo_Nombre,")
            loConsulta.AppendLine("                             Traslado_Fecha, Traslado_Origen, Traslado_Destino, ")
            loConsulta.AppendLine("                             Traslado_Documento, Traslado_Renglon, ")
            loConsulta.AppendLine("                             Traslado_Unidad, Traslado_Cantidad, Traslado_Costo)")
            loConsulta.AppendLine("SELECT      Traslados.Fec_Ini                        AS Fecha,")
            loConsulta.AppendLine("            Articulos.Cod_Tip                        AS Articulo_Cod_Tip,")
            loConsulta.AppendLine("            Articulos.Cod_Art                        AS Articulo_Codigo,")
            loConsulta.AppendLine("            Articulos.Nom_Art                        AS Articulo_Nombre,")
            loConsulta.AppendLine("            Traslados.Fec_Ini                        AS Traslado_Fecha,")
            loConsulta.AppendLine("            Traslados.Alm_Ori                        AS Traslado_Origen,")
            loConsulta.AppendLine("            Traslados.Alm_Des                        AS Traslado_Destino,")
            loConsulta.AppendLine("            Renglones_Traslados.Documento            AS Traslado_Documento,")
            loConsulta.AppendLine("            Renglones_Traslados.Renglon              AS Traslado_Renglon,")
            loConsulta.AppendLine("            Renglones_Traslados.Cod_Uni              AS Traslado_Unidad,")
            loConsulta.AppendLine("            Renglones_Traslados.Can_Art1             AS Traslado_Cantidad,")
            loConsulta.AppendLine("            ROUND(Renglones_Traslados.Cos_Ult1       ")
            loConsulta.AppendLine("                *Renglones_Traslados.Can_Art1, 2)    AS Traslado_Costo")
            loConsulta.AppendLine("FROM        Traslados")
            loConsulta.AppendLine("    JOIN    Renglones_Traslados")
            loConsulta.AppendLine("        ON  Renglones_Traslados.Documento = Traslados.Documento")
            loConsulta.AppendLine("        AND Traslados.Status IN ('Confirmado', 'Procesado')")
            loConsulta.AppendLine("    JOIN    Articulos")
            loConsulta.AppendLine("        ON  Articulos.Cod_Art = Renglones_Traslados.Cod_Art")
            loConsulta.AppendLine("WHERE       Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("        AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Traslados.Fec_Ini BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("        AND " & lcParametro1Hasta)
            loConsulta.AppendLine("        AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("        AND " & lcParametro4Hasta)
            loConsulta.AppendLine("        AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("        AND " & lcParametro5Hasta)
            loConsulta.AppendLine("        AND Articulos.Cod_Mar BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("        AND " & lcParametro6Hasta)
            loConsulta.AppendLine("        AND Traslados.Alm_Ori BETWEEN " & lcParametro10Desde)
            loConsulta.AppendLine("        AND " & lcParametro10Hasta)
            loConsulta.AppendLine("        AND Traslados.Alm_Des BETWEEN " & lcParametro11Desde)
            loConsulta.AppendLine("        AND " & lcParametro11Hasta)
            loConsulta.AppendLine("        AND Traslados.Cod_Rev BETWEEN " & lcParametro12Desde)
            loConsulta.AppendLine("        AND " & lcParametro12Hasta)
            loConsulta.AppendLine("        AND Traslados.Cod_Suc BETWEEN " & lcParametro13Desde)
            loConsulta.AppendLine("        AND " & lcParametro13Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Cuenta las compras y traslados según sus estatus")
            loConsulta.AppendLine("CREATE TABLE #tmpConteo(    N_Compras_Total INT,  ")
            loConsulta.AppendLine("                            N_Compras_Anuladas INT,")
            loConsulta.AppendLine("                            N_Compras_Pendientes INT,")
            loConsulta.AppendLine("                            N_Compras_Confirmadas INT,")
            loConsulta.AppendLine("                            N_Compras_Afectadas INT,")
            loConsulta.AppendLine("                            N_Compras_Procesadas INT,")
            loConsulta.AppendLine("                            N_Traslados_Total INT,")
            loConsulta.AppendLine("                            N_Traslados_Anulados INT,")
            loConsulta.AppendLine("                            N_Traslados_Pendientes INT,")
            loConsulta.AppendLine("                            N_Traslados_Confirmados INT,")
            loConsulta.AppendLine("                            N_Traslados_Procesados INT")
            loConsulta.AppendLine("                        );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpConteo( N_Compras_Total, N_Compras_Anuladas, N_Compras_Pendientes, N_Compras_Confirmadas, ")
            loConsulta.AppendLine("                        N_Compras_Afectadas, N_Compras_Procesadas,")
            loConsulta.AppendLine("                        N_Traslados_Total, N_Traslados_Anulados, N_Traslados_Pendientes, ")
            loConsulta.AppendLine("                        N_Traslados_Confirmados, N_Traslados_Procesados)")
            loConsulta.AppendLine("SELECT  Conteo_Compras.Total,")
            loConsulta.AppendLine("        Conteo_Compras.Compras_Anuladas, ")
            loConsulta.AppendLine("        Conteo_Compras.Compras_Pendientes, ")
            loConsulta.AppendLine("        Conteo_Compras.Compras_Confirmadas, ")
            loConsulta.AppendLine("        Conteo_Compras.Compras_Afectadas, ")
            loConsulta.AppendLine("        Conteo_Compras.Compras_Procesadas,")
            loConsulta.AppendLine("        Conteo_Traslados.Total,")
            loConsulta.AppendLine("        Conteo_Traslados.Traslados_Anulados,")
            loConsulta.AppendLine("        Conteo_Traslados.Traslados_Pendientes,")
            loConsulta.AppendLine("        Conteo_Traslados.Traslados_Confirmados, ")
            loConsulta.AppendLine("        Conteo_Traslados.Traslados_Procesados")
            loConsulta.AppendLine("FROM    (   SELECT      COUNT(*) Total,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Compras.Status = 'Anulado' THEN 1 ELSE 0 END)     AS Compras_Anuladas,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Compras.Status = 'Pendiente' THEN 1 ELSE 0 END)   AS Compras_Pendientes,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Compras.Status = 'Confirmado' THEN 1 ELSE 0 END)  AS Compras_Confirmadas,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Compras.Status = 'Afectado' THEN 1 ELSE 0 END)    AS Compras_Afectadas,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Compras.Status = 'Procesado' THEN 1 ELSE 0 END)   AS Compras_Procesadas")
            loConsulta.AppendLine("            FROM        Compras")
            loConsulta.AppendLine("                JOIN    (   SELECT  Renglones_Compras.Documento ")
            loConsulta.AppendLine("                            FROM    Renglones_Compras ")
            loConsulta.AppendLine("                            JOIN    Articulos")
            loConsulta.AppendLine("                                ON  Articulos.Cod_Art = Renglones_Compras.Cod_Art")
            loConsulta.AppendLine("                            WHERE   Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("                                AND " & lcParametro0Hasta)
            loConsulta.AppendLine("                                AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("                                AND " & lcParametro4Hasta)
            loConsulta.AppendLine("                                AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("                                AND " & lcParametro5Hasta)
            loConsulta.AppendLine("                                AND Articulos.Cod_Mar BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("                                AND " & lcParametro6Hasta)
            loConsulta.AppendLine("                                AND Renglones_Compras.Cod_Alm BETWEEN " & lcParametro7Desde)
            loConsulta.AppendLine("                                AND " & lcParametro7Hasta)
            loConsulta.AppendLine("                            GROUP BY Renglones_Compras.Documento")
            loConsulta.AppendLine("                        ) Documentos_Compras")
            loConsulta.AppendLine("                    ON  Compras.Documento = Documentos_Compras.Documento")
            loConsulta.AppendLine("            WHERE       Compras.Fec_Ini BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("                    AND " & lcParametro1Hasta)
            loConsulta.AppendLine("                    AND Compras.Cod_Pro BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("                    AND " & lcParametro2Hasta)
            loConsulta.AppendLine("                    AND Compras.Cod_Ven BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("                    AND " & lcParametro3Hasta)
            loConsulta.AppendLine("                    AND Compras.Cod_Rev BETWEEN " & lcParametro8Desde)
            loConsulta.AppendLine("                    AND " & lcParametro8Hasta)
            loConsulta.AppendLine("                    AND Compras.Cod_Suc BETWEEN " & lcParametro9Desde)
            loConsulta.AppendLine("                    AND " & lcParametro9Hasta)
            loConsulta.AppendLine("        ) Conteo_Compras")
            loConsulta.AppendLine("CROSS JOIN (SELECT      COUNT(*) Total,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Traslados.Status = 'Anulado' THEN 1 ELSE 0 END)     AS Traslados_Anulados,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Traslados.Status = 'Pendiente' THEN 1 ELSE 0 END)   AS Traslados_Pendientes,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Traslados.Status = 'Confirmado' THEN 1 ELSE 0 END)  AS Traslados_Confirmados,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Traslados.Status = 'Procesado' THEN 1 ELSE 0 END)   AS Traslados_Procesados")
            loConsulta.AppendLine("            FROM        Traslados")
            loConsulta.AppendLine("                JOIN    (   SELECT  Renglones_Traslados.Documento ")
            loConsulta.AppendLine("                            FROM    Renglones_Traslados ")
            loConsulta.AppendLine("                            JOIN    Articulos")
            loConsulta.AppendLine("                                ON  Articulos.Cod_Art = Renglones_Traslados.Cod_Art")
            loConsulta.AppendLine("                            WHERE   Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("                                AND " & lcParametro0Hasta)
            loConsulta.AppendLine("                                AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("                                AND " & lcParametro4Hasta)
            loConsulta.AppendLine("                                AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("                                AND " & lcParametro5Hasta)
            loConsulta.AppendLine("                                AND Articulos.Cod_Mar BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("                                AND " & lcParametro6Hasta)
            loConsulta.AppendLine("                            GROUP BY Renglones_Traslados.Documento")
            loConsulta.AppendLine("                        ) Documentos_Compras")
            loConsulta.AppendLine("                    ON  Traslados.Documento = Documentos_Compras.Documento")
            loConsulta.AppendLine("            WHERE       Traslados.Fec_Ini BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("                    AND " & lcParametro1Hasta)
            loConsulta.AppendLine("                    AND Traslados.Alm_Ori BETWEEN " & lcParametro10Desde)
            loConsulta.AppendLine("                    AND " & lcParametro10Hasta)
            loConsulta.AppendLine("                    AND Traslados.Alm_Des BETWEEN " & lcParametro11Desde)
            loConsulta.AppendLine("                    AND " & lcParametro11Hasta)
            loConsulta.AppendLine("                    AND Traslados.Cod_Rev BETWEEN " & lcParametro12Desde)
            loConsulta.AppendLine("                    AND " & lcParametro12Hasta)
            loConsulta.AppendLine("                    AND Traslados.Cod_Suc BETWEEN " & lcParametro13Desde)
            loConsulta.AppendLine("                    AND " & lcParametro13Hasta)
            loConsulta.AppendLine("        ) Conteo_Traslados; ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  Tipos_de_Articulos.Articulo_Nom_Tip,")
            loConsulta.AppendLine("        Tipos_de_Articulos.Articulo_Cod_Tip,")
            loConsulta.AppendLine("        Tipos_de_Articulos.Compra_Cantidad,")
            loConsulta.AppendLine("        Tipos_de_Articulos.Compra_Bruto,")
            loConsulta.AppendLine("        Tipos_de_Articulos.Compra_Impuesto,")
            loConsulta.AppendLine("        Tipos_de_Articulos.Compra_Neto,")
            loConsulta.AppendLine("        Tipos_de_Articulos.Traslado_Cantidad,")
            loConsulta.AppendLine("        Tipos_de_Articulos.Traslado_Costo,")
            loConsulta.AppendLine("        #tmpConteo.* ")
            loConsulta.AppendLine("FROM    (   SELECT      Tipos_Articulos.Nom_Tip                             Articulo_Nom_Tip,")
            loConsulta.AppendLine("                        #tmpOperaciones.Articulo_Cod_Tip                    Articulo_Cod_Tip,")
            loConsulta.AppendLine("                        SUM(COALESCE(#tmpOperaciones.Compra_Cantidad, 0))   Compra_Cantidad, ")
            loConsulta.AppendLine("                        SUM(COALESCE(#tmpOperaciones.Compra_Bruto, 0))      Compra_Bruto,")
            loConsulta.AppendLine("                        SUM(COALESCE(#tmpOperaciones.Compra_Impuesto, 0))   Compra_Impuesto, ")
            loConsulta.AppendLine("                        SUM(COALESCE(#tmpOperaciones.Compra_Neto, 0))       Compra_Neto, ")
            loConsulta.AppendLine("                        SUM(COALESCE(#tmpOperaciones.Traslado_Cantidad, 0)) Traslado_Cantidad, ")
            loConsulta.AppendLine("                        SUM(COALESCE(#tmpOperaciones.Traslado_Costo, 0))    Traslado_Costo")
            loConsulta.AppendLine("            FROM        #tmpOperaciones")
            loConsulta.AppendLine("                JOIN    Tipos_Articulos ON Tipos_Articulos.Cod_Tip = #tmpOperaciones.Articulo_Cod_Tip")
            loConsulta.AppendLine("            GROUP BY    Tipos_Articulos.Nom_Tip,")
            loConsulta.AppendLine("                        #tmpOperaciones.Articulo_Cod_Tip")
            loConsulta.AppendLine("        ) Tipos_de_Articulos")
            loConsulta.AppendLine("    CROSS JOIN #tmpConteo")
            loConsulta.AppendLine("ORDER BY    " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rTrazabilidad_CyT_TipoArticulo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTrazabilidad_CyT_TipoArticulo.ReportSource = loObjetoReporte	

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
' RJG: 04/06/15: Codigo inicial, a partir de rTrazabilidad_Compras_Traslados.               '
'-------------------------------------------------------------------------------------------'
