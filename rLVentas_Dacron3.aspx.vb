'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLVentas_Dacron3"
'-------------------------------------------------------------------------------------------'
Partial Class rLVentas_Dacron3
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            
            Dim lcParametro3Desde As String = cusAplicacion.goReportes.paParametrosIniciales(3) 

            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10) 'Compras no Gravadas
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10) 'Base Importaciones
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10) 'IVA Importaciones
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10) 'Base Compras
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10) 'IVA Compras
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10) 'Base Ajuste
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10), goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10) 'IVA Ajuste
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()
		
            loConsulta.AppendLine("")
            loConsulta.AppendLine("DECLARE @lnCero AS Decimal(28,10)")
            loConsulta.AppendLine("SET	 @lnCero = 0;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--***************************************************************************************")
            loConsulta.AppendLine("-- Busca las facturas y N/CR que apareceran en el reporte.                              *")
            loConsulta.AppendLine("--***************************************************************************************")
            loConsulta.AppendLine("CREATE TABLE #tmpRegistros( Orden               INT,")
            loConsulta.AppendLine("                            Grupo               INT,")
            loConsulta.AppendLine("                            Fecha               DATE,")
            loConsulta.AppendLine("                            Tipo                CHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Documento           CHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Numero_Fiscal       CHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Impresora           CHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Cierre_Z            CHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Factura_Afectada    CHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Numero_Devolucion   CHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Cliente             CHAR(200) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Contribuyente       BIT,")
            loConsulta.AppendLine("                            Rif                 CHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Total_Venta         DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Comprobante_Fecha   DATE DEFAULT ('19000101'),")
            loConsulta.AppendLine("                            Comprobante_Factura CHAR(20) COLLATE DATABASE_DEFAULT DEFAULT '',")
            loConsulta.AppendLine("                            Comprobante_Numero  VARCHAR(30) COLLATE DATABASE_DEFAULT DEFAULT '',")
            loConsulta.AppendLine("                            Comprobante_Monto   DECIMAL(28, 10) DEFAULT 0,")
            loConsulta.AppendLine("                            Bruto_C             DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Impuesto_C          DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Bruto_NC            DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Impuesto_NC         DECIMAL(28, 10)")
            loConsulta.AppendLine("                            );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpRegistros(   Fecha, Tipo, Documento, Numero_Fiscal, Impresora, Cierre_Z, ")
            loConsulta.AppendLine("                            Numero_Devolucion, Cliente, Contribuyente, Rif,")
            loConsulta.AppendLine("                            Total_Venta, Bruto_C, Impuesto_C, Bruto_NC, Impuesto_NC)")
            loConsulta.AppendLine("SELECT      Cuentas_Cobrar.fec_ini                                      AS Fecha,")
            loConsulta.AppendLine("            Cuentas_Cobrar.Cod_Tip                                      AS Tipo,")
            loConsulta.AppendLine("            Cuentas_Cobrar.Documento                                    AS Documento,")
            loConsulta.AppendLine("            Cuentas_Cobrar.Fiscal2                                      AS Numero_Fiscal,")
            loConsulta.AppendLine("            Cuentas_Cobrar.Fiscal1                                      AS Impresora,")
            loConsulta.AppendLine("            Cuentas_Cobrar.Fiscal3                                      AS Cierre_Z,")
            loConsulta.AppendLine("            Cuentas_Cobrar.Doc_Ori                                      AS Numero_Devolucion,")
            loConsulta.AppendLine("            Clientes.Nom_Cli                                            AS Cliente,")
            loConsulta.AppendLine("            Clientes.Fiscal                                             AS Contibuyente,")
            loConsulta.AppendLine("            Clientes.Rif                                                AS Rif,")
            loConsulta.AppendLine("            Cuentas_Cobrar.Mon_net*(CASE WHEN Cuentas_Cobrar.Cod_Tip = 'FACT' ")
            loConsulta.AppendLine("                THEN 1 ELSE -1 END)                                     AS Total_Venta,")
            loConsulta.AppendLine("            (CASE WHEN Clientes.fiscal = 1")
            loConsulta.AppendLine("                    THEN Cuentas_Cobrar.Mon_net-Cuentas_Cobrar.mon_imp1")
            loConsulta.AppendLine("                    ELSE 0")
            loConsulta.AppendLine("                END)*(CASE WHEN Cuentas_Cobrar.Cod_Tip = 'FACT' ")
            loConsulta.AppendLine("                THEN 1 ELSE -1 END)                                     AS Bruto_C,")
            loConsulta.AppendLine("            (CASE WHEN clientes.fiscal = 1")
            loConsulta.AppendLine("                THEN (Cuentas_Cobrar.Mon_net-Cuentas_Cobrar.mon_imp1)*Cuentas_Cobrar.por_imp1/100")
            loConsulta.AppendLine("                ELSE 0")
            loConsulta.AppendLine("            END)*(CASE WHEN Cuentas_Cobrar.Cod_Tip = 'FACT' ")
            loConsulta.AppendLine("                THEN 1 ELSE -1 END)                                      AS Impuesto_C,")
            loConsulta.AppendLine("            (CASE WHEN clientes.fiscal = 0")
            loConsulta.AppendLine("                THEN Cuentas_Cobrar.Mon_net-Cuentas_Cobrar.mon_imp1")
            loConsulta.AppendLine("                ELSE 0")
            loConsulta.AppendLine("            END)*(CASE WHEN Cuentas_Cobrar.Cod_Tip = 'FACT' ")
            loConsulta.AppendLine("                THEN 1 ELSE -1 END)                                       AS Bruto_NC,")
            loConsulta.AppendLine("            (CASE WHEN clientes.fiscal = 0")
            loConsulta.AppendLine("                THEN (Cuentas_Cobrar.Mon_net-Cuentas_Cobrar.mon_imp1)*Cuentas_Cobrar.por_imp1/100")
            loConsulta.AppendLine("                ELSE 0")
            loConsulta.AppendLine("            END)*(CASE WHEN Cuentas_Cobrar.Cod_Tip = 'FACT' ")
            loConsulta.AppendLine("                THEN 1 ELSE -1 END)                                      AS Impuesto_NC")
            loConsulta.AppendLine("FROM        Cuentas_Cobrar")
            loConsulta.AppendLine("	JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli ")
            loConsulta.AppendLine("WHERE		Cuentas_Cobrar.Fec_Ini    BETWEEN " & lcParametro0Desde)    '@ldFechaDesde
            loConsulta.AppendLine("			AND " & lcParametro0Hasta)                                      '@ldFechaHasta
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Cod_Tip IN ('FACT','N/CR')")
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Status <> 'Anulado'")
            loConsulta.AppendLine("       	AND Cuentas_Cobrar.Cod_Suc    BETWEEN " & lcParametro1Desde)    '@lcSucursalDesde
            loConsulta.AppendLine("			AND " & lcParametro1Hasta)                                      '@lcSucursalHasta
            If lcParametro3Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro2Desde)       '@lcRevisionDesde
            Else
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro2Desde)   '@lcRevisionDesde
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro2Hasta)                                  '@lcRevisionHasta
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE NONCLUSTERED INDEX PK_tmpRegistros_Tipo_Documento ON #tmpRegistros([Tipo],[Documento]);")
            loConsulta.AppendLine("CREATE NONCLUSTERED INDEX PK_tmpRegistros_Documento_Numero_Devolucion ON #tmpRegistros([Documento]) INCLUDE ([Numero_Devolucion]);")
            loConsulta.AppendLine("CREATE NONCLUSTERED INDEX PK_tmpRegistros_Documento_Contribuyente ON #tmpRegistros([Contribuyente]) INCLUDE ([Impresora],[Cierre_Z],[Rif]);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--***************************************************************************************")
            loConsulta.AppendLine("-- Busca la información de la factura de origen de las devoluciones.                    *")
            loConsulta.AppendLine("--***************************************************************************************")
            loConsulta.AppendLine("UPDATE  #tmpRegistros")
            loConsulta.AppendLine("SET     Factura_Afectada = Origen.Numero_Fiscal")
            loConsulta.AppendLine("FROM    (   SELECT      NCR.Documento AS Devolucion,")
            loConsulta.AppendLine("                        Origenes.Doc_Ori AS Factura,")
            loConsulta.AppendLine("                        Origenes.Fiscal2 AS Numero_Fiscal")
            loConsulta.AppendLine("            FROM        #tmpRegistros AS NCR")
            loConsulta.AppendLine("	            JOIN    Renglones_dClientes AS RC")
            loConsulta.AppendLine("	                ON  RC.Documento = NCR.Numero_Devolucion")
            loConsulta.AppendLine("	                AND RC.Renglon = 1")
            loConsulta.AppendLine("	            JOIN    Cuentas_Cobrar AS Origenes")
            loConsulta.AppendLine("	                ON  Origenes.Documento = RC.Doc_Ori")
            loConsulta.AppendLine("                    AND Origenes.Cod_Tip = 'FACT'")
            loConsulta.AppendLine("        )AS Origen")
            loConsulta.AppendLine("WHERE   Tipo = 'N/CR'")
            loConsulta.AppendLine("    AND Documento = Origen.Devolucion;")
            loConsulta.AppendLine("    ")
            loConsulta.AppendLine("--***************************************************************************************")
            loConsulta.AppendLine("-- Busca la información de las retenciones.                                             *")
            loConsulta.AppendLine("--***************************************************************************************")
            loConsulta.AppendLine("INSERT INTO #tmpRegistros(   Fecha, Tipo, Documento, Numero_Fiscal, Impresora, Cierre_Z, ")
            loConsulta.AppendLine("                            Factura_Afectada, Cliente, Contribuyente, Rif,")
            loConsulta.AppendLine("                            Comprobante_Fecha, Comprobante_Factura, Comprobante_Numero, Comprobante_Monto)")
            loConsulta.AppendLine("SELECT      COALESCE(C.Fec_Ini, Cuentas_Cobrar.fec_ini)                 AS Fecha,")
            loConsulta.AppendLine("            Cuentas_Cobrar.Cod_Tip                                      AS Tipo,")
            loConsulta.AppendLine("            Cuentas_Cobrar.Documento                                    AS Documento,")
            loConsulta.AppendLine("            ''                                                          AS Numero_Fiscal,")
            loConsulta.AppendLine("            COALESCE(F.Fiscal1, '')                                     AS Impresora,")
            loConsulta.AppendLine("            COALESCE(F.Fiscal3, '')                                     AS Cierre_Z,")
            loConsulta.AppendLine("            ''                                                          AS Factura_Afectada,")
            loConsulta.AppendLine("            Clientes.Nom_Cli                                            AS Cliente,")
            loConsulta.AppendLine("            Clientes.Fiscal                                             AS Contibuyente,")
            loConsulta.AppendLine("            Clientes.Rif                                                AS Rif,")
            loConsulta.AppendLine("            Cuentas_Cobrar.Fec_Rec                                      AS Comprobante_Fecha,")
            loConsulta.AppendLine("            COALESCE(F.Fiscal2, '')                                     AS Comprobante_Factura,")
            loConsulta.AppendLine("            (CASE WHEN COALESCE(R.Num_Com, '') > ''")
            loConsulta.AppendLine("                THEN COALESCE(R.Num_Com, '')")
            loConsulta.AppendLine("                ELSE Cuentas_Cobrar.Num_Com")
            loConsulta.AppendLine("            END)                                                        AS Comprobante_Numero,")
            loConsulta.AppendLine("            Cuentas_Cobrar.Mon_net                                      AS Comprobante_Monto")
            loConsulta.AppendLine("FROM        Cuentas_Cobrar")
            loConsulta.AppendLine("	JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli ")
            loConsulta.AppendLine("	LEFT JOIN retenciones_documentos AS R")
            loConsulta.AppendLine("	    ON  R.doc_des = Cuentas_Cobrar.documento")
            loConsulta.AppendLine("	    AND  R.tip_des = 'Cuentas_Cobrar'")
            loConsulta.AppendLine("	    AND  R.cla_des = 'RETIVA'")
            loConsulta.AppendLine("	LEFT JOIN Cuentas_Cobrar AS F")
            loConsulta.AppendLine("	    ON  F.Documento = R.Doc_Ori")
            loConsulta.AppendLine("	    AND F.Cod_Tip = 'FACT'")
            loConsulta.AppendLine("	LEFT JOIN Cobros AS C")
            loConsulta.AppendLine("	    ON  C.Documento = R.Documento")
            loConsulta.AppendLine("	    AND R.ORigen = 'Cobros'")
            loConsulta.AppendLine("WHERE		Cuentas_Cobrar.Fec_Ini    BETWEEN " & lcParametro0Desde)    '@ldFechaDesde
            loConsulta.AppendLine("			AND " & lcParametro0Hasta)                                      '@ldFechaHasta
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Cod_Tip = 'RETIVA'")
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Status <> 'Anulado'")
            loConsulta.AppendLine("       	AND Cuentas_Cobrar.Cod_Suc    BETWEEN " & lcParametro1Desde)    '@lcSucursalDesde
            loConsulta.AppendLine("			AND " & lcParametro1Hasta)                                      '@lcSucursalHasta
            If lcParametro3Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro2Desde)       '@lcRevisionDesde
            Else
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro2Desde)   '@lcRevisionDesde
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro2Hasta)                                  '@lcRevisionHasta
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--'**************************************************************************************")
            loConsulta.AppendLine("--' Ajusta el cliente+Rif de los documentos a No Contribuyentes.                        *")
            loConsulta.AppendLine("--'**************************************************************************************")
            loConsulta.AppendLine("UPDATE #tmpRegistros")
            loConsulta.AppendLine("SET Cliente = ( CASE WHEN Cierre_Z > ''")
            loConsulta.AppendLine("                    THEN 'REPORTE DIARIO Z Nº ' + Cierre_Z")
            loConsulta.AppendLine("                    ELSE Cliente END),")
            loConsulta.AppendLine("    Rif     = ( CASE WHEN Impresora > '' THEN '' ELSE Rif END)")
            loConsulta.AppendLine("WHERE Contribuyente = 0 AND Tipo='FACT';")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--'**************************************************************************************")
            loConsulta.AppendLine("--' Enumera los resultados en el orden final.                                           *")
            loConsulta.AppendLine("--'**************************************************************************************")
            loConsulta.AppendLine("UPDATE  #tmpRegistros")
            loConsulta.AppendLine("SET     Orden = X.Orden")
            loConsulta.AppendLine("FROM (  SELECT  ROW_NUMBER() OVER(ORDER BY Fecha, Impresora, Tipo, Numero_Fiscal) AS Orden,")
            loConsulta.AppendLine("                Impresora, Tipo, Numero_Fiscal")
            loConsulta.AppendLine("        FROM    #tmpRegistros")
            loConsulta.AppendLine(") AS x")
            loConsulta.AppendLine("WHERE   x.Impresora = #tmpRegistros.Impresora")
            loConsulta.AppendLine("    AND x.Tipo = #tmpRegistros.Tipo")
            loConsulta.AppendLine("    AND x.Numero_Fiscal = #tmpRegistros.Numero_Fiscal;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE NONCLUSTERED INDEX PK_tmpRegistros_Orden ON #tmpRegistros([Orden]);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--'**************************************************************************************")
            loConsulta.AppendLine("--' Genera números de grupo: para agrupar las filas 'manteniendo el orden'.             *")
            loConsulta.AppendLine("--'**************************************************************************************")
            loConsulta.AppendLine("DECLARE @lnPosicion INT; ")
            loConsulta.AppendLine("DECLARE @lnGrupo INT; ")
            loConsulta.AppendLine("DECLARE @lcTipo CHAR(20); ")
            loConsulta.AppendLine("DECLARE @lnTotal INT; ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SET @lnPosicion = 0;")
            loConsulta.AppendLine("SET @lnGrupo = 0;")
            loConsulta.AppendLine("SET @lnTotal = (SELECT COUNT(*) FROM #tmpRegistros);")
            loConsulta.AppendLine("SET @lcTipo = '';")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SET NOCOUNT ON;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("WHILE (@lnPosicion <@lnTotal)")
            loConsulta.AppendLine("BEGIN")
            loConsulta.AppendLine("    SET @lnPosicion = @lnPosicion + 1;")
            loConsulta.AppendLine("    ")
            loConsulta.AppendLine("    DECLARE @lcNuevoTipo CHAR(4);")
            loConsulta.AppendLine("    DECLARE @lcTipo_Operacion CHAR(20);")
            loConsulta.AppendLine("    DECLARE @llFiscal BIT;")
            loConsulta.AppendLine("    ")
            loConsulta.AppendLine("    SELECT  @lcTipo_Operacion = Tipo,")
            loConsulta.AppendLine("            @llFiscal = Contribuyente")
            loConsulta.AppendLine("    FROM    #tmpRegistros ")
            loConsulta.AppendLine("    WHERE   Orden = @lnPosicion;")
            loConsulta.AppendLine("    ")
            loConsulta.AppendLine("    IF (@lcTipo_Operacion = 'FACT')")
            loConsulta.AppendLine("    BEGIN ")
            loConsulta.AppendLine("        IF (@llFiscal = 0)")
            loConsulta.AppendLine("            SET @lcNuevoTipo = 'F1';")
            loConsulta.AppendLine("        ELSE ")
            loConsulta.AppendLine("            SET @lcNuevoTipo = 'F2';")
            loConsulta.AppendLine("    END")
            loConsulta.AppendLine("    ELSE IF (@lcTipo_Operacion = 'N/CR')")
            loConsulta.AppendLine("    BEGIN ")
            loConsulta.AppendLine("        IF (@llFiscal = 0)")
            loConsulta.AppendLine("            SET @lcNuevoTipo = 'C1';")
            loConsulta.AppendLine("        ELSE ")
            loConsulta.AppendLine("            SET @lcNuevoTipo = 'C2';")
            loConsulta.AppendLine("    END")
            loConsulta.AppendLine("    ELSE --RETIVA")
            loConsulta.AppendLine("    BEGIN ")
            loConsulta.AppendLine("        IF (@llFiscal = 0)")
            loConsulta.AppendLine("            SET @lcNuevoTipo = 'R1';")
            loConsulta.AppendLine("        ELSE ")
            loConsulta.AppendLine("            SET @lcNuevoTipo = 'R2';")
            loConsulta.AppendLine("    END")
            loConsulta.AppendLine("    ")
            loConsulta.AppendLine("    IF ( (@lcNuevoTipo <> @lcTipo) OR (@llFiscal = 1) )")
            loConsulta.AppendLine("    BEGIN")
            loConsulta.AppendLine("        SET @lcTipo = @lcNuevoTipo;")
            loConsulta.AppendLine("        SET @lnGrupo = @lnGrupo + 1;")
            loConsulta.AppendLine("    END ")
            loConsulta.AppendLine("    ")
            loConsulta.AppendLine("    UPDATE  #tmpRegistros")
            loConsulta.AppendLine("    SET     Grupo = @lnGrupo")
            loConsulta.AppendLine("    WHERE   Orden = @lnPosicion;")
            loConsulta.AppendLine("    ")
            loConsulta.AppendLine("END;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--'**************************************************************************************")
            loConsulta.AppendLine("--' Tabla de soporte: días de la semana.                                        	    *")
            loConsulta.AppendLine("--'**************************************************************************************")
            loConsulta.AppendLine("CREATE TABLE #tmpDias(Dia DATE);")
            loConsulta.AppendLine("DECLARE @ldFechaInicio DATE;")
            loConsulta.AppendLine("DECLARE @ldFechaFin DATE;")
            loConsulta.AppendLine("SET @ldFechaInicio = " & lcParametro0Desde & ";")
            loConsulta.AppendLine("SET @ldFechaFin = " & lcParametro0Hasta & ";")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("WHILE (@ldFechaInicio <= @ldFechaFin)")
            loConsulta.AppendLine("BEGIN ")
            loConsulta.AppendLine("    INSERT INTO #tmpDias(Dia) VALUES (@ldFechaInicio); ")
            loConsulta.AppendLine("    SET @ldFechaInicio = DATEADD(DAY, 1, @ldFechaInicio);")
            loConsulta.AppendLine("END; ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SET NOCOUNT OFF;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--'**************************************************************************************")
            loConsulta.AppendLine("--' select final agrupado.                                                              *")
            loConsulta.AppendLine("--'**************************************************************************************")
            loConsulta.AppendLine("SELECT      D.Dia                                   AS Fecha,")
            loConsulta.AppendLine("            COALESCE(MIN(Numero_Fiscal), 'NO LAB.') AS Desde,")
            loConsulta.AppendLine("            MAX(Numero_Fiscal)                      AS Hasta,")
            loConsulta.AppendLine("            COALESCE(CASE Tipo ")
            loConsulta.AppendLine("                WHEN 'FACT' THEN 'Factura' ")
            loConsulta.AppendLine("                WHEN 'N/CR' THEN 'N/CR' ")
            loConsulta.AppendLine("                WHEN 'RETIVA' THEN 'Ret/IVA' ")
            loConsulta.AppendLine("                ELSE '' ")
            loConsulta.AppendLine("            END, '')                                                 AS Tipo, ")
            loConsulta.AppendLine("            MAX(Factura_Afectada)                                    AS Factura_Afectada,")
            loConsulta.AppendLine("            (CASE WHEN Impresora='' THEN 'SIF' ELSE Impresora END)   AS Impresora,      ")
            loConsulta.AppendLine("            COALESCE(Cliente, 'NO LABORABLE')                        AS Cliente,")
            loConsulta.AppendLine("            Rif                                                      AS Rif,")
            loConsulta.AppendLine("            SUM(Total_Venta)                                         AS Total_Venta,")
            loConsulta.AppendLine("            MAX(Comprobante_Fecha)                                   AS Comprobante_Fecha,")
            loConsulta.AppendLine("            MAX(Comprobante_Factura)                                 AS Comprobante_Factura,")
            loConsulta.AppendLine("            MAX(Comprobante_Numero)                                  AS Comprobante_Numero,")
            loConsulta.AppendLine("            MAX(Comprobante_Monto)                                   AS Comprobante_Monto,")
            loConsulta.AppendLine("            SUM(Bruto_C)                                             AS Bruto_C,")
            loConsulta.AppendLine("            SUM(Impuesto_C)                                          AS Impuesto_C,")
            loConsulta.AppendLine("            SUM(Bruto_NC)                                            AS Bruto_NC,")
            loConsulta.AppendLine("            SUM(Impuesto_NC)                                         AS Impuesto_NC,")
            loConsulta.AppendLine("            CAST(" & lcParametro4Desde & " AS DECIMAL(28,10))        AS Compras_NG,")
            loConsulta.AppendLine("            CAST(" & lcParametro5Desde & " AS DECIMAL(28,10))        AS C_Base_I,")
            loConsulta.AppendLine("            CAST(" & lcParametro6Desde & " AS DECIMAL(28,10))        AS IVA_Base_I,")
            loConsulta.AppendLine("            CAST(" & lcParametro7Desde & " AS DECIMAL(28,10))        AS Base_Compras,")
            loConsulta.AppendLine("            CAST(" & lcParametro8Desde & " AS DECIMAL(28,10))        AS IVA_Compras,")
            loConsulta.AppendLine("            CAST(" & lcParametro9Desde & " AS DECIMAL(28,10))        AS Base_Ajuste,")
            loConsulta.AppendLine("            CAST(" & lcParametro10Desde & " AS DECIMAL(28,10))       AS IVA_Ajuste")
            loConsulta.AppendLine("FROM        #tmpRegistros")
            loConsulta.AppendLine("    RIGHT JOIN   #tmpDias As D ON D.Dia = #tmpRegistros.Fecha")
            loConsulta.AppendLine("GROUP BY    D.Dia, COALESCE(Grupo, 0), ")
            loConsulta.AppendLine("            (CASE Tipo WHEN 'FACT' THEN 'Factura' WHEN 'N/CR' THEN 'N/CR' WHEN 'RETIVA' THEN 'Ret/IVA' ELSE '' END), ")
            loConsulta.AppendLine("            Impresora, COALESCE(Cliente, 'NO LABORABLE'), Rif")
            loConsulta.AppendLine("ORDER BY    D.Dia, Impresora, ")
            loConsulta.AppendLine("            (CASE Tipo WHEN 'FACT' THEN 'Factura' WHEN 'N/CR' THEN 'N/CR' WHEN 'RETIVA' THEN 'Ret/IVA' ELSE '' END), ")
            loConsulta.AppendLine("            MIN(Numero_Fiscal);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("DROP TABLE #tmpRegistros;")
            loConsulta.AppendLine("DROP TABLE #tmpDias;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLVentas_Dacron3", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLVentas_Resumido.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 19/12/14: Codigo inicial (A partir de rLVentas_Dracon2): este reporte calcula el IVA '
'                "al momento" en lugar de leer el IVA guardado, para recuperar los decimales'
'                perdidos al redondear cada factura (para ajustarlo a los cierres Z).		'
'-------------------------------------------------------------------------------------------'
