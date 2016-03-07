'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRDeclaracion_pImpuestos_Auto"
'-------------------------------------------------------------------------------------------'
Partial Class rRDeclaracion_pImpuestos_Auto
    Inherits vis2formularios.frmReporteAutomatico

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load



        Dim lcFechaDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcFechaHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcClienteDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcClienteHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcProveedorDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcProveedorHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcSucursalDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcSucursalHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcRevisionDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcRevisionHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

        Dim lcRevisionIgualDif As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(5)).Trim().ToUpper()
        Dim llRevisionIgual As Boolean = (lcRevisionIgualDif = "IGUAL")

        Dim lcEstatus As String = cusAplicacion.goReportes.paParametrosIniciales(6)
        Dim lcEstatusSQL As String = goServicios.mObtenerListaFormatoSQL(lcEstatus)

        Dim lcCreditoFiscalAnterior As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        Dim lcRetencionesAcumuladas As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
        Dim llIncluirOP As Boolean = (CStr(cusAplicacion.goReportes.paParametrosIniciales(9)).Trim().ToUpper() = "SI")
        Dim llIncluirOPExentas As Boolean = (CStr(cusAplicacion.goReportes.paParametrosIniciales(10)).Trim().ToUpper() = "SI")

        Dim lcFiltroEstatus As String

        If lcEstatus.ToUpper() = "ANULADO" Then
            lcFiltroEstatus = "Status = 'Anulado'"
        ElseIf lcEstatus.ToUpper() = "PENDIENTE,CONFIRMADO,PROCESADO,PAGADO,CERRADO,AFECTADO,SERIALIZADO,CONTABILIZADO,INICIADO,CONCILIADO,OTRO,ANULADO" Then
            lcFiltroEstatus = "Status <> 'Anulado'"
        Else
            lcFiltroEstatus = "Status IN (" & lcEstatusSQL & ")"
        End If


        'Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim loComandoSeleccionar As New StringBuilder()

        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10)")
        loComandoSeleccionar.AppendLine("SET	@lnCero = 0")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("-- Exedente Credito Fiscal del mes anterior, y Retención Acumulada por Descuento del mes anterior")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("DECLARE @lnExcedenteCF DECIMAL(28, 10)")
        loComandoSeleccionar.AppendLine("DECLARE @lnAcumuladoCF DECIMAL(28, 10)")
        loComandoSeleccionar.AppendLine("DECLARE @ldFecha AS DATETIME; ")
        loComandoSeleccionar.AppendLine("SET @ldFecha = CAST(GETDATE() AS DATE); ")
        loComandoSeleccionar.AppendLine("SET @ldFecha = DATEADD(DAY,-DAY(@ldFecha) + 1, @ldFecha); ")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("SET	@lnExcedenteCF =  COALESCE((SELECT TOP 1 Val_num ")
        loComandoSeleccionar.AppendLine("                                   FROM Renglones_Series ")
        loComandoSeleccionar.AppendLine("								    WHERE cod_ser = 'EXCCEFISM'  ")
        loComandoSeleccionar.AppendLine("										AND num_ini = YEAR(@ldFecha)  ")
        loComandoSeleccionar.AppendLine("										AND num_fin = MONTH(@ldFecha)-1),0) ")
        loComandoSeleccionar.AppendLine("SET	@lnAcumuladoCF = " & lcRetencionesAcumuladas)
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("-- Base Imponible en Ventas y Débito Fiscal  ")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("SELECT		Ventas.Nacional						AS Nacional, ")
        loComandoSeleccionar.AppendLine("			Ventas.Codigo						AS Codigo, ")
        loComandoSeleccionar.AppendLine("			Ventas.Porcentaje					AS Porcentaje, ")
        loComandoSeleccionar.AppendLine("			SUM(Ventas.Base*Ventas.Signo)		AS Base, ")
        loComandoSeleccionar.AppendLine("			SUM(Ventas.Exento*Ventas.Signo)		AS Exento, ")
        loComandoSeleccionar.AppendLine("			SUM(Ventas.Impuesto*Ventas.Signo)	AS Impuesto")
        loComandoSeleccionar.AppendLine("INTO		#tmpVentas")
        loComandoSeleccionar.AppendLine("FROM	(	SELECT		DatosCC.Signo														AS Signo,")
        loComandoSeleccionar.AppendLine("						DatosCC.Nacional													AS Nacional,")
        loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./codigo)[1]', 		'varchar(MAX)')		AS Codigo,")
        loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./porcentaje)[1]', 	'decimal(28,10)')	AS Porcentaje,")
        loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./base)[1]', 		'decimal(28,10)')	AS Base,")
        loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./exento)[1]', 		'decimal(28,10)') 	AS Exento,")
        loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./monto)[1]', 		'decimal(28,10)')	AS Impuesto")
        loComandoSeleccionar.AppendLine("			FROM	(	SELECT		CAST(Cuentas_Cobrar.Dis_Imp AS XML)									AS Distribucion,")
        loComandoSeleccionar.AppendLine("									(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' THEN 1 ELSE -1 END)	AS Signo,")
        loComandoSeleccionar.AppendLine("									Clientes.Nacional													AS Nacional")
        loComandoSeleccionar.AppendLine("						FROM		Cuentas_Cobrar")
        loComandoSeleccionar.AppendLine("							JOIN	Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
        loComandoSeleccionar.AppendLine("								AND	Clientes.Cod_Cli BETWEEN " & lcClienteDesde & " AND " & lcClienteHasta)
        loComandoSeleccionar.AppendLine("						WHERE		Cuentas_Cobrar.Cod_Tip IN ('FACT', 'N/CR', 'N/DB')")
        loComandoSeleccionar.AppendLine("								AND	Cuentas_Cobrar.Fec_Ini BETWEEN " & lcFechaDesde & " AND " & lcFechaHasta)
        loComandoSeleccionar.AppendLine("								AND	Cuentas_Cobrar.Cod_Suc BETWEEN " & lcSucursalDesde & " AND " & lcSucursalHasta)
        loComandoSeleccionar.AppendLine("								AND	Cuentas_Cobrar." & lcFiltroEstatus)
        If llRevisionIgual Then
            loComandoSeleccionar.AppendLine("								AND	Cuentas_Cobrar.Cod_Rev BETWEEN " & lcRevisionDesde & " AND " & lcRevisionHasta)
        Else
            loComandoSeleccionar.AppendLine("								AND	Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcRevisionDesde & " AND " & lcRevisionHasta)
        End If
        loComandoSeleccionar.AppendLine("					) AS DatosCC")
        loComandoSeleccionar.AppendLine("				CROSS APPLY DatosCC.Distribucion.nodes('//impuestos/impuesto') AS Impuestos(Impuesto) ")
        loComandoSeleccionar.AppendLine("		) AS Ventas")
        loComandoSeleccionar.AppendLine("GROUP BY	Ventas.Nacional, Ventas.Codigo, Ventas.Porcentaje")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("-- Base Imponible en Compras y Crédito Fiscal")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("SELECT		Compras.Nacional					AS Nacional, ")
        loComandoSeleccionar.AppendLine("			Compras.Codigo						AS Codigo, ")
        loComandoSeleccionar.AppendLine("			Compras.Porcentaje					AS Porcentaje, ")
        loComandoSeleccionar.AppendLine("			SUM(Compras.Base*Compras.Signo)		AS Base, ")
        loComandoSeleccionar.AppendLine("			SUM(Compras.Exento*Compras.Signo)	AS Exento, ")
        loComandoSeleccionar.AppendLine("			SUM(Compras.Impuesto*Compras.Signo)	AS Impuesto, ")
        loComandoSeleccionar.AppendLine("			'CXP'                               AS Clase")
        loComandoSeleccionar.AppendLine("INTO		#tmpCompras")
        loComandoSeleccionar.AppendLine("FROM	(	SELECT		DatosCP.Signo														AS Signo,")
        loComandoSeleccionar.AppendLine("						DatosCP.Nacional													AS Nacional,")
        loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./codigo)[1]', 		'varchar(MAX)')		AS Codigo,")
        loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./porcentaje)[1]', 	'decimal(28,10)')	AS Porcentaje,")
        loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./base)[1]', 		'decimal(28,10)')	AS Base,")
        loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./exento)[1]', 		'decimal(28,10)') 	AS Exento,")
        loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./monto)[1]', 		'decimal(28,10)')	AS Impuesto")
        loComandoSeleccionar.AppendLine("			FROM	(	SELECT		CAST(Cuentas_Pagar.Dis_Imp AS XML)								AS Distribucion,")
        loComandoSeleccionar.AppendLine("									(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito' THEN 1 ELSE -1 END)	AS Signo,")
        loComandoSeleccionar.AppendLine("									Proveedores.Nacional											AS Nacional")
        loComandoSeleccionar.AppendLine("						FROM		Cuentas_Pagar")
        loComandoSeleccionar.AppendLine("							JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
        loComandoSeleccionar.AppendLine("								AND	Proveedores.Cod_Pro BETWEEN " & lcProveedorDesde & " AND " & lcProveedorHasta)
        loComandoSeleccionar.AppendLine("						WHERE		Cuentas_Pagar.Cod_Tip IN ('FACT', 'N/CR', 'N/DB')")
        loComandoSeleccionar.AppendLine("								AND	Cuentas_Pagar.Fec_Ini BETWEEN " & lcFechaDesde & " AND " & lcFechaHasta)
        loComandoSeleccionar.AppendLine("								AND	Cuentas_Pagar." & lcFiltroEstatus)
        loComandoSeleccionar.AppendLine("								AND	Cuentas_Pagar.Cod_Suc BETWEEN " & lcSucursalDesde & " AND " & lcSucursalHasta)
        If llRevisionIgual Then
            loComandoSeleccionar.AppendLine("								AND	Cuentas_Pagar.Cod_Rev BETWEEN " & lcRevisionDesde & " AND " & lcRevisionHasta)
        Else
            loComandoSeleccionar.AppendLine("								AND	Cuentas_Pagar.Cod_Rev NOT BETWEEN " & lcRevisionDesde & " AND " & lcRevisionHasta)
        End If
        loComandoSeleccionar.AppendLine("					) AS DatosCP")
        loComandoSeleccionar.AppendLine("				CROSS APPLY DatosCP.Distribucion.nodes('//impuestos/impuesto') AS Impuestos(Impuesto) ")
        loComandoSeleccionar.AppendLine("		) AS Compras")
        loComandoSeleccionar.AppendLine("GROUP BY	Compras.Nacional, Compras.Codigo, Compras.Porcentaje")

        If llIncluirOP Then

            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
            loComandoSeleccionar.AppendLine("-- Base Imponible en Órdenes de Pago y Crédito Fiscal de OP")
            loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
            loComandoSeleccionar.AppendLine("SELECT		OPagos.Nacional			AS Nacional, ")
            loComandoSeleccionar.AppendLine("			OPagos.Codigo			AS Codigo, ")
            loComandoSeleccionar.AppendLine("			OPagos.Porcentaje		AS Porcentaje, ")
            loComandoSeleccionar.AppendLine("			SUM(OPagos.Base)		AS Base, ")
            loComandoSeleccionar.AppendLine("			SUM(OPagos.Exento)		AS Exento, ")
            loComandoSeleccionar.AppendLine("			SUM(OPagos.Impuesto)	AS Impuesto, ")
            loComandoSeleccionar.AppendLine("			'OP'                    AS Clase")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT		DatosOP.Nacional													AS Nacional,")
            loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./codigo)[1]', 		'varchar(MAX)')		AS Codigo,")
            loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./porcentaje)[1]', 	'decimal(28,10)')	AS Porcentaje,")
            loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./base)[1]', 		'decimal(28,10)')	AS Base,")
            loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./exento)[1]', 		'decimal(28,10)') 	AS Exento,")
            loComandoSeleccionar.AppendLine("						Impuestos.Impuesto.value('(./monto)[1]', 		'decimal(28,10)')	AS Impuesto")
            loComandoSeleccionar.AppendLine("			FROM	(	SELECT		CAST(Ordenes_Pagos.Dis_Imp AS XML)	AS Distribucion,")
            loComandoSeleccionar.AppendLine("									Proveedores.Nacional				AS Nacional")
            loComandoSeleccionar.AppendLine("						FROM		Ordenes_Pagos")
            loComandoSeleccionar.AppendLine("							JOIN	Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("								AND	Proveedores.Cod_Pro BETWEEN " & lcProveedorDesde & " AND " & lcProveedorHasta)
            loComandoSeleccionar.AppendLine("						WHERE		Ordenes_Pagos.Fec_Ini BETWEEN " & lcFechaDesde & " AND " & lcFechaHasta)
            loComandoSeleccionar.AppendLine("								AND	Ordenes_Pagos." & lcFiltroEstatus)
            loComandoSeleccionar.AppendLine("								AND	Ordenes_Pagos.Cod_Suc BETWEEN " & lcSucursalDesde & " AND " & lcSucursalHasta)
            If llRevisionIgual Then
                loComandoSeleccionar.AppendLine("								AND	Ordenes_Pagos.Cod_Rev BETWEEN " & lcRevisionDesde & " AND " & lcRevisionHasta)
            Else
                loComandoSeleccionar.AppendLine("								AND	Ordenes_Pagos.Cod_Rev NOT BETWEEN " & lcRevisionDesde & " AND " & lcRevisionHasta)
            End If
            If Not llIncluirOPExentas Then
                loComandoSeleccionar.AppendLine("								AND	Ordenes_Pagos.Mon_Imp > 0")
            End If
            loComandoSeleccionar.AppendLine("					) AS DatosOP")
            loComandoSeleccionar.AppendLine("				CROSS APPLY DatosOP.Distribucion.nodes('//impuestos/impuesto') AS Impuestos(Impuesto) ")
            loComandoSeleccionar.AppendLine("		) AS OPagos")
            loComandoSeleccionar.AppendLine("GROUP BY	OPagos.Nacional, OPagos.Codigo, OPagos.Porcentaje")
            loComandoSeleccionar.AppendLine("")

        End If


        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("-- Impuesto excluido en compras (GACETA 5162)")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("SELECT  Documento, Cod_Tip, Signo ")
        loComandoSeleccionar.AppendLine("INTO    #tmpCxPIncluidas")
        loComandoSeleccionar.AppendLine("FROM (  SELECT		Cuentas_Pagar.Documento,")
        loComandoSeleccionar.AppendLine("                    Cuentas_Pagar.Cod_Tip,")
        loComandoSeleccionar.AppendLine("                    (CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito' THEN 1 ELSE -1 END)	AS Signo")
        loComandoSeleccionar.AppendLine("		 FROM		Cuentas_Pagar")
        loComandoSeleccionar.AppendLine("		 	 JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
        loComandoSeleccionar.AppendLine("		 		AND	Proveedores.Cod_Pro BETWEEN " & lcProveedorDesde & " AND " & lcProveedorHasta)
        loComandoSeleccionar.AppendLine("		 WHERE		Cuentas_Pagar.Cod_Tip IN ('FACT', 'N/CR', 'N/DB')")
        loComandoSeleccionar.AppendLine("		 		AND	Cuentas_Pagar.Fec_Ini BETWEEN " & lcFechaDesde & " AND " & lcFechaHasta)
        loComandoSeleccionar.AppendLine("		 		AND	Cuentas_Pagar." & lcFiltroEstatus)
        loComandoSeleccionar.AppendLine("		 		AND	Cuentas_Pagar.Cod_Suc BETWEEN " & lcSucursalDesde & " AND " & lcSucursalHasta)
        If llRevisionIgual Then
            loComandoSeleccionar.AppendLine("	 			AND	Cuentas_Pagar.Cod_Rev BETWEEN " & lcRevisionDesde & " AND " & lcRevisionHasta)
        Else
            loComandoSeleccionar.AppendLine("	 			AND	Cuentas_Pagar.Cod_Rev NOT BETWEEN " & lcRevisionDesde & " AND " & lcRevisionHasta)
        End If
        loComandoSeleccionar.AppendLine(") Documentos;")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("DECLARE @lnImpuestoExcluidoCompras DECIMAL(28,10);")
        loComandoSeleccionar.AppendLine("SET @lnImpuestoExcluidoCompras = (")
        loComandoSeleccionar.AppendLine("    SELECT  SUM(Excluir.Mon_Imp1) Mon_Imp1")
        loComandoSeleccionar.AppendLine("    FROM  ( SELECT      (Cuentas_Pagar.Mon_Imp1*#tmpCxPIncluidas.Signo) Mon_Imp1")
        loComandoSeleccionar.AppendLine("            FROM        #tmpCxPIncluidas")
        loComandoSeleccionar.AppendLine("                JOIN    Cuentas_Pagar ")
        loComandoSeleccionar.AppendLine("                    ON  Cuentas_Pagar.Documento = #tmpCxPIncluidas.Documento")
        loComandoSeleccionar.AppendLine("                    AND Cuentas_Pagar.Cod_Tip = #tmpCxPIncluidas.Cod_tip")
        loComandoSeleccionar.AppendLine("                    AND Cuentas_Pagar.Automatico = 0")
        loComandoSeleccionar.AppendLine("                    AND CAST(Not_Sta AS XML).value('(clasificacion/logico1)[1]', 'VARCHAR(MAX)') = 'true'")
        loComandoSeleccionar.AppendLine("            UNION ALL  ")
        loComandoSeleccionar.AppendLine("            SELECT      (Compras.Mon_Imp1*#tmpCxPIncluidas.Signo) Mon_Imp1")
        loComandoSeleccionar.AppendLine("            FROM        #tmpCxPIncluidas")
        loComandoSeleccionar.AppendLine("                JOIN    Compras ")
        loComandoSeleccionar.AppendLine("                    ON  Compras.Documento = #tmpCxPIncluidas.Documento")
        loComandoSeleccionar.AppendLine("            WHERE       CAST(Not_Sta AS XML).value('(clasificacion/logico1)[1]', 'VARCHAR(MAX)') = 'true'")
        loComandoSeleccionar.AppendLine("            UNION ALL  ")
        loComandoSeleccionar.AppendLine("            SELECT      (Renglones_Compras.Mon_Imp1*#tmpCxPIncluidas.Signo)  Mon_Imp1")
        loComandoSeleccionar.AppendLine("            FROM        #tmpCxPIncluidas")
        loComandoSeleccionar.AppendLine("                JOIN    Compras ")
        loComandoSeleccionar.AppendLine("                    ON  Compras.Documento = #tmpCxPIncluidas.Documento")
        loComandoSeleccionar.AppendLine("                    AND CAST(Compras.Not_Sta AS XML).value('(clasificacion/logico1)[1]', 'VARCHAR(MAX)') <> 'true'")
        loComandoSeleccionar.AppendLine("                JOIN    Renglones_Compras ")
        loComandoSeleccionar.AppendLine("                    ON  Renglones_Compras.Documento = Compras.Documento")
        loComandoSeleccionar.AppendLine("                    AND CAST(Renglones_Compras.Not_Sta AS XML).value('(clasificacion/logico1)[1]', 'VARCHAR(MAX)') = 'true'")
        loComandoSeleccionar.AppendLine("        ) Excluir ")
        loComandoSeleccionar.AppendLine("    ); ")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")

        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("-- Impuesto Retenido a Proveedor")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("DECLARE	 @lnRetProveedor DECIMAL(28,10)")
        loComandoSeleccionar.AppendLine("SET @lnRetProveedor = (")
        loComandoSeleccionar.AppendLine("	SELECT		ISNULL(SUM(Cuentas_Pagar.Mon_Net), @lnCero)")
        loComandoSeleccionar.AppendLine("	FROM		Cuentas_Pagar")
        'loComandoSeleccionar.AppendLine("		LEFT JOIN Revisiones ON Revisiones.Cod_Rev = Cuentas_Pagar.Cod_Rev")
        loComandoSeleccionar.AppendLine("	WHERE		Cuentas_Pagar.Cod_Tip = 'RETIVA'")
        loComandoSeleccionar.AppendLine("		AND		Cuentas_Pagar." & lcFiltroEstatus)
        loComandoSeleccionar.AppendLine("		AND		Cuentas_Pagar.Fec_Ini BETWEEN " & lcFechaDesde & " AND " & lcFechaHasta)
        loComandoSeleccionar.AppendLine("		AND		Cuentas_Pagar.Cod_Pro BETWEEN " & lcProveedorDesde & " AND " & lcProveedorHasta)
        loComandoSeleccionar.AppendLine("		AND		Cuentas_Pagar.Cod_Suc BETWEEN " & lcSucursalDesde & " AND " & lcSucursalHasta)
        If llRevisionIgual Then
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Cod_Rev BETWEEN " & lcRevisionDesde & " AND " & lcRevisionHasta)
        Else
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Cod_Rev NOT BETWEEN " & lcRevisionDesde & " AND " & lcRevisionHasta)
        End If
        'loComandoSeleccionar.AppendLine("		AND		ISNULL(Revisiones.Tipo,'') BETWEEN " & lcTipoRevisionDesde & " AND " & lcTipoRevisionHasta)
        loComandoSeleccionar.AppendLine(")")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("-- Impuesto Retenido por Cliente")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("DECLARE	 @lnRetCliente DECIMAL(28,10)")
        loComandoSeleccionar.AppendLine("SET	@lnRetCliente = (")
        loComandoSeleccionar.AppendLine("	SELECT		ISNULL(SUM(Cuentas_Cobrar.Mon_Net), @lnCero)")
        loComandoSeleccionar.AppendLine("	FROM		Cuentas_Cobrar")
        'loComandoSeleccionar.AppendLine("		LEFT JOIN Revisiones ON Revisiones.Cod_Rev = Cuentas_Cobrar.Cod_Rev")
        loComandoSeleccionar.AppendLine("	WHERE		Cuentas_Cobrar.Cod_Tip = 'RETIVA'")
        loComandoSeleccionar.AppendLine("		AND		Cuentas_Cobrar." & lcFiltroEstatus)
        loComandoSeleccionar.AppendLine("		AND		Cuentas_Cobrar.Fec_Ini BETWEEN " & lcFechaDesde & " AND " & lcFechaHasta)
        loComandoSeleccionar.AppendLine("		AND		Cuentas_Cobrar.Cod_Cli BETWEEN " & lcClienteDesde & " AND " & lcClienteHasta)
        loComandoSeleccionar.AppendLine("		AND		Cuentas_Cobrar.Cod_Suc BETWEEN " & lcSucursalDesde & " AND " & lcSucursalHasta)
        If llRevisionIgual Then
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Cobrar.Cod_Rev BETWEEN " & lcRevisionDesde & " AND " & lcRevisionHasta)
        Else
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcRevisionDesde & " AND " & lcRevisionHasta)
        End If
        'loComandoSeleccionar.AppendLine("		AND		ISNULL(Revisiones.Tipo,'') BETWEEN " & lcTipoRevisionDesde & " AND " & lcTipoRevisionHasta)
        loComandoSeleccionar.AppendLine(")")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("-- SELECT Final")
        loComandoSeleccionar.AppendLine("-- ---------------------------------------------------------------------------------------------")
        loComandoSeleccionar.AppendLine("SELECT		1						    AS Tipo,")
        loComandoSeleccionar.AppendLine("			Nacional				    AS Nacional,")
        loComandoSeleccionar.AppendLine("			Codigo					    As V_Codigo,")
        loComandoSeleccionar.AppendLine("			Porcentaje				    AS V_Porcentaje,")
        loComandoSeleccionar.AppendLine("			(Base+Exento)			    AS V_Base,")
        loComandoSeleccionar.AppendLine("			Impuesto				    AS V_Impuesto,")
        loComandoSeleccionar.AppendLine("			''						    As C_Codigo,")
        loComandoSeleccionar.AppendLine("			@lnCero					    AS C_Porcentaje,")
        loComandoSeleccionar.AppendLine("			@lnCero					    AS C_Base,")
        loComandoSeleccionar.AppendLine("			@lnCero					    AS C_Impuesto,")
        loComandoSeleccionar.AppendLine("			@lnExcedenteCF			    AS Excedente_CF,")
        loComandoSeleccionar.AppendLine("			@lnRetProveedor			    AS RetProveedor,")
        loComandoSeleccionar.AppendLine("			@lnRetCliente			    AS RetCliente,")
        loComandoSeleccionar.AppendLine("			@lnAcumuladoCF			    AS Acumulado_CF,")
        loComandoSeleccionar.AppendLine("			@lnImpuestoExcluidoCompras  AS ImpuestoExcluido_C")
        loComandoSeleccionar.AppendLine("FROM		#tmpVentas ")
        loComandoSeleccionar.AppendLine("UNION ALL")
        loComandoSeleccionar.AppendLine("SELECT		2						    AS Tipo,")
        loComandoSeleccionar.AppendLine("			Nacional				    AS Nacional,")
        loComandoSeleccionar.AppendLine("			''						    As V_Codigo,")
        loComandoSeleccionar.AppendLine("			@lnCero					    AS V_Porcentaje,")
        loComandoSeleccionar.AppendLine("			@lnCero					    AS V_Base,")
        loComandoSeleccionar.AppendLine("			@lnCero					    AS V_Impuesto,")
        loComandoSeleccionar.AppendLine("			Codigo					    As C_Codigo,")
        loComandoSeleccionar.AppendLine("			Porcentaje				    AS C_Porcentaje,")
        loComandoSeleccionar.AppendLine("			SUM(Base+Exento)		    AS C_Base,")
        loComandoSeleccionar.AppendLine("			SUM(Impuesto)			    AS C_Impuesto,")
        loComandoSeleccionar.AppendLine("			@lnExcedenteCF			    AS Excedente_CF,")
        loComandoSeleccionar.AppendLine("			@lnRetProveedor			    AS RetProveedor,")
        loComandoSeleccionar.AppendLine("			@lnRetCliente			    AS RetCliente,")
        loComandoSeleccionar.AppendLine("			@lnAcumuladoCF			    AS Acumulado_CF,")
        loComandoSeleccionar.AppendLine("			@lnImpuestoExcluidoCompras  AS ImpuestoExcluido_C")
        loComandoSeleccionar.AppendLine("FROM		#tmpCompras ")
        loComandoSeleccionar.AppendLine("GROUP BY	#tmpCompras.Nacional, #tmpCompras.Codigo, #tmpCompras.Porcentaje ")
        loComandoSeleccionar.AppendLine("UNION ALL")
        loComandoSeleccionar.AppendLine("SELECT		3						    AS Tipo,")
        loComandoSeleccionar.AppendLine("			CAST(0 AS BIT)			    AS Nacional,")
        loComandoSeleccionar.AppendLine("			''						    As V_Codigo,")
        loComandoSeleccionar.AppendLine("			@lnCero					    AS V_Porcentaje,")
        loComandoSeleccionar.AppendLine("			@lnCero					    AS V_Base,")
        loComandoSeleccionar.AppendLine("			@lnCero					    AS V_Impuesto,")
        loComandoSeleccionar.AppendLine("			''						    As C_Codigo,")
        loComandoSeleccionar.AppendLine("			@lnCero					    AS C_Porcentaje,")
        loComandoSeleccionar.AppendLine("			@lnCero					    AS C_Base,")
        loComandoSeleccionar.AppendLine("			@lnCero					    AS C_Impuesto,")
        loComandoSeleccionar.AppendLine("			@lnExcedenteCF			    AS Excedente_CF,")
        loComandoSeleccionar.AppendLine("			@lnRetProveedor			    AS RetProveedor,")
        loComandoSeleccionar.AppendLine("			@lnRetCliente			    AS RetCliente,")
        loComandoSeleccionar.AppendLine("			@lnAcumuladoCF			    AS Acumulado_CF,")
        loComandoSeleccionar.AppendLine("			@lnImpuestoExcluidoCompras  AS ImpuestoExcluido_C")
        loComandoSeleccionar.AppendLine("ORDER BY	Tipo ASC, Nacional ASC, V_Porcentaje ASC, C_Porcentaje ASC, Codigo ASC")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("DROP TABLE #tmpVentas")
        loComandoSeleccionar.AppendLine("DROP TABLE #tmpCompras ")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("")

        'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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


        loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRDeclaracion_pImpuestos_Auto", laDatosReporte)

        Me.mTraducirReporte(loObjetoReporte)

        Me.mFormatearCamposReporte(loObjetoReporte)

        Me.crvrRDeclaracion_pImpuestos_Auto.ReportSource = loObjetoReporte


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
' RJG: 25/06/12: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
' RJG: 17/07/12: Se corrigió el parámetro de Tipo de revisión. Se corrigió las etiquetas de '
'				 totales en el RPT (Estaban mal escritas).									'
'-------------------------------------------------------------------------------------------'
' RJG: 20/04/15: Se agregó la exclusión de los impuestos de compras/renglones que están     '
'                marcados como excluidos en la clasificación del documento (GACETA 6152).   '
'-------------------------------------------------------------------------------------------'
