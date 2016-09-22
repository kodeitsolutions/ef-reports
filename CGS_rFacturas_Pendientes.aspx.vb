'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rFacturas_Pendientes"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rFacturas_Pendientes
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFechaDesde	DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaHasta	DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcProDesde	VARCHAR(10) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcProHasta	VARCHAR(10) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcFactDesde	VARCHAR(10) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcFactHasta	VARCHAR(10) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero		DECIMAL(28, 10) = 0 ;")
            loComandoSeleccionar.AppendLine("DECLARE @lcVacio		VARCHAR(10) = '';")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  'FACT'						AS Tipo,")
            loComandoSeleccionar.AppendLine("		Compras.Documento			AS Documento,")
            loComandoSeleccionar.AppendLine("		@lcVacio					AS Origen,	")
            loComandoSeleccionar.AppendLine("		Compras.Factura				AS Factura,")
            loComandoSeleccionar.AppendLine("		Compras.Control				AS Control,")
            loComandoSeleccionar.AppendLine("       Compras.Fec_Ini				AS Fecha,")
            loComandoSeleccionar.AppendLine("		Compras.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro			AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Compras.Status				AS Estatus,")
            loComandoSeleccionar.AppendLine("		Compras.Mon_Bas1			AS Base,")
            loComandoSeleccionar.AppendLine("		Compras.Mon_Imp1			AS Impuesto,")
            loComandoSeleccionar.AppendLine("		Compras.Mon_Net				AS Neto,")
            loComandoSeleccionar.AppendLine("       Compras.Mon_Sal				AS Saldo,")
            loComandoSeleccionar.AppendLine("		@lnCero						AS Abonado")
            loComandoSeleccionar.AppendLine("INTO #tmpFacturas")
            loComandoSeleccionar.AppendLine("FROM Compras")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE Compras.Mon_Sal > 0")
            loComandoSeleccionar.AppendLine("	AND Compras.Fec_Ini BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Compras.Cod_Pro BETWEEN @lcProDesde AND @lcProHasta")
            loComandoSeleccionar.AppendLine("	AND Compras.Factura BETWEEN @lcFactDesde AND @lcFactHasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Cuentas_Pagar.Cod_Tip		AS Tipo,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Documento		AS Documento,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Doc_Ori		AS Origen, ")
            loComandoSeleccionar.AppendLine("		Documentos.Factura			AS Factura,")
            loComandoSeleccionar.AppendLine("		Documentos.Control			AS Control,")
            loComandoSeleccionar.AppendLine("       Cuentas_Pagar.Fec_Ini		AS Fecha,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Pro		AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro			AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Status		AS Estatus,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Mon_Bas1		AS Base,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Mon_Imp1		AS Impuesto,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Mon_Net		AS Neto,")
            loComandoSeleccionar.AppendLine("       Cuentas_Pagar.Mon_Sal		AS Saldo,")
            loComandoSeleccionar.AppendLine("		@lnCero						AS Abonado")
            loComandoSeleccionar.AppendLine("INTO #tmpRetenciones")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("	JOIN	Cuentas_Pagar AS Documentos ON	Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loComandoSeleccionar.AppendLine("	JOIN	Proveedores ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Cod_Tip IN ('RETIVA', 'ISLR', 'RETPAT')")
            loComandoSeleccionar.AppendLine("	AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Fec_Ini BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Cod_Pro BETWEEN @lcProDesde AND @lcProHasta")
            loComandoSeleccionar.AppendLine("	AND Documentos.Documento IN (SELECT Documento FROM #tmpFacturas)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT 'Pago'						AS Tipo,")
            loComandoSeleccionar.AppendLine("		Pagos.Documento				AS Documento,")
            loComandoSeleccionar.AppendLine("		@lcVacio					AS Origen,")
            loComandoSeleccionar.AppendLine("		Renglones_Pagos.Factura		AS Factura,")
            loComandoSeleccionar.AppendLine("		Renglones_Pagos.Control		AS Control,")
            loComandoSeleccionar.AppendLine("       Pagos.Fec_Ini				AS Fecha,")
            loComandoSeleccionar.AppendLine("		Pagos.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro			AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Pagos.Status				AS Estatus,")
            loComandoSeleccionar.AppendLine("		Renglones_Pagos.Mon_Bru		AS Base,")
            loComandoSeleccionar.AppendLine("		Renglones_Pagos.Mon_Imp		AS Impuesto,")
            loComandoSeleccionar.AppendLine("		Renglones_Pagos.Mon_Net		AS Neto,")
            loComandoSeleccionar.AppendLine("       Renglones_Pagos.Mon_Sal		AS Saldo,")
            loComandoSeleccionar.AppendLine("		Renglones_Pagos.Mon_Abo		AS Abonado")
            loComandoSeleccionar.AppendLine("INTO #tmpPagos")
            loComandoSeleccionar.AppendLine("FROM Pagos")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE Pagos.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("	AND Pagos.Cod_Pro BETWEEN @lcProDesde AND @lcProHasta")
            loComandoSeleccionar.AppendLine("	AND Renglones_Pagos.Doc_Ori IN (SELECT Documento FROM #tmpFacturas WHERE Renglones_Pagos.Cod_Tip = Tipo)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT DISTINCT")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Documento		AS Doc_Factura,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Factura		AS Factura,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Control		AS Control,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Fecha			AS Fecha_Factura,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Cod_Pro		AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Nom_Pro		AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Neto			AS Neto_Factura,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Saldo			AS Saldo_Factura,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Estatus		AS Estatus_Factura,")
            loComandoSeleccionar.AppendLine("		#tmpRetenciones.Tipo        AS Tipo,")
            loComandoSeleccionar.AppendLine("		#tmpRetenciones.Origen		AS Origen,")
            loComandoSeleccionar.AppendLine("		#tmpRetenciones.Documento	AS Documento,")
            loComandoSeleccionar.AppendLine("		#tmpRetenciones.Fecha		AS Fecha,")
            loComandoSeleccionar.AppendLine("		#tmpRetenciones.Estatus		AS Estatus,")
            loComandoSeleccionar.AppendLine("		#tmpRetenciones.Neto		AS Neto,")
            loComandoSeleccionar.AppendLine("		#tmpRetenciones.Saldo		AS Saldo,")
            loComandoSeleccionar.AppendLine("		#tmpRetenciones.Abonado		AS Abonado")
            loComandoSeleccionar.AppendLine("FROM #tmpFacturas")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpRetenciones ON #tmpFacturas.Documento = #tmpRetenciones.Origen")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT DISTINCT")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Documento		AS Doc_Factura,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Factura		AS Factura,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Control		AS Control,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Fecha			AS Fecha_Factura,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Cod_Pro		AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Nom_Pro		AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Neto			AS Neto_Factura,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Saldo			AS Saldo_Factura,")
            loComandoSeleccionar.AppendLine("		#tmpFacturas.Estatus		AS Estatus_Factura,")
            loComandoSeleccionar.AppendLine("		#tmpPagos.Tipo          	AS Tipo,")
            loComandoSeleccionar.AppendLine("		#tmpPagos.Origen			AS Origen,")
            loComandoSeleccionar.AppendLine("		#tmpPagos.Documento			AS Documento,")
            loComandoSeleccionar.AppendLine("		#tmpPagos.Fecha				AS Fecha,")
            loComandoSeleccionar.AppendLine("		#tmpPagos.Estatus			AS Estatus,")
            loComandoSeleccionar.AppendLine("		#tmpPagos.Neto				AS Neto,")
            loComandoSeleccionar.AppendLine("		#tmpPagos.Saldo				AS Saldo,")
            loComandoSeleccionar.AppendLine("		#tmpPagos.Abonado			AS Abonado")
            loComandoSeleccionar.AppendLine("FROM #tmpFacturas")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpPagos ON #tmpPagos.Cod_Pro = #tmpFacturas.Cod_Pro")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Pro")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpFacturas")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpRetenciones")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpPagos")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rFacturas_Pendientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rFacturas_Pendientes.ReportSource = loObjetoReporte

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
' MJP: 16/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GS:  14/03/16: Cambio a Listado de Artículos.
'-------------------------------------------------------------------------------------------'

