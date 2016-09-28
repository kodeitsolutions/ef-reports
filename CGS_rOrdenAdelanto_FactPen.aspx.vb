'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rOrdenAdelanto_FactPen"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rOrdenAdelanto_FactPen
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFechaDesde	DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaHasta	DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcProDesde	VARCHAR(10) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcProHasta	VARCHAR(10) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero			DECIMAL(28, 10) = 0 ;")
            loComandoSeleccionar.AppendLine("DECLARE @lcVacio		VARCHAR(10) = '';")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Órdenes de Compras'		AS Tipo,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Documento	AS Documento,")
            loComandoSeleccionar.AppendLine("		@lcVacio					AS Factura,")
            loComandoSeleccionar.AppendLine("		@lcVacio					AS Control,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Fec_Ini		AS Fecha,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Cod_Pro		AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro			AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Status		AS Estatus,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Net		AS Monto,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Sal		AS Saldo,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Comentario	AS Comentario,")
            loComandoSeleccionar.AppendLine("		@lnCero						AS Abonado_Pago,")
            loComandoSeleccionar.AppendLine("		SUM(Cuentas_Pagar.Mon_Net)	AS Monto_Adel,")
            loComandoSeleccionar.AppendLine("		@lnCero						AS Neto_Ret,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Net - SUM(Cuentas_Pagar.Mon_Net) AS Deuda")
            loComandoSeleccionar.AppendLine("FROM Ordenes_Compras")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Ordenes_Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine("	JOIN Pagos ON Pagos.Ord_Com = Ordenes_Compras.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Pagos ON Pagos.Documento = Renglones_Pagos.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Pagar ON Cuentas_Pagar.Documento = Renglones_Pagos.Doc_Ori")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Tip = 'ADEL'")
            loComandoSeleccionar.AppendLine("WHERE Ordenes_Compras.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("	--AND Ordenes_Compras.Fec_Ini > '01/06/2016'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Status <> 'Pagado'")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Fec_Ini BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Cod_Pro BETWEEN @lcProDesde AND @lcProHasta")
            loComandoSeleccionar.AppendLine("GROUP BY Ordenes_Compras.Documento, Ordenes_Compras.Fec_Ini, Ordenes_Compras.Cod_Pro, Proveedores.Nom_Pro, Ordenes_Compras.Status,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Net, Ordenes_Compras.Mon_Sal, Ordenes_Compras.Comentario, Pagos.Documento,")
            loComandoSeleccionar.AppendLine("		Pagos.Fec_Ini, Cuentas_Pagar.Documento, Cuentas_Pagar.Fec_Ini, Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Facturas'										AS Tipo,")
            loComandoSeleccionar.AppendLine("		Compras.Documento								AS Documento,")
            loComandoSeleccionar.AppendLine("		Compras.Factura									AS Factura,")
            loComandoSeleccionar.AppendLine("		Compras.Control									AS Control,")
            loComandoSeleccionar.AppendLine("		Compras.Fec_Ini									AS Fecha,")
            loComandoSeleccionar.AppendLine("		Compras.Cod_Pro									AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro								AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Compras.Status									AS Estatus,")
            loComandoSeleccionar.AppendLine("		Compras.Mon_Net									AS Monto,")
            loComandoSeleccionar.AppendLine("		Compras.Mon_Sal									AS Saldo,")
            loComandoSeleccionar.AppendLine("		Compras.Comentario								AS Comentario,")
            loComandoSeleccionar.AppendLine("		COALESCE(SUM(Renglones_Pagos.Mon_Abo), @lnCero)	AS Abonado_Pago,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Monto_Adel,")
            loComandoSeleccionar.AppendLine("		COALESCE(SUM(Cuentas_Pagar.Mon_Net), @lnCero)	AS Neto_Ret,")
            loComandoSeleccionar.AppendLine("		Compras.Mon_Net ")
            loComandoSeleccionar.AppendLine("		- COALESCE(SUM(Cuentas_Pagar.Mon_Net), @lnCero)")
            loComandoSeleccionar.AppendLine("		- COALESCE(SUM(Renglones_Pagos.Mon_Abo), @lnCero)	AS Deuda")
            loComandoSeleccionar.AppendLine("FROM Compras")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Pagar ON Cuentas_Pagar.Doc_Ori = Compras.Documento")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Tip IN ('RETIVA', 'ISLR', 'RETPAT')")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Pro = Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Status NOT IN ('Anulado', 'Pagado')")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Pagos ")
            loComandoSeleccionar.AppendLine("		INNER JOIN Pagos ")
            loComandoSeleccionar.AppendLine("			ON (Renglones_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine("			AND Pagos.Status = 'Confirmado')	")
            loComandoSeleccionar.AppendLine("	ON Renglones_Pagos.Doc_Ori = Compras.Documento ")
            loComandoSeleccionar.AppendLine("		AND Renglones_Pagos.Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("WHERE Compras.Mon_Sal > 0	")
            loComandoSeleccionar.AppendLine("	AND Compras.Fec_Ini BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Compras.Cod_Pro BETWEEN @lcProDesde AND @lcProHasta")
            loComandoSeleccionar.AppendLine("GROUP BY Compras.Documento, Compras.Factura, Compras.Control, Compras.Fec_Ini, Compras.Cod_Pro, Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Compras.Status, Compras.Mon_Net	, Compras.Mon_Sal, Compras.Comentario")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Pro")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rOrdenAdelanto_FactPen", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rOrdenAdelanto_FactPen.ReportSource = loObjetoReporte

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

