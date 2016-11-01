'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rLista_Cheques"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rLista_Cheques
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            'Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            'Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFechaDesde DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaHasta DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Pagos.Fec_Ini				AS Fecha,")
            loComandoSeleccionar.AppendLine("		Pagos.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro			AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Detalles_Pagos.Mon_Net		AS Monto,")
            loComandoSeleccionar.AppendLine("		Detalles_Pagos.Num_Doc		AS Cheque,")
            loComandoSeleccionar.AppendLine("		SUBSTRING(Pagos.Comentario, 16, LEN(RTRIM(Pagos.Comentario)))			AS Comentario,")
            loComandoSeleccionar.AppendLine("       CONCAT('Pago N°: ', Pagos.Documento)	AS Origen,")
            loComandoSeleccionar.AppendLine("       @ldFechaDesde               AS Desde,")
            loComandoSeleccionar.AppendLine("       @ldFechaHasta               AS Hasta")
            loComandoSeleccionar.AppendLine("FROM Pagos")
            loComandoSeleccionar.AppendLine("	JOIN Detalles_Pagos ON Pagos.Documento = Detalles_Pagos.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE Pagos.Fec_Ini	BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Detalles_Pagos.Tip_Ope = 'Cheque'")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Ordenes_Pagos.Fec_Ini		AS Fecha,")
            loComandoSeleccionar.AppendLine("		Ordenes_Pagos.Cod_Pro		AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro			AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Detalles_OPagos.Mon_Net		AS Monto,")
            loComandoSeleccionar.AppendLine("		Detalles_OPagos.Num_Doc		AS Cheque,")
            loComandoSeleccionar.AppendLine("		SUBSTRING(Ordenes_Pagos.Motivo, 16, LEN(RTRIM(Ordenes_Pagos.Motivo)))	AS Comentario,")
            loComandoSeleccionar.AppendLine("       CONCAT('Orden de Pago N°: ', Ordenes_Pagos.Documento)	AS Origen,")
            loComandoSeleccionar.AppendLine("       @ldFechaDesde               AS Desde,")
            loComandoSeleccionar.AppendLine("       @ldFechaHasta               AS Hasta")
            loComandoSeleccionar.AppendLine("FROM Ordenes_Pagos")
            loComandoSeleccionar.AppendLine("	JOIN Detalles_OPagos ON Detalles_OPagos.Documento = Ordenes_Pagos.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE Ordenes_Pagos.Fec_Ini BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Detalles_OPagos.Tip_Ope = 'Cheque'")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rLista_Cheques", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvTRF_rLista_Cheques.ReportSource = loObjetoReporte

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
' DLC: 02/09/2010: Programacion inicial (Replica del reporte rLEstadoCuenta_HistoricoVentas)'
'                   - Cambio de la consulta a procedimiento almacenado.						'
'-------------------------------------------------------------------------------------------'
' DLC: 15/09/2010: Ajuste en la forma de obtener los detalles de Pagos, asi como también,	'
'                ajustar en el RPT, la forma de mostrar los detalles de Pagos.				'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Reprogramación del Reporte y su respectivo Store Procedure					'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Ajuste de la vista de Diseño.												'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Se elimino el filtro Detalle												'
'-------------------------------------------------------------------------------------------'
' RJG: 05/12/11: Eliminado el SP: ahora la consulta se hace desde un Query en línea para	'
'				 corregir cálculo de saldo y optimizar.										'
'-------------------------------------------------------------------------------------------'
