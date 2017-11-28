'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rPagos_Proveedores"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rPagos_Proveedores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
           
            Dim Empresa As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("DECLARE @ldFechaDesde DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaHasta DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcProDesde VARCHAR(10) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcProHasta VARCHAR(10) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT DISTINCT")
            loComandoSeleccionar.AppendLine("		Pagos.Documento							AS Documento,")
            loComandoSeleccionar.AppendLine("		Pagos.Fec_Ini							AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Pagos.Cod_Pro							AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro						AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Renglones_Pagos.Cod_Tip					AS Tipo_Doc,")
            loComandoSeleccionar.AppendLine("		Renglones_Pagos.Mon_Abo					AS Mon_Abo,")
            loComandoSeleccionar.AppendLine("		CASE WHEN Renglones_Pagos.Cod_Tip = 'ADEL'")
            loComandoSeleccionar.AppendLine("			 THEN Pagos.Ord_Com")
            loComandoSeleccionar.AppendLine("			 ELSE COALESCE(Ordenes_Compras.Documento,'')")
            loComandoSeleccionar.AppendLine("		END										AS Orden_Compra,")
            loComandoSeleccionar.AppendLine("		Renglones_Pagos.Factura					AS Factura,")
            loComandoSeleccionar.AppendLine("		Detalles_Pagos.Tip_Ope					AS Tipo_Pago,")
            loComandoSeleccionar.AppendLine("		Detalles_Pagos.Num_Doc					AS Referencia,")
            loComandoSeleccionar.AppendLine("		Detalles_Pagos.Cod_Cue					AS Cod_Cue,")
            loComandoSeleccionar.AppendLine("       Bancos.Nom_Ban							AS Nom_Ban,")
            loComandoSeleccionar.AppendLine("		Detalles_Pagos.Mon_Net					AS Mon_Net,")
            loComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFechaDesde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFechaHasta AS DATE),103))	AS Fecha,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcProDesde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro FROM Proveedores  WHERE Cod_Pro = @lcProDesde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				                        AS Pro_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcProHasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro  FROM Proveedores  WHERE Cod_Pro = @lcProHasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				                        AS Pro_Hasta")
            loComandoSeleccionar.AppendLine("FROM Pagos")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_Pagos.Cod_Tip IN ('FACT','ADEL')")
            loComandoSeleccionar.AppendLine("	JOIN Detalles_Pagos ON Detalles_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Compras ON Renglones_Pagos.Doc_Ori = Compras.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_Pagos.Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Compras ON Compras.Documento = Renglones_Compras.Documento")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Ordenes_Compras ON Renglones_Compras.Doc_Ori = Ordenes_Compras.Documento	")
            loComandoSeleccionar.AppendLine("   JOIN Proveedores ON Pagos.Cod_pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("   JOIN Bancos ON Detalles_Pagos.Cod_Ban = Bancos.Cod_Ban")
            loComandoSeleccionar.AppendLine("WHERE Pagos.Status = 'Confirmado'	")
            loComandoSeleccionar.AppendLine("	AND Pagos.Fec_Ini BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Pagos.Cod_Pro BETWEEN @lcProDesde AND @lcProHasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("ORDER BY Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rPagos_Proveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rPagos_Proveedores.ReportSource = loObjetoReporte


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
