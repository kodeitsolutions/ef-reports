'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_fComprobante_Diario"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_fComprobante_Diario
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	Comprobantes.Documento					AS Documento,")
            loComandoSeleccionar.AppendLine("		Comprobantes.Fec_Ini					AS Fecha_Comprobante,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Renglon			AS Renglon,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Cue			AS Cod_Cuenta, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Nom_Cue			AS Nom_Cuenta, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Aux			AS Auxiliar, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Mon_Deb			AS Debe, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Mon_Hab			AS Haber, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Doc_Ori			AS Doc_Ori, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Tip_Ori			AS Tip_Ori, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Comentario		AS Comentario ")
            loComandoSeleccionar.AppendLine("FROM Comprobantes")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Comprobantes ON Comprobantes.Documento = Renglones_Comprobantes.Documento")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("TRF_fComprobante_Diario", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvTRF_fComprobante_Diario.ReportSource = loObjetoReporte

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
