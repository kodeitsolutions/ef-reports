'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fMedios"
'-------------------------------------------------------------------------------------------'
Partial Class fMedios
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

				loComandoSeleccionar.AppendLine(" SELECT ")
				loComandoSeleccionar.AppendLine(" 			cod_med, ")
				loComandoSeleccionar.AppendLine(" 			nom_med, ")
				loComandoSeleccionar.AppendLine(" 			CASE")
				loComandoSeleccionar.AppendLine(" 				WHEN Status = 'A' THEN 'Activo'")
				loComandoSeleccionar.AppendLine(" 				WHEN Status = 'I' THEN 'Inactivo'")
				loComandoSeleccionar.AppendLine(" 				WHEN Status = 'S' THEN 'Suspendido'")
				loComandoSeleccionar.AppendLine(" 			END AS Status,")
				loComandoSeleccionar.AppendLine(" 			fec_ini, ")
				loComandoSeleccionar.AppendLine(" 			fec_fin, ")
				loComandoSeleccionar.AppendLine(" 			tarifa,  ")
				loComandoSeleccionar.AppendLine(" 			tiempo,  ")
				loComandoSeleccionar.AppendLine(" 			ancho,  ")
				loComandoSeleccionar.AppendLine(" 			alto,  ")
				loComandoSeleccionar.AppendLine(" 			fondo, ")
				loComandoSeleccionar.AppendLine(" 			peso,  ")            
				loComandoSeleccionar.AppendLine(" 			volumen,  ")
				loComandoSeleccionar.AppendLine(" 			espacio, ")
				loComandoSeleccionar.AppendLine(" 			alcance, ")
				loComandoSeleccionar.AppendLine(" 			tipo,  ")
				loComandoSeleccionar.AppendLine(" 			clase, ")
				loComandoSeleccionar.AppendLine(" 			grupo, ")
				loComandoSeleccionar.AppendLine(" 			contacto,  ")
				loComandoSeleccionar.AppendLine(" 			direccion, ")
				loComandoSeleccionar.AppendLine(" 			telefonos, ")
				loComandoSeleccionar.AppendLine(" 			directo,  ")
				loComandoSeleccionar.AppendLine(" 			fax,  ")
				loComandoSeleccionar.AppendLine(" 			correo,  ")               
				loComandoSeleccionar.AppendLine(" 			correo2,  ")               
				loComandoSeleccionar.AppendLine(" 			correo3,  ")               
				loComandoSeleccionar.AppendLine(" 			web,  ")
				loComandoSeleccionar.AppendLine(" 			comentario ")
				loComandoSeleccionar.AppendLine(" FROM		Medios ")
				loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")
            
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fMedios", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfMedios.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' CMS:  27/02/10 : Codigo inicial
'-------------------------------------------------------------------------------------------'
