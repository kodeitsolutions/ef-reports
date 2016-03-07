'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDatosBasicos_Cajeros"
'-------------------------------------------------------------------------------------------'
Partial Class fDatosBasicos_Cajeros
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine(" SELECT")				
			loComandoSeleccionar.AppendLine(" 			Cajeros.Cod_Caj,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Nom_Caj,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Cajeros.Status = 'A' THEN 'Activo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Cajeros.Status = 'I' THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Cajeros.Status = 'S' THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine(" 			END AS Status,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Fec_Ini,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Fec_Fin,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Caja,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Mon_Fon,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Clave,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Correo,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Telefonos,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Movil,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Nivel,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Prioridad,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Supervisor,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Tipo,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Clase,")
			loComandoSeleccionar.AppendLine(" 			Cajeros.Comentario,")
			loComandoSeleccionar.AppendLine(" 			Cajas.Nom_Caj AS Nom_Caja")
			loComandoSeleccionar.AppendLine(" FROM Cajeros")
			loComandoSeleccionar.AppendLine(" JOIN Cajas ON Cajas.Cod_Caj = Cajeros.Caja")
			loComandoSeleccionar.AppendLine(" WHERE")	   			
            loComandoSeleccionar.AppendLine("           " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDatosBasicos_Cajeros", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfDatosBasicos_Cajeros.ReportSource = loObjetoReporte

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
' MAT: 11/01/11 : Codigo inicial
'-------------------------------------------------------------------------------------------'
