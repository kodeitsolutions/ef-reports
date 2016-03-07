'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDatosBasicos_Idiomas"
'-------------------------------------------------------------------------------------------'
Partial Class fDatosBasicos_Idiomas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine(" SELECT")				
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Cod_Idi,")
			loComandoSeleccionar.AppendLine(" 			Idiomas.Nom_Idi AS Nom_Idi,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Renglon,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Tip_Obj,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Framework,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Renglones,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Sistema,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Modulo,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Seccion,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Objeto,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Texto,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Traduccion,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Ayuda,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Comentario")
			loComandoSeleccionar.AppendLine(" FROM Renglones_Idiomas")
			loComandoSeleccionar.AppendLine(" JOIN Idiomas ON Idiomas.Cod_Idi = Renglones_Idiomas.Cod_Idi")
			loComandoSeleccionar.AppendLine(" WHERE")	   			
            loComandoSeleccionar.AppendLine("           " & cusAplicacion.goFormatos.pcCondicionPrincipal)

			goDatos.pcNombreAplicativoExterno = "Framework"
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDatosBasicos_Idiomas", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfDatosBasicos_Idiomas.ReportSource = loObjetoReporte

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
