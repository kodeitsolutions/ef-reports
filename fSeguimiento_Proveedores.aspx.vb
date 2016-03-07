'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fSeguimiento_Proveedores"
'-------------------------------------------------------------------------------------------'
Partial Class fSeguimiento_Proveedores
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" 	SELECT 		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Cod_Ven,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Fec_Ini,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Hor_Ini,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Lugar,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Contacto,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Accion,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Notas,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Comentario,		")
			loComandoSeleccionar.AppendLine(" 				Proveedores.Cod_Pro,		")
			loComandoSeleccionar.AppendLine(" 				Proveedores.Nom_Pro,		")
			loComandoSeleccionar.AppendLine(" 				Vendedores.Nom_Ven		")
			loComandoSeleccionar.AppendLine(" 	FROM Seguimientos		")
			loComandoSeleccionar.AppendLine(" 	JOIN Proveedores ON Proveedores.Cod_Pro = Seguimientos.cod_reg	 ")
			loComandoSeleccionar.AppendLine(" 	JOIN Vendedores ON Vendedores.Cod_Ven = Seguimientos.cod_Ven ")
			loComandoSeleccionar.AppendLine(" WHERE")	   			
            loComandoSeleccionar.AppendLine("          Seguimientos.tipo = 'Proveedores' AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            
            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
           	'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
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

            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fSeguimiento_Proveedores", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfSeguimiento_Proveedores.ReportSource = loObjetoReporte

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
' CMS:  29/05/10 : Codigo inicial
'-------------------------------------------------------------------------------------------'

