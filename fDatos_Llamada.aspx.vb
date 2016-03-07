'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDatos_Llamada"
'-------------------------------------------------------------------------------------------'
Partial Class fDatos_Llamada
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

		
            Dim loComandoSeleccionar As New StringBuilder()
			

			loComandoSeleccionar.AppendLine(" 	SELECT 		")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Unico,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Documento,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Num_Ori,")		
			loComandoSeleccionar.AppendLine(" 			Llamadas.Num_Des,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Extension,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Fec_Ini,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Fec_Fin,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Status,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Duracion,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Cod_Usu,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Comentario,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Origen,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Cla_Ori,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Doc_Ori,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Destino,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Cla_Des,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Doc_Des,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Facturar,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Cod_Mon,")
			loComandoSeleccionar.AppendLine(" 			Monedas.Nom_Mon,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Tasa,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Precio,")
			loComandoSeleccionar.AppendLine(" 			Llamadas.Costo AS Monto")			
			loComandoSeleccionar.AppendLine(" FROM	Llamadas		")
			loComandoSeleccionar.AppendLine(" JOIN Monedas ON (Monedas.Cod_Mon = Llamadas.Cod_Mon)")
			loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			
			'me.mEscribirConsulta(loComandoSeleccionar.ToString)

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
					  
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDatos_Llamada", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfDatos_Llamada.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' MAT: 08/08/11 : Codigo inicial															'
'-------------------------------------------------------------------------------------------'

