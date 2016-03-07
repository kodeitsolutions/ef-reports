'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDatos_Medidores_Manuales"
'-------------------------------------------------------------------------------------------'
Partial Class fDatos_Medidores_Manuales
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

		
            Dim loComandoSeleccionar As New StringBuilder()
			

			loComandoSeleccionar.AppendLine(" 	SELECT 		")
			loComandoSeleccionar.AppendLine(" 			Medidores.Cod_Med,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Nom_Med,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Medidores.Status = 'A' THEN 'Activo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Medidores.Status = 'I' THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Medidores.Status = 'S' THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine(" 			END AS Status,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Responsable,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Fec_Ini,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Fec_Fin,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Clase,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Tipo,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Grupo,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Manual,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Por_Sup,")		
			loComandoSeleccionar.AppendLine(" 			Medidores.Por_Inf,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Por_Pro,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Mon_Sup,")		
			loComandoSeleccionar.AppendLine(" 			Medidores.Mon_Inf,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Mon_Pro,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Can_Sup,")		
			loComandoSeleccionar.AppendLine(" 			Medidores.Can_Inf,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Can_Pro,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Doc_Sup,")		
			loComandoSeleccionar.AppendLine(" 			Medidores.Doc_Inf,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Doc_Pro,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Formulario,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Reporte,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Cod_Mon,")
			loComandoSeleccionar.AppendLine(" 			Monedas.Nom_Mon,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Tasa,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Prioridad,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Importancia,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Nivel,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Comentario,")
			loComandoSeleccionar.AppendLine(" 			Medidores.Objetivo,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Renglon,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Mes,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Año,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Mon_Est,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Mon_Eje,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Mon_Des,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Can_Est,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Can_Eje,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Can_Des,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Doc_Est,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Doc_Eje,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Doc_Des,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Por_Est,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Por_Eje,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Por_Des,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Medidores.Comentario AS Comentario_Renglon")
			loComandoSeleccionar.AppendLine(" FROM	Medidores		")
			loComandoSeleccionar.AppendLine(" JOIN Renglones_Medidores ON (Renglones_Medidores.Cod_Med = Medidores.Cod_Med)")
			loComandoSeleccionar.AppendLine(" JOIN Monedas ON (Monedas.Cod_Mon = Medidores.Cod_Mon)")
			loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" ORDER BY Cod_Med, Renglon ASC ")
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
					  
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDatos_Medidores_Manuales", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfDatos_Medidores_Manuales.ReportSource = loObjetoReporte

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
' MAT: 10/08/11 : Codigo inicial															'
'-------------------------------------------------------------------------------------------'

