'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fContabilizacion_Informacion"
'-------------------------------------------------------------------------------------------'
Partial Class fContabilizacion_Informacion
     Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine(" SELECT")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Documento,")
			loComandoSeleccionar.AppendLine(" 			CONVERT(nchar(15), Integraciones.Fec_Ini,103) AS Fec_Ini,")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Status,")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Comentario,")
			loComandoSeleccionar.AppendLine(" 			CONVERT(nchar(15), Integraciones.Int_Ini,103) AS Int_Ini,")
			loComandoSeleccionar.AppendLine(" 			CONVERT(nchar(15), Integraciones.Int_Fin,103) AS Int_Fin,")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Doc_Des,")
			loComandoSeleccionar.AppendLine(" 			CONVERT(nchar(15), Integraciones.Fec_Com,103) AS Fec_Com,")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Agrupacion,")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Tip_Cos,")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Mar_Doc,")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Doc_Cua,")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Agr_Fec,")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Agr_Tip,")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Doc_Fal,")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Logico1,")
			loComandoSeleccionar.AppendLine(" 			Integraciones.Act_Fec")
			loComandoSeleccionar.AppendLine(" INTO #tmpTemporal1")
            loComandoSeleccionar.AppendLine(" FROM Integraciones")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ")
            
            
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 		Renglones_Integraciones.Documento,")
			loComandoSeleccionar.AppendLine(" 		Renglones_Integraciones.Renglon,")
			loComandoSeleccionar.AppendLine(" 		Renglones_Integraciones.Cod_Tip,")
			loComandoSeleccionar.AppendLine(" 		Renglones_Integraciones.Nom_Tip,")
			loComandoSeleccionar.AppendLine(" 		Renglones_Integraciones.Comentario AS Comentario_Renglon")
			loComandoSeleccionar.AppendLine(" INTO #tmpTemporal2")
			loComandoSeleccionar.AppendLine(" FROM Renglones_Integraciones")
			loComandoSeleccionar.AppendLine(" JOIN Integraciones ON Renglones_Integraciones.Documento	=	Integraciones.Documento")
            loComandoSeleccionar.AppendLine(" WHERE    " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" ")
			
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 		#tmpTemporal1.*,")
			loComandoSeleccionar.AppendLine(" 		#tmpTemporal2.Renglon,")
			loComandoSeleccionar.AppendLine(" 		#tmpTemporal2.Cod_Tip,")
			loComandoSeleccionar.AppendLine(" 		#tmpTemporal2.Nom_Tip,")
			loComandoSeleccionar.AppendLine(" 		#tmpTemporal2.Comentario_Renglon")
			loComandoSeleccionar.AppendLine(" FROM #tmpTemporal1")
			loComandoSeleccionar.AppendLine(" JOIN #tmpTemporal2 ON #tmpTemporal2.Documento	= #tmpTemporal1.Documento")
			loComandoSeleccionar.AppendLine(" ORDER BY Documento,Renglon ASC")
   
   
   
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
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

             
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fContabilizacion_Informacion", laDatosReporte)


            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfContabilizacion_Informacion.ReportSource = loObjetoReporte

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
' CMS:  16/09/09 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT:  04/03/11 : Ajuste del SELECT, el formato no mostraba información.
'-------------------------------------------------------------------------------------------'
' MAT:  04/03/11 : Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT:  04/03/11 : Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
