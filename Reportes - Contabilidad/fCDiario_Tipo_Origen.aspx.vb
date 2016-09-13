Imports System.Data
Partial Class fCDiario_Tipo_Origen

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Comprobantes.Documento							AS Documento, ")
            loComandoSeleccionar.AppendLine("           YEAR(Comprobantes.Fec_Ini)						AS Anno, ")
            loComandoSeleccionar.AppendLine("           MONTH(Comprobantes.Fec_Ini)						AS Mes, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Fec_Ini 							AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Fec_Fin 							AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Resumen							AS Resumen, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Tipo								As Tipo, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Origen								AS Origen, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Integracion						AS Integracion, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Status								As Status")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTemporal1 ")
            loComandoSeleccionar.AppendLine(" FROM      Comprobantes ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

           
            loComandoSeleccionar.AppendLine("SELECT		Comprobantes.Documento								AS Doc_Tipo, ")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Comprobantes.Mon_Deb)					As Mon_Deb, ")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Comprobantes.Mon_Hab)					AS Mon_hab, ")
            loComandoSeleccionar.AppendLine("			(CASE	")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Comprobantes.Tip_Ori <> '')")
            loComandoSeleccionar.AppendLine("				THEN Renglones_Comprobantes.Tip_Ori")
            loComandoSeleccionar.AppendLine("				ELSE '(Sin Tipo)'")
            loComandoSeleccionar.AppendLine("			END)												AS Tip_Ori ")
            loComandoSeleccionar.AppendLine("INTO		#tmpTemporal2")
            loComandoSeleccionar.AppendLine("FROM		Comprobantes ")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Comprobantes ON (Comprobantes.Documento  =  Renglones_Comprobantes.Documento)")
            loComandoSeleccionar.AppendLine("WHERE		" & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("GROUP BY	Renglones_Comprobantes.Tip_Ori, Comprobantes.Documento")

            
            
            loComandoSeleccionar.AppendLine("SELECT		#tmpTemporal1.*, ")
            loComandoSeleccionar.AppendLine("			#tmpTemporal2.*")
            loComandoSeleccionar.AppendLine("FROM		#tmpTemporal1 ")
            loComandoSeleccionar.AppendLine("	JOIN	#tmpTemporal2 ON (#tmpTemporal2.Doc_Tipo  =  #tmpTemporal1.Documento)")
            loComandoSeleccionar.AppendLine("ORDER BY	Tip_Ori, Documento ")
             
            



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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCDiario_Tipo_Origen", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCDiario_Tipo_Origen.ReportSource = loObjetoReporte

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
' JJD: 24/02/09: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' MAT: 03/03/11: Mejora en la vista de diseño												'
'-------------------------------------------------------------------------------------------'
' MAT: 04/03/11: Se aplicaron los metodos carga de imagen y validacion de registro cero		'
'-------------------------------------------------------------------------------------------'
' RJG: 08/11/11: Ajuste en el SELECT para que los renglones sin origen muestren una			'
'				 leyenda en lugar de aparecer en blanco.									'
'-------------------------------------------------------------------------------------------'		
