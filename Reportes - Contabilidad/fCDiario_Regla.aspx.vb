Imports System.Data
Partial Class fCDiario_Regla

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

           
     
			loComandoSeleccionar.AppendLine("SELECT		Comprobantes.Documento			AS Documento, ")
            loComandoSeleccionar.AppendLine("			YEAR(Comprobantes.Fec_Ini)  	AS Anno, ")
            loComandoSeleccionar.AppendLine("			MONTH(Comprobantes.Fec_Ini) 	AS Mes, ")
            loComandoSeleccionar.AppendLine("			Comprobantes.Fec_Ini 			AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Comprobantes.Fec_Fin 			AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine("			Comprobantes.Resumen			AS Resumen, ")
            loComandoSeleccionar.AppendLine("			Comprobantes.Tipo				As Tipo, ")
            loComandoSeleccionar.AppendLine("			Comprobantes.Origen				AS Origen, ")
            loComandoSeleccionar.AppendLine("			Comprobantes.Integracion		AS Integracion, ")
            loComandoSeleccionar.AppendLine("			Comprobantes.Status				As Status")
            loComandoSeleccionar.AppendLine("INTO		#tmpTemporal1 ")
            loComandoSeleccionar.AppendLine("FROM		Comprobantes ")
            loComandoSeleccionar.AppendLine("WHERE		" & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

           
            loComandoSeleccionar.AppendLine("SELECT		Comprobantes.Documento								AS Doc_Tipo, ")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Comprobantes.Mon_Deb)					AS Mon_Deb, ")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Comprobantes.Mon_Hab)					AS Mon_hab, ")
            loComandoSeleccionar.AppendLine("			ISNULL(Reglas_Integracion.Nom_Reg, '(Sin Regla)')	AS Cod_Reg, ")
            loComandoSeleccionar.AppendLine("			ISNULL(Reglas_Integracion.Nom_Reg, '(Sin Regla)')	AS Nom_Reg ")
            loComandoSeleccionar.AppendLine("INTO		#tmpTemporal2")
            loComandoSeleccionar.AppendLine("FROM		Comprobantes ")
            loComandoSeleccionar.AppendLine("	JOIN 	Renglones_Comprobantes")
            loComandoSeleccionar.AppendLine("		ON	(Comprobantes.Documento = Renglones_Comprobantes.Documento)")
			loComandoSeleccionar.AppendLine("		AND	(Renglones_Comprobantes.Adicional = Comprobantes.Adicional)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Reglas_Integracion")
            loComandoSeleccionar.AppendLine("		ON	(Reglas_Integracion.Cod_Reg = Renglones_Comprobantes.Cod_Reg)")
            
            loComandoSeleccionar.AppendLine("WHERE		" & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("GROUP BY	Renglones_Comprobantes.Cod_Reg,Reglas_Integracion.Nom_Reg,Comprobantes.Documento")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            
            
            loComandoSeleccionar.AppendLine("SELECT		#tmpTemporal1.*, ")
            loComandoSeleccionar.AppendLine("			#tmpTemporal2.*")
            loComandoSeleccionar.AppendLine("FROM		#tmpTemporal1 ")
            loComandoSeleccionar.AppendLine("	JOIN	#tmpTemporal2 ON (#tmpTemporal2.Doc_Tipo  =  #tmpTemporal1.Documento)")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCDiario_Regla", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCDiario_Regla.ReportSource = loObjetoReporte

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
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 24/02/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 03/03/11: Mejora en la vista de diseño
'-------------------------------------------------------------------------------------------'
' MAT:  04/03/11 : Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' RJG: 09/11/11: Ajuste para que aparezcan los renglones que no tienen una regla de			'
'				 integración asociada.														'
'-------------------------------------------------------------------------------------------'
' RJG: 20/01/12: Se agregó el campo Adicional a la unión entre el encabezado y los renglones'
'-------------------------------------------------------------------------------------------'
