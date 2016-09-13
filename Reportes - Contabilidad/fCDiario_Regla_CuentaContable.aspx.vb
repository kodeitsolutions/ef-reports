'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCDiario_Regla_CuentaContable"
'-------------------------------------------------------------------------------------------'
Partial Class fCDiario_Regla_CuentaContable
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
			
			loComandoSeleccionar.AppendLine("SELECT		Renglones_Comprobantes.Documento					AS Documento,")
			loComandoSeleccionar.AppendLine(" 			SUM(Renglones_Comprobantes.Mon_Deb) 				AS Mon_Deb,")
			loComandoSeleccionar.AppendLine(" 			SUM(Renglones_Comprobantes.Mon_Hab) 				AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("			ISNULL(Reglas_Integracion.Nom_Reg, '(Sin Regla)')	AS Cod_Reg, ")
            loComandoSeleccionar.AppendLine("			ISNULL(Reglas_Integracion.Nom_Reg, '(Sin Regla)')	AS Nom_Reg, ")
			loComandoSeleccionar.AppendLine(" 			Cuentas_Contables.Nom_Cue							AS Nom_Cue,")
			loComandoSeleccionar.AppendLine(" 			Cuentas_Contables.Cod_Cue							As Cod_Cue ")
			loComandoSeleccionar.AppendLine("INTO		#tmpTemporal2 ")
			loComandoSeleccionar.AppendLine("FROM		Comprobantes")
			loComandoSeleccionar.AppendLine("	JOIN 	Renglones_Comprobantes ON Renglones_Comprobantes.Documento	=	Comprobantes.Documento")
			loComandoSeleccionar.AppendLine("		AND	Renglones_Comprobantes.Adicional = Comprobantes.Adicional")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Reglas_Integracion ON (Reglas_Integracion.Cod_Reg = Renglones_Comprobantes.Cod_Reg)")
			loComandoSeleccionar.AppendLine("	JOIN 	Cuentas_Contables ON Cuentas_Contables.Cod_Cue	=	Renglones_Comprobantes.Cod_Cue   ")
			loComandoSeleccionar.AppendLine("WHERE		" & cusAplicacion.goFormatos.pcCondicionPrincipal)  
			loComandoSeleccionar.AppendLine("GROUP BY 	Renglones_Comprobantes.Documento,Renglones_Comprobantes.Cod_Reg,Reglas_Integracion.Nom_Reg,Cuentas_Contables.Cod_Cue, Cuentas_Contables.Nom_Cue")
			loComandoSeleccionar.AppendLine("ORDER BY 	Renglones_Comprobantes.Cod_Reg ASC,Cuentas_Contables.Cod_Cue ASC	")
			
		
            
            loComandoSeleccionar.AppendLine("SELECT		#tmpTemporal1.*, ")
            loComandoSeleccionar.AppendLine("			#tmpTemporal2.*")
            loComandoSeleccionar.AppendLine("FROM		#tmpTemporal1 ")
            loComandoSeleccionar.AppendLine("	JOIN	#tmpTemporal2 ON (#tmpTemporal2.Documento  =  #tmpTemporal1.Documento)")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCDiario_Regla_CuentaContable", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCDiario_Regla_CuentaContable.ReportSource = loObjetoReporte

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
' MAT: 09/08/11: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' RJG: 09/11/11: Ajuste para que aparezcan los renglones que no tienen una regla de			'
'				 integración asociada.														'
'-------------------------------------------------------------------------------------------'
' RJG: 20/01/12: Se agregó el campo Adicional a la unión entre el encabezado y los renglones'
'-------------------------------------------------------------------------------------------'
