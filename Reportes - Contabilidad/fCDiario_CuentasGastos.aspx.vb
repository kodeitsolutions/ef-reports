'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCDiario_CuentasGastos"
'-------------------------------------------------------------------------------------------'
Partial Class fCDiario_CuentasGastos
     Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine("SELECT		Comprobantes.Documento							AS Documento,")
			loComandoSeleccionar.AppendLine(" 			CONVERT(nchar(15), Comprobantes.Fec_Ini,103)	AS Fec_Ini,")
			loComandoSeleccionar.AppendLine(" 			CONVERT(nchar(15), Comprobantes.Fec_Fin,103)	AS Fec_Fin,")
			loComandoSeleccionar.AppendLine(" 			Comprobantes.Status								AS Status,")
			loComandoSeleccionar.AppendLine(" 			Comprobantes.Resumen							AS Resumen,")
			loComandoSeleccionar.AppendLine(" 			Comprobantes.Tipo								AS Tipo")
			loComandoSeleccionar.AppendLine("INTO		#Temp1")
            loComandoSeleccionar.AppendLine("FROM		Comprobantes")
            loComandoSeleccionar.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("")
            
            
			loComandoSeleccionar.AppendLine("SELECT		Renglones_Comprobantes.Documento						AS Documento,")
			loComandoSeleccionar.AppendLine("			SUM(Renglones_Comprobantes.Mon_Deb)						AS Mon_Deb2,")
			loComandoSeleccionar.AppendLine("			SUM(Renglones_Comprobantes.Mon_Hab)						AS Mon_Hab2,")
			loComandoSeleccionar.AppendLine("			ISNULL(Cuentas_Gastos.Cod_Gas, '*Sin gasto')			AS Cod_Gas,")
            loComandoSeleccionar.AppendLine("			ISNULL(Cuentas_Gastos.Nom_Gas, '(Sin Cuenta de Gasto)')	AS Nom_Gas	")
			loComandoSeleccionar.AppendLine("INTO		#Temp2")
			loComandoSeleccionar.AppendLine("FROM		Comprobantes")
			loComandoSeleccionar.AppendLine("	JOIN	Renglones_Comprobantes ON Renglones_Comprobantes.Documento = Comprobantes.Documento")
			loComandoSeleccionar.AppendLine("		AND	Renglones_Comprobantes.Adicional = Comprobantes.Adicional")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Gastos ON Cuentas_Gastos.Cod_Gas = Renglones_Comprobantes.Cod_Gas")
            loComandoSeleccionar.AppendLine("WHERE		Comprobantes.Documento	=  Renglones_Comprobantes.Documento ")
            loComandoSeleccionar.AppendLine("       AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine("GROUP BY 	Renglones_Comprobantes.Documento,Cuentas_Gastos.Cod_Gas, Cuentas_Gastos.Nom_Gas")
			loComandoSeleccionar.AppendLine("")
			
			loComandoSeleccionar.AppendLine("SELECT #Temp1.Documento,")
			loComandoSeleccionar.AppendLine(" 		#Temp1.Fec_Ini,")
			loComandoSeleccionar.AppendLine(" 		#Temp1.Fec_Fin,")
			loComandoSeleccionar.AppendLine(" 		#Temp1.Status,")
			loComandoSeleccionar.AppendLine(" 		#Temp1.Resumen,")
			loComandoSeleccionar.AppendLine(" 		#Temp1.Tipo,")
			loComandoSeleccionar.AppendLine(" 		#Temp2.Mon_Deb2,")
			loComandoSeleccionar.AppendLine(" 		#Temp2.Mon_Hab2,")
			loComandoSeleccionar.AppendLine(" 		#Temp2.Cod_Gas,")
			loComandoSeleccionar.AppendLine(" 		#Temp2.Nom_Gas")
			loComandoSeleccionar.AppendLine("FROM #Temp1")
			loComandoSeleccionar.AppendLine("	JOIN #Temp2 ON #Temp2.Documento = #Temp1.Documento")
			loComandoSeleccionar.AppendLine("ORDER BY 	(CASE WHEN(Mon_Deb2>0) THEN 0 ELSE 1 END) ASC,")
			loComandoSeleccionar.AppendLine("			Cod_Gas ASC")
			loComandoSeleccionar.AppendLine("")
   
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

             
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCDiario_CuentasGastos", laDatosReporte)


            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCDiario_CuentasGastos.ReportSource = loObjetoReporte

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
' CMS: 16/09/09: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' MAT: 04/03/11: Ajuste del SELECT, el formato no mostraba información.						'
'-------------------------------------------------------------------------------------------'
' MAT: 04/03/11: Se aplicaron los metodos carga de imagen y validacion de registro cero		'
'-------------------------------------------------------------------------------------------'
' MAT: 04/03/11: Ajuste de la vista de diseño												'
'-------------------------------------------------------------------------------------------'
' RJG: 09/11/11: Ajuste en el JOIN para incluir los renglones que no tengan una cuenta de	'
'				 gastos asociada. Ajuste en el ordenamiento para que primero coloque los	'
'				 montos por el Debe y luego los del haber, y en segundo lugar ordena por	'
'				 código de Cuenta de Gastos.												'
'-------------------------------------------------------------------------------------------'
' RJG: 20/01/12: Se agregó el campo Adicional a la unión entre el encabezado y los renglones'
'-------------------------------------------------------------------------------------------'
