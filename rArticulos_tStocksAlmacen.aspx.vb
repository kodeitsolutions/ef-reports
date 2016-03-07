'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_tStocksAlmacen"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_tStocksAlmacen
   Inherits vis2Formularios.frmReporte
   
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosFinales(8)
        Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro11Desde As String = cusAplicacion.goReportes.paParametrosIniciales(11)
        Dim lcParametro12Desde As String = cusAplicacion.goReportes.paParametrosIniciales(12)
        

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try

			lcComandoSeleccionar.AppendLine(" SELECT    Articulos.Cod_Art, ")
            lcComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
			lcComandoSeleccionar.AppendLine("           Renglones_Almacenes.Exi_Act1, ")
			lcComandoSeleccionar.AppendLine("           Renglones_Almacenes.Exi_Ped1, ")
            lcComandoSeleccionar.AppendLine("           (Renglones_Almacenes.Exi_Act1 - Renglones_Almacenes.Exi_Ped1) AS Exi_Dis, ")
			lcComandoSeleccionar.AppendLine("           Renglones_Almacenes.Exi_Por1, ")
			lcComandoSeleccionar.AppendLine("           Renglones_Almacenes.Exi_Des1, ")
			lcComandoSeleccionar.AppendLine("           Almacenes.Cod_Alm, ")
			
			Select Case lcParametro11Desde
			
				Case "Todos"
							lcComandoSeleccionar.AppendLine("'Todos' As Visible, ")
								
				Case "Actual"
							lcComandoSeleccionar.AppendLine("'Actual' As Visible, ")	
							
				Case "Disponible"	
							lcComandoSeleccionar.AppendLine("'Disponible' As Visible, ")	
							
				Case "Comprometida"	
							lcComandoSeleccionar.AppendLine("'Comprometida' As Visible, ")	
							
				Case "Por_Llegar"	
							lcComandoSeleccionar.AppendLine("'Por_Llegar' As Visible, ")	
						
				Case "Por_Despachar"	
							lcComandoSeleccionar.AppendLine("'Por_Despachar' As Visible, ")	
						
			End Select
			
			Select Case lcParametro12Desde
				Case "Si"	
							lcComandoSeleccionar.AppendLine("'Si' As Color, ")	
				Case "No"	
							lcComandoSeleccionar.AppendLine("'No' As Color, ")	
			End Select 
			
			lcComandoSeleccionar.AppendLine("           Almacenes.Nom_Alm, ")
			lcComandoSeleccionar.AppendLine("           Articulos.Cod_Dep, ")
			lcComandoSeleccionar.AppendLine("           Articulos.Cod_Sec, ")
			lcComandoSeleccionar.AppendLine("           Articulos.Cod_Mar, ")
			lcComandoSeleccionar.AppendLine("           Articulos.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Cla, ")            
            lcComandoSeleccionar.AppendLine("           Case when Renglones_Almacenes.Exi_Act1 > Articulos.Exi_Max then 1 else 0 end As Exi_Max, ")
            lcComandoSeleccionar.AppendLine("           Case when Renglones_Almacenes.Exi_Act1 < Articulos.Exi_Min then 1 else 0 end As Exi_Min  ")            
            lcComandoSeleccionar.AppendLine(" FROM      Articulos ")
            lcComandoSeleccionar.AppendLine(" JOIN Renglones_Almacenes on Renglones_Almacenes.Cod_Art = Articulos.Cod_Art ")
            lcComandoSeleccionar.AppendLine(" JOIN Almacenes on Almacenes.Cod_Alm = Renglones_Almacenes.Cod_Alm ")

            Select Case lcParametro8Desde
                Case "Todos"
                    lcComandoSeleccionar.AppendLine(" WHERE     ")
                Case "Igual"
                    lcComandoSeleccionar.AppendLine(" WHERE     (Renglones_Almacenes.Exi_Act1 - Renglones_Almacenes.Exi_Ped1)          =   0 AND ")
                Case "Mayor"
                    lcComandoSeleccionar.AppendLine(" WHERE     (Renglones_Almacenes.Exi_Act1 - Renglones_Almacenes.Exi_Ped1)          >   0 AND ")
                Case "Menor"
                    lcComandoSeleccionar.AppendLine(" WHERE    (Renglones_Almacenes.Exi_Act1 - Renglones_Almacenes.Exi_Ped1)         <   0 AND ")
            End Select
            

            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Art           Between " & lcParametro0Desde)
			lcComandoSeleccionar.AppendLine("           And "  & lcParametro0Hasta)
			lcComandoSeleccionar.AppendLine("           And Articulos.Status            IN (" & lcParametro1Desde & ")")
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep       Between " & lcParametro2Desde)
			lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec           Between " & lcParametro3Desde)
			lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar              Between " & lcParametro4Desde)
			lcComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Tip     Between " & lcParametro5Desde)
			lcComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Cla    Between " & lcParametro6Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Ubi    Between " & lcParametro7Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            lcComandoSeleccionar.AppendLine("           And Renglones_Almacenes.Cod_Alm    Between " & lcParametro9Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Pro    Between " & lcParametro10Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro10Hasta)
            lcComandoSeleccionar.AppendLine("ORDER BY   Almacenes.Cod_Alm,  " & lcOrdenamiento)

            ' Me.mEscribirConsulta(lcComandoSeleccionar.ToString)
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

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


			loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rArticulos_tStocksAlmacen", laDatosReporte)
			
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
	
            Me.crvrArticulos_tStocksAlmacen.ReportSource = loObjetoReporte

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
' CMS: 23/06/10: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' MAT: 11/07/11: Reprogramación del reporte, Ajuste de la vista de diseño					'
'-------------------------------------------------------------------------------------------'
' MAT: 29/09/11: Ajuste del Filtro de Exi_Act1 -Exi_Ped1, según requerimientos				'
'-------------------------------------------------------------------------------------------'
