'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_tStocks"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_tStocks
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

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try

			lcComandoSeleccionar.AppendLine(" SELECT    Articulos.Cod_Art, ")
            lcComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
			lcComandoSeleccionar.AppendLine("           Articulos.Exi_Act1, ")
			lcComandoSeleccionar.AppendLine("           Articulos.Exi_Ped1, ")
			lcComandoSeleccionar.AppendLine("           (Articulos.Exi_Act1 - Articulos.Exi_Ped1) Exi_Dis, ")
			lcComandoSeleccionar.AppendLine("           Articulos.Exi_Por1, ")
			lcComandoSeleccionar.AppendLine("           Articulos.Exi_Des1, ")
			lcComandoSeleccionar.AppendLine("           Articulos.Cod_Dep, ")
			lcComandoSeleccionar.AppendLine("           Articulos.Cod_Sec, ")
			lcComandoSeleccionar.AppendLine("           Articulos.Cod_Mar, ")
			lcComandoSeleccionar.AppendLine("           Articulos.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Cla ")
            lcComandoSeleccionar.AppendLine(" FROM      Articulos ")
            'lcComandoSeleccionar.AppendLine(" FROM      Articulos, ")
            'lcComandoSeleccionar.AppendLine("           Departamentos, ")
            'lcComandoSeleccionar.AppendLine("           Secciones, ")
            'lcComandoSeleccionar.AppendLine("           Marcas, ")
            'lcComandoSeleccionar.AppendLine("           Tipos_Articulos, ")
            'lcComandoSeleccionar.AppendLine("           Clases_Articulos ")

            Select Case lcParametro8Desde
                Case "Todos"
                    lcComandoSeleccionar.AppendLine(" WHERE     ")
                Case "Igual"
                    lcComandoSeleccionar.AppendLine(" WHERE     Articulos.Exi_Act1          =   0 AND ")
                Case "Mayor"
                    lcComandoSeleccionar.AppendLine(" WHERE     Articulos.Exi_Act1          >   0 AND ")
                Case "Menor"
                    lcComandoSeleccionar.AppendLine(" WHERE     Articulos.Exi_Act1          <   0 AND ")
            End Select

            'lcComandoSeleccionar.AppendLine("           Articulos.Cod_Dep               =   Departamentos.Cod_Dep ")
            'lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec           =   Secciones.Cod_Sec ")
            'lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar           =   Marcas.Cod_Mar ")
            'lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Tip           =   Tipos_Articulos.Cod_Tip ")
            'lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Cla           =   Clases_Articulos.Cod_Cla ")
            'lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Art           Between " & lcParametro0Desde)
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
            'lcComandoSeleccionar.AppendLine(" ORDER BY Articulos.Cod_Art, Articulos.Nom_Art ")
            lcComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


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


			loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rArticulos_tStocks", laDatosReporte)
			
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
	
            Me.crvrArticulos_tStocks.ReportSource = loObjetoReporte

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
' MVP: 10/07/08: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' MVP: 11/07/08: Adición de loObjetoReporte para eliminar los archivos temp en Uranus		'
'-------------------------------------------------------------------------------------------'
' JJD: 15/10/08: Cambio de parametros para su correcta ejecucion							'
'-------------------------------------------------------------------------------------------'
' CMS: 05/05/09: Ordenamiento																'
'-------------------------------------------------------------------------------------------'
' CMS:  11/08/09: Verificacion de registros y se agregaro el filtro_ Ubicación				'
'-------------------------------------------------------------------------------------------'
' CMS:  11/08/09: Se agrego el filtro existencia con la finalidad de unificar en un solo	'
'                 reporte los siguientes 4 reportes:										'
'                    - Listado de Artículos con Todos sus Stocks (Menor a Cero)				'
'                    - Listado de Artículos con Todos sus Stocks (Mayor a Cero)				'
'                    - Listado de Artículos con Todos sus Stocks (Igual a Cero)				'
'                    - Listado de Artículos con Todos sus Stocks							'
'                 Se eliminaron las uniones innecesarias, para evitar time out.				'
'-------------------------------------------------------------------------------------------'
' CMS:  07/06/10: Se agrrego la columna de articulos disponibles							'
'					(Articulos.Exi_Act1 - Articulos.Exi_Ped1) Exi_Dis						'
'-------------------------------------------------------------------------------------------'