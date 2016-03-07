'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAPrecios_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rAPrecios_Fechas 
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		Dim loComandoSeleccionar As New StringBuilder()
       

            loComandoSeleccionar.AppendLine(" SELECT		Ajustes_Precios.Documento, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes_Precios.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes_Precios.Comentario, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes_Precios.Status, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes_Precios.Cod_Mon, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes_Precios.Tasa, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_APrecios.Renglon, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_APrecios.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_APrecios.Pre_Ant, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_APrecios.Pre_Nue, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_APrecios.Tip_Pre, ")
            loComandoSeleccionar.AppendLine(" 				Articulos.Nom_Art, ")				 
            loComandoSeleccionar.AppendLine(" 				Monedas.Nom_Mon ")
            loComandoSeleccionar.AppendLine(" FROM			Ajustes_Precios")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Renglones_APrecios ON Renglones_APrecios.Documento = Ajustes_Precios.Documento ")
			loComandoSeleccionar.AppendLine("				AND Renglones_APrecios.Cod_Art between " & lcParametro2Desde & "AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" LEFT JOIN Monedas ON (Ajustes_Precios.Cod_Mon	=	Monedas.Cod_Mon)  ")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Articulos ON (Articulos.Cod_Art	=	Renglones_APrecios.Cod_Art )") 	
            loComandoSeleccionar.AppendLine(" WHERE			Ajustes_Precios.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Ajustes_Precios.Fec_Ini		Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Ajustes_Precios.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Ajustes_Precios.Cod_Rev		Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY       Convert(nchar(30), Ajustes_Precios.Fec_Ini, 112) DESC, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
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


            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rAPrecios_Fechas", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrAPrecios_Fechas.ReportSource =	 loObjetoReporte	

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
' GMO: 07/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GMO: 24/03/09: Estandarizacion de codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS:  14/07/09: Verificacion de registros, :etodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' MAT:  01/02/11: Modificación del Código. No mostraba documentos.
'-------------------------------------------------------------------------------------------'