'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAPrecios_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rAPrecios_Articulos 
    Inherits vis2formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

	Try
	
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim loComandoSeleccionar As New StringBuilder()
       
    

            loComandoSeleccionar.AppendLine(" SELECT		Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("				Ajustes_Precios.Status, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes_Precios.Documento, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes_Precios.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_APrecios.Renglon, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_APrecios.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_APrecios.Pre_Ant, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_APrecios.Pre_Nue, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_APrecios.Tip_Pre ")
            loComandoSeleccionar.AppendLine(" FROM			Articulos, ")
            loComandoSeleccionar.AppendLine("				Ajustes_Precios, ")
            loComandoSeleccionar.AppendLine("				Renglones_APrecios ")
            loComandoSeleccionar.AppendLine(" WHERE			Ajustes_Precios.Documento		=	Renglones_APrecios.Documento ")
            loComandoSeleccionar.AppendLine(" 				And Articulos.Cod_Art			=	Renglones_APrecios.Cod_Art ")
            loComandoSeleccionar.AppendLine(" 				And Renglones_APrecios.Cod_Art	Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				And Ajustes_Precios.Documento	Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				And Ajustes_Precios.Fec_Ini		Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				And Ajustes_Precios.Status		IN(" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("      	        AND Ajustes_Precios.Cod_Rev between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 		        AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY       Renglones_APrecios.Cod_Art, " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY		Renglones_APrecios.Cod_Art, Ajustes_Precios.Fec_Ini, Ajustes_Precios.Documento ")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rAPrecios_Articulos", laDatosReporte)
            
            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.crvrAPrecios_Articulos.ReportSource =	 loObjetoReporte	


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
' GCR: 11/03/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS:  15/07/09: Metodo de Ordenamiento, Verificacion de registros
'-------------------------------------------------------------------------------------------'