'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCotizaciones_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rCotizaciones_Renglones 
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
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))            
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.Appendline("SELECT	Cotizaciones.Documento, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Fec_Ini, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Fec_Fin, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Cod_Cli, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Cod_Ven, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Cod_Tra, "  )
			loComandoSeleccionar.Appendline("		Renglones_Cotizaciones.Renglon, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Cod_For, "  )
			loComandoSeleccionar.Appendline("		Renglones_Cotizaciones.Cod_Art, "  )
			loComandoSeleccionar.Appendline("		Renglones_Cotizaciones.Cod_Alm, "  )
			loComandoSeleccionar.Appendline("		Renglones_Cotizaciones.Precio1, "  )

            Select Case lcParametro8Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("             Renglones_Cotizaciones.Can_Art1, ")
                Case "Backorder"
                    loComandoSeleccionar.AppendLine("             Renglones_Cotizaciones.Can_Pen1 AS Can_Art1, ")
                Case "Procesado"
                    loComandoSeleccionar.AppendLine("             (Renglones_Cotizaciones.Can_Art1 - Renglones_Cotizaciones.Can_Pen1) AS Can_Art1, ")
            End Select

			loComandoSeleccionar.Appendline("		Renglones_Cotizaciones.Cod_Uni, "  )
			loComandoSeleccionar.Appendline("		Renglones_Cotizaciones.Por_Imp1, "  )
			loComandoSeleccionar.Appendline("		Renglones_Cotizaciones.Por_Des, "  )
			loComandoSeleccionar.Appendline("		Renglones_Cotizaciones.Mon_Net, "  )
			loComandoSeleccionar.Appendline("		Articulos.Nom_Art, "  )
			loComandoSeleccionar.Appendline("		Vendedores.Nom_Ven, "  )
			loComandoSeleccionar.Appendline("		Transportes.Nom_Tra, "  )
			loComandoSeleccionar.Appendline("		Formas_Pagos.Nom_For, "  )
			loComandoSeleccionar.Appendline("		Clientes.Nom_Cli "  )
			loComandoSeleccionar.Appendline("FROM	 Cotizaciones, "  )
			loComandoSeleccionar.Appendline("		Renglones_Cotizaciones, "  )
			loComandoSeleccionar.Appendline("		Clientes, "  )
			loComandoSeleccionar.Appendline("		Articulos, "  )
			loComandoSeleccionar.Appendline("		Vendedores, "  )
			loComandoSeleccionar.Appendline("		Formas_Pagos, "  )
			loComandoSeleccionar.Appendline("		Transportes "  )
			loComandoSeleccionar.Appendline("WHERE	Cotizaciones.Documento = Renglones_Cotizaciones.Documento "  )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Cod_Cli = Clientes.Cod_Cli "  )
			loComandoSeleccionar.Appendline("		And Renglones_Cotizaciones.Cod_Art = Articulos.Cod_Art "  )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Cod_Ven = Vendedores.Cod_Ven "  )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Cod_For = Formas_Pagos.Cod_For "  )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Cod_Tra = Transportes.Cod_Tra "  )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Documento Between " & lcParametro0Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro0Hasta )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Fec_Ini between " & lcParametro1Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro1Hasta )
			loComandoSeleccionar.Appendline("		And Renglones_Cotizaciones.Cod_Art between " & lcParametro2Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro2Hasta )	   			
			loComandoSeleccionar.Appendline("		And Clientes.Cod_Cli between " & lcParametro3Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro3Hasta )
			loComandoSeleccionar.Appendline("		And Vendedores.Cod_Ven between " & lcParametro4Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro4Hasta )
            loComandoSeleccionar.AppendLine("		And Cotizaciones.Status IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Cotizaciones.Cod_Rev between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("    	AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("       AND Cotizaciones.Cod_Suc between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	AND " & lcParametro7Hasta)
            'loComandoSeleccionar.AppendLine("ORDER BY Cotizaciones.Documento")
            loComandoSeleccionar.AppendLine("ORDER BY  Cotizaciones.Documento, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString , "curReportes")

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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rCotizaciones_Renglones", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrCotizaciones_Renglones.ReportSource =	 loObjetoReporte	

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
' MVP: 11/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  29/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  21/07/09: Filtro BackOrder, lo conllevo al anexo del campo Can_Pen1
'                 Metodeo de ordenamiento, verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS: 19/03/10: se agrego el filtro cod_art
'-------------------------------------------------------------------------------------------'
' CMS: 27/04/10: se agrego el filtro Vendedores
'-------------------------------------------------------------------------------------------'
' MAT: 04/04/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'