'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPedidos_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rPedidos_Renglones 
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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)
		
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.Appendline("SELECT	Pedidos.Documento, "  )
			loComandoSeleccionar.Appendline("		Pedidos.Fec_Ini, "  )
			loComandoSeleccionar.Appendline("		Pedidos.Fec_Fin, "  )
			loComandoSeleccionar.Appendline("		Clientes.Cod_Cli, "  )
			loComandoSeleccionar.Appendline("		Vendedores.Cod_Ven, "  )
			loComandoSeleccionar.Appendline("		Transportes.Cod_Tra, "  )
			loComandoSeleccionar.Appendline("		Renglones_Pedidos.Renglon, "  )
			loComandoSeleccionar.Appendline("		Formas_Pagos.Cod_For, "  )
			loComandoSeleccionar.Appendline("		Renglones_Pedidos.Cod_Art, "  )
			loComandoSeleccionar.Appendline("		Renglones_Pedidos.Cod_Alm, "  )
			loComandoSeleccionar.Appendline("		Renglones_Pedidos.Precio1, "  )

            Select Case lcParametro9Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Can_Art1, ")
                Case "Backorder"
                    loComandoSeleccionar.AppendLine("             Renglones_Pedidos.Can_Pen1 AS Can_Art1, ")
                Case "Procesado"
                    loComandoSeleccionar.AppendLine("             (Renglones_Pedidos.Can_Art1 - Renglones_Pedidos.Can_Pen1) AS Can_Art1, ")
            End Select

			loComandoSeleccionar.Appendline("		Renglones_Pedidos.Cod_Uni, "  )
			loComandoSeleccionar.Appendline("		Renglones_Pedidos.Por_Imp1, "  )
			loComandoSeleccionar.Appendline("		Renglones_Pedidos.Por_Des, "  )
			loComandoSeleccionar.Appendline("		Renglones_Pedidos.Mon_Net, "  )
			loComandoSeleccionar.Appendline("		Articulos.Nom_Art, "  )
			loComandoSeleccionar.Appendline("		Vendedores.Nom_Ven, "  )
			loComandoSeleccionar.Appendline("		Transportes.Nom_Tra, "  )
			loComandoSeleccionar.Appendline("		Formas_Pagos.Nom_For, "  )
			loComandoSeleccionar.Appendline("		Clientes.Nom_Cli "  )
			loComandoSeleccionar.Appendline("FROM	Pedidos, "  )
			loComandoSeleccionar.Appendline("		Renglones_Pedidos, "  )
			loComandoSeleccionar.Appendline("		Clientes, "  )
			loComandoSeleccionar.Appendline("		Articulos, "  )
			loComandoSeleccionar.Appendline("		Vendedores, "  )
			loComandoSeleccionar.Appendline("		Formas_Pagos, "  )
			loComandoSeleccionar.Appendline("		Transportes, "  )
			loComandoSeleccionar.Appendline("		Monedas "  )
			loComandoSeleccionar.Appendline("WHERE	Pedidos.Documento = Renglones_Pedidos.Documento "  )
			loComandoSeleccionar.Appendline(" 		AND Pedidos.Cod_Cli = Clientes.Cod_Cli "  )
			loComandoSeleccionar.Appendline(" 		AND Renglones_Pedidos.Cod_Art = Articulos.Cod_Art "  )
			loComandoSeleccionar.Appendline(" 		AND Pedidos.Cod_Ven = Vendedores.Cod_Ven "  )
			loComandoSeleccionar.Appendline(" 		AND Pedidos.Cod_For = Formas_Pagos.Cod_For "  )
			loComandoSeleccionar.Appendline(" 		AND Pedidos.Cod_Tra = Transportes.Cod_Tra "  )
            loComandoSeleccionar.AppendLine(" 		AND pedidos.Cod_mon = monedas.Cod_mon ")

            Select Case lcParametro9Desde
                Case "Backorder"
                    loComandoSeleccionar.AppendLine("             AND Renglones_Pedidos.Can_Pen1 <> 0 ")
            End Select

			loComandoSeleccionar.Appendline(" 		AND Pedidos.Documento Between " & lcParametro0Desde )
			loComandoSeleccionar.Appendline(" 		AND " & lcParametro0Hasta )
			loComandoSeleccionar.Appendline(" 		AND Pedidos.Fec_Ini BETWEEN " & lcParametro1Desde )
			loComandoSeleccionar.Appendline(" 		AND " & lcParametro1Hasta )
			loComandoSeleccionar.Appendline(" 		AND Clientes.Cod_Cli BETWEEN " & lcParametro2Desde )
			loComandoSeleccionar.Appendline(" 		AND " & lcParametro2Hasta )
			loComandoSeleccionar.Appendline(" 		AND Vendedores.Cod_ven BETWEEN " & lcParametro3Desde )
			loComandoSeleccionar.Appendline(" 		AND " & lcParametro3Hasta )
			loComandoSeleccionar.Appendline(" 		AND Pedidos.Status IN ( " & lcParametro4Desde & ")" )
			loComandoSeleccionar.Appendline(" 		AND Transportes.Cod_tra BETWEEN " & lcParametro5Desde )
			loComandoSeleccionar.Appendline(" 		AND " & lcParametro5Hasta )
			loComandoSeleccionar.Appendline(" 		AND Monedas.Cod_mon BETWEEN " & lcParametro6Desde )
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("       AND Pedidos.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("       AND Pedidos.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("    	AND " & lcParametro8Hasta)
            'loComandoSeleccionar.AppendLine("ORDER BY Pedidos.Documento")
            loComandoSeleccionar.AppendLine("ORDER BY  Pedidos.Documento, " & lcOrdenamiento)


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


            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rPedidos_Renglones", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrPedidos_Renglones.ReportSource =	 loObjetoReporte	

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
' JJD: 26/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  16/04/09: Cambios Estandarización de codigo. 
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  21/07/09: Filtro BackOrder, lo conllevo al anexo del campo Can_Pen1
'-------------------------------------------------------------------------------------------'
' CMS:  13/08/09: Se Agrego la restricción Renglones_Pedidos.Can_Pen1 <> 0 cuando el filtro 
'                   BackOrder = BackOrder
'-------------------------------------------------------------------------------------------'