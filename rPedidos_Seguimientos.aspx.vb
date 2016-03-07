'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPedidos_Seguimientos "
'-------------------------------------------------------------------------------------------'
Partial Class rPedidos_Seguimientos 
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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
	    	Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))  

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			 Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT		Renglones_Pedidos.Cod_Art, " ) 
			loComandoSeleccionar.AppendLine(" 				Articulos.Nom_art, " )
			loComandoSeleccionar.AppendLine(" 				Pedidos.Documento, " )
			loComandoSeleccionar.AppendLine(" 				Pedidos.Cod_Cli, " )
			loComandoSeleccionar.AppendLine(" 				Clientes.Nom_Cli, " )
			loComandoSeleccionar.AppendLine(" 				Pedidos.Fec_Ini, " )
			loComandoSeleccionar.AppendLine(" 				Pedidos.Notas, " )
			loComandoSeleccionar.AppendLine(" 				Pedidos.Cod_Ven, " )
			loComandoSeleccionar.AppendLine(" 				Vendedores.Nom_Ven, " )
			loComandoSeleccionar.AppendLine(" 				Renglones_Pedidos.Can_Art1, " )
			loComandoSeleccionar.AppendLine(" 				Renglones_Pedidos.Cod_Uni, " )
			loComandoSeleccionar.AppendLine(" 				Renglones_Pedidos.Precio1, " )
			loComandoSeleccionar.AppendLine(" 				Renglones_Pedidos.Mon_Net " )
			loComandoSeleccionar.AppendLine(" FROM			Articulos, " )
			loComandoSeleccionar.AppendLine(" 				Pedidos, " )
			loComandoSeleccionar.AppendLine(" 				Renglones_Pedidos, " )
            loComandoSeleccionar.AppendLine(" 				Clientes, ")
            loComandoSeleccionar.AppendLine(" 				Vendedores ")
            'loComandoSeleccionar.AppendLine(" 				Vendedores, " )
            'loComandoSeleccionar.AppendLine(" 				Monedas, " )
            'loComandoSeleccionar.AppendLine(" 				Transportes, " )
            'loComandoSeleccionar.AppendLine(" 				Almacenes, " )
            'loComandoSeleccionar.AppendLine(" 				Departamentos, " )
            'loComandoSeleccionar.AppendLine(" 				Clases_Articulos " )
			loComandoSeleccionar.AppendLine(" WHERE			Articulos.Cod_Art = Renglones_Pedidos.Cod_Art " )
			loComandoSeleccionar.AppendLine(" 				AND Renglones_Pedidos.Documento = Pedidos.Documento "  )
			loComandoSeleccionar.AppendLine(" 				AND Pedidos.Cod_Cli = Clientes.Cod_Cli "  )
			loComandoSeleccionar.AppendLine(" 				AND Pedidos.Cod_Ven = Vendedores.Cod_Ven" )		
            'loComandoSeleccionar.AppendLine(" 				AND Pedidos.Cod_Mon = Monedas.Cod_Mon" )								
            'loComandoSeleccionar.AppendLine(" 				AND Pedidos.Cod_Tra = Transportes.Cod_Tra" )	
            'loComandoSeleccionar.AppendLine(" 				AND Renglones_Pedidos.Cod_Alm = Almacenes.Cod_Alm" )
            'loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Dep = Departamentos.Cod_Dep " )
            'loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Cla = Clases_Articulos.Cod_Cla " )
			loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,4) <>	'VIA-' " ) 		
			loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,4) <> 'GEN-' " )		
			loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,5) <> 'NOTAS' " )		
			loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,7) <> 'SES-CAN' " )	
			loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,7) <> 'INC-CAN' " )	
			loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,7) <> 'INC-PER' " )	
			loComandoSeleccionar.AppendLine(" 				AND SUBSTRING(Articulos.Cod_Art,1,4) <> 'REQ-' " )
			loComandoSeleccionar.AppendLine(" 				AND Renglones_Pedidos.Cod_Art between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Pedidos.Fec_Ini between " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Pedidos.Cod_Cli between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Pedidos.Cod_Ven between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Dep  between" & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Cla  between" & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Pedidos.Status IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Pedidos.Cod_Alm between " & lcParametro7Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro7Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Pedidos.Cod_Mon between " & lcParametro8Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro8Hasta)
			loComandoSeleccionar.AppendLine(" 				AND (( " & lcParametro9Desde & " = 'SI' AND Datepart(dw, Pedidos.fec_ini)>= 7) " )
			loComandoSeleccionar.AppendLine("				Or ( " & lcParametro9Desde & " <> 'SI' AND Datepart(dw, Pedidos.fec_ini)< 7))"  )
            'loComandoSeleccionar.AppendLine(" ORDER BY		Pedidos.Documento,Renglones_Pedidos.Cod_Art, Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("ORDER BY      Pedidos.Cod_Ven, " & lcOrdenamiento)


   
   '& " And Pedidos.Status between " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))  _
			'					& " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))  _

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rPedidos_Seguimientos", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrPedidos_Seguimientos.ReportSource =	 loObjetoReporte	


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
' MVP: 12/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GCR: 10/03/09: Estandarizacion de codigo y ajustes en el diseño
'-------------------------------------------------------------------------------------------'
' CMS:  31/08/09: Metodo de ordenamiento, verificacionde registros, se quitaron las 
'                 siguientes tablas: Monedas, Transportes, Almacenes, Departamentos,
'                 Clases_Articulos, para mejorar los tiempo de respuesta de la consulta
'-------------------------------------------------------------------------------------------'