'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDCLientes_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rDCLientes_Articulos 
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

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
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

        
			Dim loComandoSeleccionar As New StringBuilder()
			 
			loComandoSeleccionar.AppendLine( " SELECT		Articulos.Cod_Art, " ) 
			loComandoSeleccionar.AppendLine( "				Articulos.Nom_art, " )
			loComandoSeleccionar.AppendLine( "				Articulos.Cod_Mar, " )
			loComandoSeleccionar.AppendLine( " 				Articulos.Status, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.Documento, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.Cod_Cli, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.Fec_Ini, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.Cod_Ven, " )
			loComandoSeleccionar.AppendLine( " 				Renglones_dClientes.Renglon, " )
			loComandoSeleccionar.AppendLine( " 				Renglones_dClientes.Cod_Alm, " )
			loComandoSeleccionar.AppendLine( " 				Renglones_dClientes.Can_Art1, " )
			loComandoSeleccionar.AppendLine( " 				Renglones_dClientes.Cod_Uni, " )
			loComandoSeleccionar.AppendLine( " 				Renglones_dClientes.Precio1, " )
			loComandoSeleccionar.AppendLine( " 				Renglones_dClientes.Por_Des, " )
			loComandoSeleccionar.AppendLine( " 				Renglones_dClientes.Mon_Net " )
			loComandoSeleccionar.AppendLine( " FROM			Articulos, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes, " )
			loComandoSeleccionar.AppendLine( " 				Renglones_dClientes, " )
			loComandoSeleccionar.AppendLine( " 				Clientes, " )
			loComandoSeleccionar.AppendLine( " 				Vendedores, " )
			loComandoSeleccionar.AppendLine( " 				Almacenes, " )
			loComandoSeleccionar.AppendLine( " 				Marcas " )
			loComandoSeleccionar.AppendLine( " WHERE		Articulos.Cod_Art = Renglones_dClientes.Cod_Art ")
            loComandoSeleccionar.AppendLine("             AND 			Renglones_dClientes.Documento = Devoluciones_Clientes.Documento ")
            loComandoSeleccionar.AppendLine("             AND 			Articulos.Cod_Mar = Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("             AND 			Devoluciones_Clientes.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("             AND 			Devoluciones_Clientes.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("             AND 			Renglones_dClientes.Cod_Alm = Almacenes.Cod_Alm ")
            loComandoSeleccionar.AppendLine("             AND 			Renglones_dClientes.Cod_Art between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("             AND 			" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("             AND 			Devoluciones_Clientes.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("             AND 			" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("             AND 			Devoluciones_Clientes.Cod_Cli between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("             AND 			" & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("             AND 			Devoluciones_Clientes.Cod_Ven between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("             AND 			" & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("             AND 			Articulos.Cod_Mar  between" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("             AND 			" & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("             AND 			Devoluciones_Clientes.Status IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("             AND 			Almacenes.Cod_Alm between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("             AND 			" & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("             AND           Devoluciones_Clientes.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	      AND           " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("             AND 		    Devoluciones_Clientes.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("             AND 			" & lcParametro8Hasta)

            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY		Articulos.Cod_Art")



       Dim loServicios As New cusDatos.goDatos

       Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

       loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rDClientes_Articulos", laDatosReporte)
       
       Me.mTraducirReporte(loObjetoReporte)
       
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDClientes_Articulos.ReportSource =	 loObjetoReporte	


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
' MJP: 16/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
