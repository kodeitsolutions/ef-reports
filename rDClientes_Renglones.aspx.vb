'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDClientes_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rDClientes_Renglones
     Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
		
        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
	Try
	
		Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT	Devoluciones_Clientes.Documento, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Fec_Ini, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Fec_Fin, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Cod_Cli, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Cod_Ven, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Cod_Tra, ")
			loComandoSeleccionar.AppendLine("			Renglones_DClientes.Renglon, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Cod_For, ")
			loComandoSeleccionar.AppendLine("			Renglones_DClientes.Cod_Art, ")
			loComandoSeleccionar.AppendLine("			Renglones_DClientes.Cod_Alm, ")
			loComandoSeleccionar.AppendLine("			Renglones_DClientes.Precio1, ")
			loComandoSeleccionar.AppendLine("			Renglones_DClientes.Can_Art1, ")
			loComandoSeleccionar.AppendLine("			Renglones_DClientes.Cod_Uni, ")
			loComandoSeleccionar.AppendLine("			Renglones_DClientes.Por_Imp1, ")
			loComandoSeleccionar.AppendLine("			Renglones_DClientes.Por_Des, ")
			loComandoSeleccionar.AppendLine("			Renglones_DClientes.Mon_Net, ")
			loComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
			loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven, ")
			loComandoSeleccionar.AppendLine("			Transportes.Nom_Tra, ")
			loComandoSeleccionar.AppendLine("			Formas_Pagos.Nom_For, ")
			loComandoSeleccionar.AppendLine("			Clientes.Nom_Cli ")
			loComandoSeleccionar.AppendLine(" FROM		Devoluciones_Clientes, ")
			loComandoSeleccionar.AppendLine("			Renglones_DClientes, ")
			loComandoSeleccionar.AppendLine("			Clientes, ")
			loComandoSeleccionar.AppendLine("			Articulos, ")
			loComandoSeleccionar.AppendLine("			Vendedores, ")
			loComandoSeleccionar.AppendLine("			Formas_Pagos, ")
			loComandoSeleccionar.AppendLine("			Transportes ")
			loComandoSeleccionar.AppendLine(" WHERE		Devoluciones_Clientes.Documento			=	Renglones_DClientes.Documento ")
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Cod_Cli		=	Clientes.Cod_Cli ")
			loComandoSeleccionar.AppendLine("			AND Renglones_DClientes.Cod_Art			=	Articulos.Cod_Art ")
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Cod_Ven		=	Vendedores.Cod_Ven ")
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Cod_For		=	Formas_Pagos.Cod_For ")
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Cod_Tra		=	Transportes.Cod_Tra ")
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Documento		BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Fec_Ini		BETWEEN " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("			AND Clientes.Cod_Cli					BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Status		IN ( " & lcParametro3Desde & " )")
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Cod_Rev between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Cod_Suc between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)


            'loComandoSeleccionar.AppendLine(" ORDER BY	Devoluciones_Clientes.Documento")
            loComandoSeleccionar.AppendLine("ORDER BY   Devoluciones_Clientes.Documento, " & lcOrdenamiento)


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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rDClientes_Renglones", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte) 

            Me.crvrDClientes_Renglones.ReportSource =	 loObjetoReporte	

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
' JJD: 13/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' MAT:  18/02/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
