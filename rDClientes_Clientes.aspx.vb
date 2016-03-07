Imports System.Data
Partial Class rDClientes_Clientes 
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine( "  SELECT		Clientes.Cod_Cli, " ) 
			loComandoSeleccionar.AppendLine( "				Clientes.Nom_Cli, " )
			loComandoSeleccionar.AppendLine( "				Clientes.Status, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.Documento, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.fec_ini, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.fec_fin, " ) 
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.Cod_Ven, " ) 
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.Cod_Tra, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.Cod_Mon, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.Control, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.Mon_Imp1, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.Mon_Net, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes.Mon_Bru  " ) 
			loComandoSeleccionar.AppendLine( " From			Clientes, " )
			loComandoSeleccionar.AppendLine( " 				Devoluciones_Clientes, " )
			loComandoSeleccionar.AppendLine( " 				Vendedores, " )
			loComandoSeleccionar.AppendLine( " 				Transportes, " )
			loComandoSeleccionar.AppendLine( " 				Monedas " )
			loComandoSeleccionar.AppendLine( " WHERE		Clientes.Cod_Cli = Devoluciones_Clientes.Cod_Cli " )
			loComandoSeleccionar.AppendLine( " 				AND Devoluciones_Clientes.Cod_Ven = Vendedores.Cod_Ven " )
			loComandoSeleccionar.AppendLine( " 				AND Devoluciones_Clientes.Cod_Tra = Transportes.Cod_Tra " )
			loComandoSeleccionar.AppendLine( " 				AND Devoluciones_Clientes.Cod_Mon = Monedas.Cod_Mon " )
			loComandoSeleccionar.AppendLine( " 				AND Devoluciones_Clientes.Documento between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine( " 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine( " 				AND Devoluciones_Clientes.Fec_Ini between " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine( " 				AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine( " 				AND Devoluciones_Clientes.Cod_Cli between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine( " 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine( " 				AND Devoluciones_Clientes.Cod_Ven between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine( " 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine( " 				AND Devoluciones_Clientes.Status IN (" & lcParametro4Desde & ")")
			loComandoSeleccionar.AppendLine( " 				AND Devoluciones_Clientes.Cod_Tra between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine( " 				AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine( " 				AND Devoluciones_Clientes.Cod_Mon between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("               AND	Devoluciones_Clientes.Cod_Suc between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("               AND " & lcParametro7Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY		Clientes.Cod_Cli, Devoluciones_Clientes.Fec_Ini, Devoluciones_Clientes.Fec_Fin, Devoluciones_Clientes.Cod_Ven ")
            loComandoSeleccionar.AppendLine("ORDER BY       Clientes.Cod_Cli, " & lcOrdenamiento)

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


            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rDClientes_Clientes", laDatosReporte)
            
			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDClientes_Clientes.ReportSource =	 loObjetoReporte	


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
' MJP: 17/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  20/03/09: Estandarizacion de codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  13/07/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  05/08/09: Verificacion de registros
'-------------------------------------------------------------------------------------------'
