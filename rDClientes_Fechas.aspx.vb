'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDClientes_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rDClientes_Fechas 
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
			Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
        
			Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT Devoluciones_Clientes.Documento, ")
            loComandoSeleccionar.AppendLine("                 Devoluciones_Clientes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("                 Devoluciones_Clientes.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("                 Devoluciones_Clientes.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("                 Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("                 Devoluciones_Clientes.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("                 Devoluciones_Clientes.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("                 Devoluciones_Clientes.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("                 Devoluciones_Clientes.Control, ")
            loComandoSeleccionar.AppendLine("                 Devoluciones_Clientes.Mon_Net, ")
            loComandoSeleccionar.AppendLine("                 Devoluciones_Clientes.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("                 Devoluciones_Clientes.Mon_Bru  ")
            loComandoSeleccionar.AppendLine(" From            Clientes, ")
            loComandoSeleccionar.AppendLine("                 Devoluciones_Clientes, ")
            loComandoSeleccionar.AppendLine("                 Vendedores, ")
            loComandoSeleccionar.AppendLine("                 Transportes, ")
            loComandoSeleccionar.AppendLine("                 Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE           Devoluciones_Clientes.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("                 And Devoluciones_Clientes.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("                 And Devoluciones_Clientes.Cod_Tra = Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("                 And Devoluciones_Clientes.Cod_Mon = Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine("                 And Devoluciones_Clientes.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("                 And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("                 And Devoluciones_Clientes.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("                 And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("                 And Devoluciones_Clientes.Cod_Cli between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("                 And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("                 And Devoluciones_Clientes.Cod_Ven between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("                 And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("                 And Devoluciones_Clientes.Status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("                 And Devoluciones_Clientes.Cod_Tra between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("                 And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("                 And Devoluciones_Clientes.Cod_Mon between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("                 And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("                 AND Devoluciones_Clientes.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	          AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("                 AND Devoluciones_Clientes.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro8Hasta)

            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Devoluciones_Clientes.Documento ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rDClientes_Fechas", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDClientes_Fechas.ReportSource =	 loObjetoReporte	


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
' MJP: 18/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  20/03/09: Estandarización de código y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  25/06/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
