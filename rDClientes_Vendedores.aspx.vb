'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDClientes_Vendedores"
'-------------------------------------------------------------------------------------------'
Partial Class rDClientes_Vendedores
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
		
		Dim lcParametro0Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro0Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
		Dim lcParametro1Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro1Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro2Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
		Dim lcParametro2Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
		Dim lcParametro3Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
		Dim lcParametro3Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
		Dim lcParametro4Desde	As  String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
		Dim lcParametro5Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
		Dim lcParametro5Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
		Dim lcParametro6Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
        Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
        Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

	Try
	
		Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT	Devoluciones_Clientes.Documento, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Fec_Ini, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Fec_Fin, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Cod_Cli, ")
			loComandoSeleccionar.AppendLine("			Clientes.Nom_Cli, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Cod_Ven, ")
			loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Cod_Tra, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Cod_Mon, ")
			loComandoSeleccionar.AppendLine("			Vendedores.Status, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Factura, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Control, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Mon_Net, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes.Mon_Sal  ")
			loComandoSeleccionar.AppendLine(" FROM		Clientes, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Clientes, ")
			loComandoSeleccionar.AppendLine("			Vendedores, ")
			loComandoSeleccionar.AppendLine("			Transportes, ")
			loComandoSeleccionar.AppendLine("			Monedas ")
			loComandoSeleccionar.AppendLine(" WHERE		Devoluciones_Clientes.Cod_Cli		=	Clientes.Cod_Cli ")
			loComandoSeleccionar.AppendLine("			And Devoluciones_Clientes.Cod_Ven	=	Vendedores.Cod_Ven ")
			loComandoSeleccionar.AppendLine("			And Devoluciones_Clientes.Cod_Tra	=	Transportes.Cod_Tra ")
			loComandoSeleccionar.AppendLine("			And Devoluciones_Clientes.Cod_Mon	=	Monedas.Cod_Mon ")
			loComandoSeleccionar.AppendLine("			And Devoluciones_Clientes.Documento	Between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("			And Devoluciones_Clientes.Fec_Ini	Between " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("			And Devoluciones_Clientes.Cod_Cli				Between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("			And Devoluciones_Clientes.Cod_Ven				Between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("			And Devoluciones_Clientes.Status				IN (" & lcParametro4Desde & ")")
			loComandoSeleccionar.AppendLine("			And Devoluciones_Clientes.Cod_Tra				Between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Devoluciones_Clientes.Cod_Mon				Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY   Devoluciones_Clientes.Cod_Ven, " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Devoluciones_Clientes.Documento ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDClientes_Vendedores", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDClientes_Vendedores.ReportSource = loObjetoReporte

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
' JJD: 14/10/08: Continuacion de la programacion del reporte
'-------------------------------------------------------------------------------------------'
' GCR: 23/03/09: Modificaciones al codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  13/07/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'