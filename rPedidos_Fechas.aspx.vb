'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPedidos_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rPedidos_Fechas
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.Appendline("SELECT Pedidos.Documento, " )
			loComandoSeleccionar.Appendline("		Pedidos.Fec_Ini, " )
			loComandoSeleccionar.Appendline("		Pedidos.Fec_Fin, " )
			loComandoSeleccionar.Appendline("		Pedidos.Cod_Cli, " )
			loComandoSeleccionar.Appendline("		Clientes.Nom_Cli, " )
			loComandoSeleccionar.Appendline("		Pedidos.Cod_Ven, " )
			loComandoSeleccionar.Appendline("		Pedidos.Cod_Tra, " )
			loComandoSeleccionar.Appendline("		Pedidos.Cod_Mon, " )
			loComandoSeleccionar.Appendline("		Pedidos.Control, " )
			loComandoSeleccionar.Appendline("		Pedidos.Mon_Net, " )
			loComandoSeleccionar.Appendline("		Pedidos.Mon_Sal,  " )
			loComandoSeleccionar.Appendline("		Pedidos.mon_bru  " )
			loComandoSeleccionar.Appendline("From Clientes, " )
			loComandoSeleccionar.Appendline("		Pedidos, " )
			loComandoSeleccionar.Appendline("		Vendedores, " )
			loComandoSeleccionar.Appendline("		Transportes, " )
			loComandoSeleccionar.Appendline("		Monedas " )
			loComandoSeleccionar.Appendline("WHERE Pedidos.Cod_Cli = Clientes.Cod_Cli " )
			loComandoSeleccionar.Appendline("		AND Pedidos.Cod_Ven = Vendedores.Cod_Ven " )
			loComandoSeleccionar.Appendline("		AND Pedidos.Cod_Tra = Transportes.Cod_Tra " )
			loComandoSeleccionar.Appendline("		AND Pedidos.Cod_Mon = Monedas.Cod_Mon " )
			loComandoSeleccionar.Appendline("		AND Pedidos.Documento BETWEEN " & lcParametro0Desde )
			loComandoSeleccionar.Appendline("		AND " & lcParametro0Hasta )
			loComandoSeleccionar.Appendline("		AND Pedidos.Fec_Ini BETWEEN " & lcParametro1Desde )
			loComandoSeleccionar.Appendline("		AND " & lcParametro1Hasta )
			loComandoSeleccionar.Appendline("		AND Clientes.Cod_Cli BETWEEN " & lcParametro2Desde )
			loComandoSeleccionar.Appendline("		AND " & lcParametro2Hasta )
			loComandoSeleccionar.Appendline("		AND Vendedores.Cod_Ven BETWEEN " & lcParametro3Desde )
			loComandoSeleccionar.Appendline("		AND " & lcParametro3Hasta )
			loComandoSeleccionar.Appendline("		AND Pedidos.Status IN ( " & lcParametro4Desde & ")" )
			loComandoSeleccionar.Appendline("		AND Transportes.Cod_Tra BETWEEN " & lcParametro5Desde )
			loComandoSeleccionar.Appendline("		AND " & lcParametro5Hasta )
			loComandoSeleccionar.Appendline("		AND Monedas.Cod_Mon BETWEEN " & lcParametro6Desde )
            loComandoSeleccionar.AppendLine("		AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("       AND Pedidos.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("       AND Pedidos.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY   CONVERT(nchar(30), Pedidos.Fec_Ini,112), " & lcOrdenamiento)
            'loComandoSeleccionar.Appendline("ORDER BY  Pedidos.Documento "	)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString , "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPedidos_Fechas", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPedidos_Fechas.ReportSource = loObjetoReporte


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
' MJP   :  18/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' CMS:  16/04/09: Cambios Estandarización de codigo, correcion campo estatus. 
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
