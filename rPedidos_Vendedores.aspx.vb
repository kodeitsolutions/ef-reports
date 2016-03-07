'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPedidos_Vendedores"
'-------------------------------------------------------------------------------------------'
Partial Class rPedidos_Vendedores
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

			loComandoSeleccionar.AppendLine("SELECT Pedidos.Documento, ")
			loComandoSeleccionar.AppendLine("		Pedidos.Fec_Ini, " )
			loComandoSeleccionar.AppendLine("		Pedidos.Fec_Fin, " )
			loComandoSeleccionar.AppendLine("		Pedidos.Cod_Cli, " )
			loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli, " )
			loComandoSeleccionar.AppendLine("		Pedidos.Cod_Ven, " )
			loComandoSeleccionar.AppendLine("		Vendedores.Nom_Ven, " )
			loComandoSeleccionar.AppendLine("		Pedidos.Cod_Tra, " )
			loComandoSeleccionar.AppendLine("		Pedidos.Cod_Mon, " )
			loComandoSeleccionar.AppendLine("		Vendedores.Status, " )
			loComandoSeleccionar.AppendLine("		Pedidos.Control, " )
			loComandoSeleccionar.AppendLine("		Pedidos.Mon_Net, " )
			loComandoSeleccionar.AppendLine("		Pedidos.Mon_Sal  " )
			loComandoSeleccionar.AppendLine("From	Clientes, " )
			loComandoSeleccionar.AppendLine("		Pedidos, " )
			loComandoSeleccionar.AppendLine("		Vendedores, " )
			loComandoSeleccionar.AppendLine("		Transportes, " )
			loComandoSeleccionar.AppendLine("		Monedas " )
			loComandoSeleccionar.AppendLine("WHERE Pedidos.Cod_Cli = Clientes.Cod_Cli " )
			loComandoSeleccionar.AppendLine("		AND Pedidos.Cod_Ven = Vendedores.Cod_Ven " )
			loComandoSeleccionar.AppendLine("		AND Pedidos.Cod_Tra = Transportes.Cod_Tra " )
			loComandoSeleccionar.AppendLine("		AND Pedidos.Cod_Mon = Monedas.Cod_Mon " )
			loComandoSeleccionar.AppendLine("		AND Pedidos.Documento BETWEEN " & lcParametro0Desde )
			loComandoSeleccionar.AppendLine("		AND " & lcParametro0Hasta )
			loComandoSeleccionar.AppendLine("		AND Pedidos.Fec_Ini BETWEEN " & lcParametro1Desde )
			loComandoSeleccionar.AppendLine("		AND " & lcParametro1Hasta )
			loComandoSeleccionar.AppendLine("		AND Clientes.Cod_Cli BETWEEN " & lcParametro2Desde )
			loComandoSeleccionar.AppendLine("		AND " & lcParametro2Hasta )
			loComandoSeleccionar.AppendLine("		AND Vendedores.Cod_Ven BETWEEN " & lcParametro3Desde )
			loComandoSeleccionar.AppendLine("		AND " & lcParametro3Hasta )
			loComandoSeleccionar.AppendLine("		AND Pedidos.Status IN (" & lcParametro4Desde & ")" )
			loComandoSeleccionar.AppendLine("		AND Transportes.Cod_Tra BETWEEN " & lcParametro5Desde )
			loComandoSeleccionar.AppendLine("		AND " & lcParametro5Hasta )
			loComandoSeleccionar.AppendLine("		AND Monedas.Cod_Mon BETWEEN " & lcParametro6Desde )
            loComandoSeleccionar.AppendLine("		AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("       AND Pedidos.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("       AND Pedidos.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("    	AND " & lcParametro8Hasta)
            'loComandoSeleccionar.AppendLine("ORDER BY  Pedidos.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY    Pedidos.Cod_Ven, " & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPedidos_Vendedores", laDatosReporte)
		
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPedidos_Vendedores.ReportSource = loObjetoReporte


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
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  13/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'