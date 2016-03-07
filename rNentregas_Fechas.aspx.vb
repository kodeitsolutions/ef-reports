'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rNentregas_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rNentregas_Fechas
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

			loComandoSeleccionar.Appendline("SELECT Entregas.Documento, " )
			loComandoSeleccionar.Appendline("		Entregas.Fec_Ini, " )
			loComandoSeleccionar.Appendline("		Entregas.Fec_Fin, " )
			loComandoSeleccionar.Appendline("		Entregas.Cod_Cli, " )
			loComandoSeleccionar.Appendline("		Clientes.Nom_Cli, " )
			loComandoSeleccionar.Appendline("		Entregas.Cod_Ven, " )
			loComandoSeleccionar.Appendline("		Entregas.Cod_Tra, " )
			loComandoSeleccionar.Appendline("		Entregas.Cod_Mon, " )
			loComandoSeleccionar.Appendline("		Entregas.Control, " )
			loComandoSeleccionar.Appendline("		Entregas.Mon_Net, " )
			loComandoSeleccionar.Appendline("		Entregas.Mon_Sal  " )
			loComandoSeleccionar.Appendline("From Clientes, " )
			loComandoSeleccionar.Appendline("		Entregas, " )
			loComandoSeleccionar.Appendline("		Vendedores, " )
			loComandoSeleccionar.Appendline("		Transportes, " )
			loComandoSeleccionar.Appendline("		Monedas " )
			loComandoSeleccionar.Appendline("WHERE Entregas.Cod_Cli = Clientes.Cod_Cli " )
			loComandoSeleccionar.Appendline("		And Entregas.Cod_Ven = Vendedores.Cod_Ven " )
			loComandoSeleccionar.Appendline("		And Entregas.Cod_Tra = Transportes.Cod_Tra " )
			loComandoSeleccionar.Appendline("		And Entregas.Cod_Mon = Monedas.Cod_Mon " )
			loComandoSeleccionar.Appendline("		And Entregas.Documento between " & lcParametro0Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro1Hasta )
			loComandoSeleccionar.Appendline("		And Entregas.Fec_Ini between " & lcParametro1Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro1Hasta  )
			loComandoSeleccionar.Appendline("		And Clientes.Cod_Cli between " & lcParametro2Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro2Hasta )
			loComandoSeleccionar.Appendline("		And Vendedores.Cod_Ven between " & lcParametro3Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro3Hasta )
			loComandoSeleccionar.Appendline("		And Entregas.Status IN ( " & lcParametro4Desde & " ) ")
			loComandoSeleccionar.Appendline("		And Transportes.Cod_Tra between " & lcParametro5Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro5Hasta )
			loComandoSeleccionar.Appendline("		And Monedas.Cod_Mon between " & lcParametro6Desde )
            loComandoSeleccionar.AppendLine("		And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("       AND Entregas.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("       AND Entregas.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro8Hasta)
            'loComandoSeleccionar.Appendline("ORDER BY  Entregas.Fec_Ini, " )
            'loComandoSeleccionar.AppendLine("       Entregas.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY   Entregas.Fec_Ini DESC, " & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rNentregas_Fechas", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrNentregas_Fechas.ReportSource = loObjetoReporte


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
' YYG:  08/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 11/08/08: Entonacion del código 
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  10/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'