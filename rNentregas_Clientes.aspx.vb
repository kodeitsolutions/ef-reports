Imports System.Data
Partial Class rNentregas_Clientes
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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
        

			 loComandoSeleccionar.Append("  SELECT		Clientes.Cod_Cli, " )
			 loComandoSeleccionar.Append(" 				Clientes.Nom_Cli, " )
			 loComandoSeleccionar.Append(" 				Clientes.Status, " )
			 loComandoSeleccionar.Append(" 				Entregas.Documento, " )
			 loComandoSeleccionar.Append(" 				Entregas.Fec_ini, " )
			 loComandoSeleccionar.Append(" 				Entregas.Fec_fin, " )
			 loComandoSeleccionar.Append(" 				Entregas.Cod_Ven, " )
			 loComandoSeleccionar.Append(" 				Entregas.Cod_Tra, " )
			 loComandoSeleccionar.Append(" 				Entregas.Cod_Mon, " )
			 loComandoSeleccionar.Append(" 				Entregas.Control, " )
			 loComandoSeleccionar.Append(" 				Entregas.Comentario, " )
			 loComandoSeleccionar.Append(" 				Entregas.Mon_Net, " )
			 loComandoSeleccionar.Append(" 				Entregas.Mon_Sal  " )
			 loComandoSeleccionar.Append(" FROM			Clientes, " )
			 loComandoSeleccionar.Append(" 				Entregas, " )
			 loComandoSeleccionar.Append(" 				Vendedores, " )
			 loComandoSeleccionar.Append(" 				Transportes, " )
			 loComandoSeleccionar.Append(" 				Monedas " )
			 loComandoSeleccionar.Append(" WHERE		Clientes.Cod_Cli = Entregas.Cod_Cli " )
			 loComandoSeleccionar.Append(" 				AND Entregas.Cod_Ven = Vendedores.Cod_Ven " )
			 loComandoSeleccionar.Append(" 				AND Entregas.Cod_Tra = Transportes.Cod_Tra " )
			 loComandoSeleccionar.Append(" 				AND Entregas.Cod_Mon = Monedas.Cod_Mon " )
			 loComandoSeleccionar.Append(" 				AND Entregas.Documento between " & lcParametro0Desde)
			 loComandoSeleccionar.Append(" 				AND " & lcParametro0Hasta) 
			 loComandoSeleccionar.Append(" 				AND Entregas.Fec_Ini between " & lcParametro1Desde) 
			 loComandoSeleccionar.Append(" 				AND " & lcParametro1Hasta)  
			 loComandoSeleccionar.Append(" 				AND Entregas.Cod_Cli between " & lcParametro2Desde) 
			 loComandoSeleccionar.Append(" 				AND " & lcParametro2Hasta) 
			 loComandoSeleccionar.Append(" 				AND Entregas.Cod_Ven between " & lcParametro3Desde) 
			 loComandoSeleccionar.Append(" 				AND " & lcParametro3Hasta) 
			 loComandoSeleccionar.Append(" 				AND Entregas.Status IN (" & lcParametro4Desde & ")")
			 loComandoSeleccionar.Append(" 				AND Entregas.Cod_Tra between " & lcParametro5Desde) 
			 loComandoSeleccionar.Append(" 				AND " & lcParametro5Hasta) 
			 loComandoSeleccionar.Append(" 				AND Entregas.Cod_Mon between " & lcParametro6Desde)
            loComandoSeleccionar.Append(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Entregas.Cod_Suc between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            'loComandoSeleccionar.Append(" ORDER BY		Entregas.Cod_Cli, " )
            'loComandoSeleccionar.Append(" 				Entregas.Fec_Ini, " )
            'loComandoSeleccionar.Append(" 				Entregas.Fec_Fin, " )
            'loComandoSeleccionar.Append(" 				Entregas.Cod_Ven ")
            loComandoSeleccionar.AppendLine("ORDER BY   Entregas.Cod_Cli, " & lcOrdenamiento)
 

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rNentregas_Clientes", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrNentregas_Clientes.ReportSource = loObjetoReporte


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
' YYG: 08/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 11/08/08: Entonacion del código 
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  10/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'