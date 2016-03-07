Imports System.Data
Partial Class rCotizaciones_cVendedor
    Inherits vis2Formularios.frmReporte
    
     Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcComandoSelect As String

        lcComandoSelect = "SELECT	Cotizaciones.Documento, " _
          & "Cotizaciones.Cod_Ven, " _
          & "Cotizaciones.Cod_Cli, " _
          & "Vendedores.Nom_Ven, " _
          & "Clientes.Nom_cli, " _
          & "Cotizaciones.Fec_Ini, " _
          & "Cotizaciones.Fec_Fin, " _
          & "Cotizaciones.Mon_Net, " _
          & "Cotizaciones.Status " _
          & "FROM	Cotizaciones, Clientes, Vendedores   " _
          & "WHERE Cotizaciones.Cod_Ven = Vendedores.Cod_Ven " _
          & " And Cotizaciones.Cod_Cli = Clientes.Cod_Cli  " _
          & " And Cotizaciones.Cod_Ven between '" & cusAplicacion.goReportes.paParametrosIniciales(0) & "'" _
          & " And '" & cusAplicacion.goReportes.paParametrosFinales(0) & "'" _
          & " And Cotizaciones.Status between '" & cusAplicacion.goReportes.paParametrosIniciales(1) & "'" _
          & " And '" & cusAplicacion.goReportes.paParametrosFinales(1) & "'" _
          & " ORDER BY Cotizaciones.Documento "
        

        Try


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")
			
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCotizaciones_cVendedor", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCotizaciones_cVendedor.ReportSource = loObjetoReporte


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
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'