Imports System.Data
Partial Class rClientes_cVendedores

     Inherits vis2Formularios.frmReporte
     
     Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcComandoSelect As String

        lcComandoSelect = "SELECT	Clientes.Cod_Cli, " _
         & "Clientes.Nom_Cli, " _
         & "Vendedores.Nom_Ven, " _
         & "Clientes.status " _
         & "FROM	 Clientes, Vendedores " _
         & "WHERE Clientes.Cod_Ven = Vendedores.Cod_Ven " _
         & " And Clientes.Cod_Cli between '" & cusAplicacion.goReportes.paParametrosIniciales(0) & "'" _
         & " And '" & cusAplicacion.goReportes.paParametrosFinales(0) & "'" _
         & " And Vendedores.Cod_Ven between '" & cusAplicacion.goReportes.paParametrosIniciales(1) & "'" _
         & " And '" & cusAplicacion.goReportes.paParametrosFinales(1) & "'" _
         & " And Clientes.Status between '" & cusAplicacion.goReportes.paParametrosIniciales(2) & "'" _
         & " And '" & cusAplicacion.goReportes.paParametrosFinales(2) & "'" _
         & " ORDER BY Cod_Cli"

        Try


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rClientes_cVendedores", laDatosReporte)
			
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
			Me.crvrClientes_cVendedores.ReportSource = loObjetoReporte


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