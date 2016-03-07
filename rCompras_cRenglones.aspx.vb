Imports System.Data
Partial Class rCompras_cRenglones
   Inherits vis2Formularios.frmReporte
   
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcComandoSelect As String

        Try

            lcComandoSelect = "SELECT	Compras.Documento, " _
              & "Compras.Cod_Pro, " _
              & "Renglones_Compras.Renglon, " _
              & "Renglones_Compras.Mon_Bru, " _
              & "Renglones_Compras.Mon_Imp1, " _
              & "Renglones_Compras.Mon_Net, " _
              & "Renglones_Compras.Mon_Sal, " _
              & "Proveedores.Nom_Pro, " _
              & "Compras.Fec_Ini, " _
              & "Renglones_Compras.Cod_Art, " _
              & "Articulos.Nom_Art, " _
              & "Renglones_Compras.Can_Art1 " _
              & "FROM	Compras, Renglones_Compras, Proveedores, Articulos " _
              & "WHERE Compras.Documento = Renglones_Compras.Documento " _
              & " AND Compras.Cod_Pro = Proveedores.Cod_Pro " _
              & " AND Renglones_Compras.Cod_Art = Articulos.Cod_Art " _
              & " AND Compras.Documento Between '" & cusAplicacion.goReportes.paParametrosIniciales(0) & "'" _
              & " AND '" & cusAplicacion.goReportes.paParametrosFinales(0) & "'" _
              & " ORDER BY Compras.Documento "

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")
			
			loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rCompras_cRenglones", laDatosReporte)
			
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
	
            Me.crvrCompras_cRenglones.ReportSource = loObjetoReporte

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