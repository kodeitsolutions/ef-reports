Imports System.Data
Partial Class rIntegracion_Accesos
   Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcComandoSelect As String
        
		Try
        lcComandoSelect = "SELECT	Cod_Int, " _
						          & "Nom_Int, " _
								  & "Status, " _
								  & "Case When Status = 'A' Then 'Activo' Else 'Inactivo' End as Status_Integracion_Accesos " _
						& "FROM	 Integracion_Accesos " _
						& "WHERE Cod_Int between '" & cusAplicacion.goReportes.paParametrosIniciales(0) & "'" _
								  & " And '" & cusAplicacion.goReportes.paParametrosFinales(0) & "'" _
         						  & " And Status between '" & cusAplicacion.goReportes.paParametrosIniciales(1) & "'" _
								  & " And '" & cusAplicacion.goReportes.paParametrosFinales(1) & "'" _
						& " ORDER BY Cod_Int, " _
								  & "Nom_Int"

        


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")
          
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rIntegracion_Accesos", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrIntegracion_Accesos.ReportSource = loObjetoReporte

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
' MJP : 10/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP : 11/07/08 : Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP : 14/07/08 : Agregacion filtro Status
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'