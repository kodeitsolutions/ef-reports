Imports System.Data
Partial Class rEmpresas
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcComandoSelect As String

	Try	
		lcComandoSelect = "SELECT	Cod_Emp, " _
						         & "Nom_Emp, " _
						         & "Status, " _
						         & "(Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Empresas " _
				         & "FROM	 Empresas " _
				         & " ORDER BY Cod_Emp, Nom_Emp"
						'        & "WHERE Cod_Emp between '" & cusAplicacion.goReportes.paParametrosIniciales(0) & "'" _
						'        & " And '" & cusAplicacion.goReportes.paParametrosFinales(0) & "'" _
  						'        & " And Status between '" & cusAplicacion.goReportes.paParametrosIniciales(1) & "'" _
						'	     & " And '" & cusAplicacion.goReportes.paParametrosFinales(1) & "'" _
			    '        & " ORDER BY Cod_Emp, Nom_Emp"

        

		
            Dim loServicios As New cusDatos.goDatos
            
            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEmpresas", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrEmpresas.ReportSource = loObjetoReporte
			
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
' MJP   :  15/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  05/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
