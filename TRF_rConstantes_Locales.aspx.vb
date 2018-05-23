'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rConstantes_Locales"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rConstantes_Locales
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	    Try	
	
	        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("DECLARE @lcCodCon_Desde AS VARCHAR(10) = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @lcCodCon_Hasta AS VARCHAR(10) = " & lcParametro0Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Cod_Con,")
            loConsulta.AppendLine("		    Nom_Con,")
            loConsulta.AppendLine("		    Val_Num")
            loConsulta.AppendLine("FROM Constantes_Locales")
            loConsulta.AppendLine("WHERE Cod_Con BETWEEN @lcCodCon_Desde AND @lcCodCon_Hasta")
            loConsulta.AppendLine("     AND Tipo = 'N'")
            loConsulta.AppendLine("     AND Status = 'A'")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

           
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rConstantes_Locales", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
			Me.crvTRF_rConstantes_Locales.ReportSource = loObjetoReporte
			  
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
