'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rResumenProspectos_Paises"
'-------------------------------------------------------------------------------------------'
Partial Class rResumenProspectos_Paises 
    Inherits vis2Formularios.frmReporte
    
	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
		    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
		    Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT      Paises.Cod_Pai,")
			loConsulta.AppendLine("            Paises.Nom_Pai,")
			loConsulta.AppendLine("            Paises.Telefonos,")
			loConsulta.AppendLine("            Monedas.Cod_Mon,")
			loConsulta.AppendLine("            Monedas.Nom_Mon,")
			loConsulta.AppendLine("            COUNT(*) Cantidad ")
			loConsulta.AppendLine("FROM        Prospectos")
			loConsulta.AppendLine("    JOIN    Paises ON Paises.Cod_Pai = Prospectos.Cod_Pai")
			loConsulta.AppendLine("    JOIN    Monedas ON Monedas.Cod_Mon = Paises.Cod_Mon")
			loConsulta.AppendLine("WHERE	   Prospectos.Cod_Pro BETWEEN " & lcParametro0Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro0Hasta)
			loConsulta.AppendLine(" 	    AND Prospectos.status IN (" & lcParametro1Desde &  ")")
			loConsulta.AppendLine(" 	    AND Prospectos.Cod_Ven BETWEEN " & lcParametro2Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro2Hasta)
			loConsulta.AppendLine(" 	    AND Prospectos.Cod_Zon BETWEEN " & lcParametro3Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro3Hasta)
			loConsulta.AppendLine(" 	    AND Paises.Cod_Pai BETWEEN " & lcParametro4Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro4Hasta)
			loConsulta.AppendLine(" 	    AND Prospectos.Cod_Suc BETWEEN " & lcParametro5Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro5Hasta)
			loConsulta.AppendLine(" 	    AND Prospectos.Cod_Cla BETWEEN " & lcParametro6Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro6Hasta)
			loConsulta.AppendLine(" 	    AND Prospectos.Cod_Tip BETWEEN " & lcParametro7Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro7Hasta)
			loConsulta.AppendLine("GROUP BY    Paises.Cod_Pai, ")
			loConsulta.AppendLine("            Paises.Nom_Pai,")
			loConsulta.AppendLine("            Paises.Telefonos,")
			loConsulta.AppendLine("            Monedas.Cod_Mon,")
			loConsulta.AppendLine("            Monedas.Nom_Mon")
			loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
			loConsulta.AppendLine("")
     
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rResumenProspectos_Paises", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrResumenProspectos_Paises.ReportSource = loObjetoReporte
            

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
' RJG: 07/05/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
