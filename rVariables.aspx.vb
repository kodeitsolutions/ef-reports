'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rVariables"
'-------------------------------------------------------------------------------------------'
Partial Class rVariables
   Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


	Try	
	
	        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()
	        
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT  Cod_Var,")
			loConsulta.AppendLine("        Nom_Var,")
			loConsulta.AppendLine("        Status,")
			loConsulta.AppendLine("        (CASE Status ")
			loConsulta.AppendLine("            WHEN 'A' THEN 'Activo'")
			loConsulta.AppendLine("            WHEN 'S' THEN 'Suspendido'")
			loConsulta.AppendLine("            ELSE 'Inactivo'")
			loConsulta.AppendLine("        END)    Estatus, ")
			loConsulta.AppendLine("        Cod_Uni,")
			loConsulta.AppendLine("        Tip_Var,")
			loConsulta.AppendLine("        Val_Max_Esp,")
			loConsulta.AppendLine("        Val_Min_Esp,")
			loConsulta.AppendLine("        Modulo, ")
			loConsulta.AppendLine("        Seccion,")
			loConsulta.AppendLine("        Opcion")
			loConsulta.AppendLine("FROM    Variables")
			loConsulta.AppendLine("WHERE   Cod_Var BETWEEN " & lcParametro0Desde)
			loConsulta.AppendLine("    AND " & lcParametro0Hasta)
			loConsulta.AppendLine("    AND Status IN (" & lcParametro1Desde & ")")
			loConsulta.AppendLine("    AND Tip_Var IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine("ORDER BY      " & lcOrdenamiento)

		
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rVariables", laDatosReporte)
	   
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrVariables.ReportSource = loObjetoReporte
			
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
' RJG: 22/12/14: Codigo inicial.
'-------------------------------------------------------------------------------------------'
