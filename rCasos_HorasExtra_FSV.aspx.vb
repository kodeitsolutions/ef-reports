'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCasos_HorasExtra_FSV"
'-------------------------------------------------------------------------------------------'
Partial Class rCasos_HorasExtra_FSV
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
		   
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loConsulta As New StringBuilder()
			 

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  RC.Cod_Eje AS Cod_Eje,")
            loConsulta.AppendLine("        Vendedores.Nom_Ven AS Nom_Eje,")
            loConsulta.AppendLine("        CAST(RC.fec_ini AS DATE) AS Fecha,")
            loConsulta.AppendLine("        (CASE dbo.udf_GetISOWeekDay(CAST(RC.fec_ini AS DATE))")
            loConsulta.AppendLine("            WHEN 1 THEN 'Lunes'")
            loConsulta.AppendLine("            WHEN 2 THEN 'Martes'")
            loConsulta.AppendLine("            WHEN 3 THEN 'Miércoles'")
            loConsulta.AppendLine("            WHEN 4 THEN 'Jueves'")
            loConsulta.AppendLine("            WHEN 5 THEN 'Viernes'")
            loConsulta.AppendLine("            WHEN 6 THEN 'Sábado'")
            loConsulta.AppendLine("            WHEN 7 THEN 'Domingo'")
            loConsulta.AppendLine("            ELSE '[N/A]'")
            loConsulta.AppendLine("        END) AS Dia,")
            loConsulta.AppendLine("        SUM(RC.Duracion) As Duracion, ")
            loConsulta.AppendLine("        (CASE WHEN dbo.udf_GetISOWeekDay(CAST(RC.fec_ini AS DATE))<6 ")
            loConsulta.AppendLine("        THEN")
            loConsulta.AppendLine("            (CASE WHEN SUM(RC.Duracion) > 8")
            loConsulta.AppendLine("                THEN SUM(RC.Duracion) - 8")
            loConsulta.AppendLine("                ELSE 0")
            loConsulta.AppendLine("            END)")
            loConsulta.AppendLine("        ELSE")
            loConsulta.AppendLine("            SUM(RC.Duracion)")
            loConsulta.AppendLine("        END)    AS Extra")
            loConsulta.AppendLine("FROM    Renglones_Casos RC")
            loConsulta.AppendLine("    JOIN Casos")
            loConsulta.AppendLine("        ON Casos.Documento = RC.Documento")
            loConsulta.AppendLine(" 	   AND Casos.Status <> 'Anulado'")
            loConsulta.AppendLine("    JOIN vendedores ON Vendedores.Cod_Ven = RC.Cod_Eje")
            loConsulta.AppendLine("WHERE   RC.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 	   AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 	   AND RC.Cod_Eje  BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine(" 	   AND " & lcParametro1Hasta)
            loConsulta.AppendLine("GROUP BY CAST(RC.fec_ini AS DATE), RC.Cod_Eje, Vendedores.Nom_Ven")
            loConsulta.AppendLine("HAVING SUM(RC.Duracion) > 8 Or dbo.udf_GetISOWeekDay(CAST(RC.fec_ini AS DATE)) > 5")
            loConsulta.AppendLine("ORDER BY    " & lcOrdenamiento & ",")
            loConsulta.AppendLine("            CAST(RC.Fec_Ini AS DATE)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCasos_HorasExtra_FSV", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCasos_HorasExtra_FSV.ReportSource = loObjetoReporte


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
' Fin del codigo.                                                                           '
'-------------------------------------------------------------------------------------------'
' RJG: 15/08/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
