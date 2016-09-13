'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTrabajadores_Cumpleaños_Ampliado"
'-------------------------------------------------------------------------------------------'
Partial Class rTrabajadores_Cumpleaños_Ampliado
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

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

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            loConsulta.AppendLine("SELECT	cod_tra, ")
            loConsulta.AppendLine("		nom_tra, ")
            loConsulta.AppendLine("                Case Month(fec_nac) ")
            loConsulta.AppendLine("			WHEN '1' THEN 'Enero' ")
            loConsulta.AppendLine("			WHEN '2' THEN 'Febrero' ")
            loConsulta.AppendLine("			WHEN '3' THEN 'Marzo' ")
            loConsulta.AppendLine("			WHEN '4' THEN 'Abril' ")
            loConsulta.AppendLine("			WHEN '5' THEN 'Mayo' ")
            loConsulta.AppendLine("			WHEN '6' THEN 'Junio' ")
            loConsulta.AppendLine("			WHEN '7' THEN 'Julio' ")
            loConsulta.AppendLine("			WHEN '8' THEN 'Agosto' ")
            loConsulta.AppendLine("			WHEN '9' THEN 'Septiembre' ")
            loConsulta.AppendLine("			WHEN '10' THEN 'Octubre' ")
            loConsulta.AppendLine("			WHEN '11' THEN 'Noviembre' ")
            loConsulta.AppendLine("			ELSE 'Diciembre' ")
            loConsulta.AppendLine("		END							AS Mes, ")
            loConsulta.AppendLine("		DAY(fec_nac)				AS dia,")
            loConsulta.AppendLine("		CASE status ")
            loConsulta.AppendLine("			WHEN 'A' THEN 'ACTIVO' ")
            loConsulta.AppendLine("			WHEN 'I' THEN 'INACTIVO' ")
            loConsulta.AppendLine("			WHEN 'S' THEN 'SUSPENIDO' ")
            loConsulta.AppendLine("		END							AS status, ")
            loConsulta.AppendLine("		MONTH(FEC_NAC)						AS MES_AUX, ")
            loConsulta.AppendLine("		CAST(YEAR(GETDATE())-YEAR(fec_nac) AS VARCHAR)+' Años' AS años_cumplir, ")
            loConsulta.AppendLine("		CAST(SUM(YEAR(GETDATE())-YEAR(fec_nac)) OVER(PARTITION BY MONTH(fec_nac)) / COUNT(COD_TRA) OVER(PARTITION BY MONTH(fec_nac)) AS VARCHAR)+' Años' AS edad_promedio, ")
            loConsulta.AppendLine("		(CASE   ")
            loConsulta.AppendLine("				WHEN (DATEDIFF(day,fec_ini,DATEADD(year,YEAR(GETDATE())-YEAR(fec_nac),fec_nac))/365) <>0  ")
            loConsulta.AppendLine("				THEN CAST((DATEDIFF(day,fec_ini,DATEADD(year,YEAR(GETDATE())-YEAR(fec_nac),fec_nac))/365) AS VARCHAR)+' Años '  ")
            loConsulta.AppendLine("				ELSE '' ")
            loConsulta.AppendLine("		END) + ")
            loConsulta.AppendLine("		CASE WHEN FLOOR((CAST(DATEDIFF(day,fec_ini,DATEADD(year,YEAR(GETDATE())-YEAR(fec_nac),fec_nac)) AS DECIMAL(28,10))/365 ")
            loConsulta.AppendLine("					- DATEDIFF(day,fec_ini,DATEADD(year,YEAR(GETDATE())-YEAR(fec_nac),fec_nac))/365)*12) = 0 ")
            loConsulta.AppendLine("				THEN '' ")
            loConsulta.AppendLine("			WHEN FLOOR((CAST(DATEDIFF(day,fec_ini,DATEADD(year,YEAR(GETDATE())-YEAR(fec_nac),fec_nac)) AS DECIMAL(28,10))/365 ")
            loConsulta.AppendLine("					- DATEDIFF(day,fec_ini,DATEADD(year,YEAR(GETDATE())-YEAR(fec_nac),fec_nac))/365)*12) <10 ")
            loConsulta.AppendLine("			THEN ' '+ CAST( FLOOR((CAST(DATEDIFF(day,fec_ini,DATEADD(year,YEAR(GETDATE())-YEAR(fec_nac),fec_nac)) AS DECIMAL(28,10))/365 ")
            loConsulta.AppendLine("					- DATEDIFF(day,fec_ini,DATEADD(year,YEAR(GETDATE())-YEAR(fec_nac),fec_nac))/365)*12)  AS VARCHAR) +' Meses' ")
            loConsulta.AppendLine("			ELSE CAST( FLOOR((CAST(DATEDIFF(day,fec_ini,DATEADD(year,YEAR(GETDATE())-YEAR(fec_nac),fec_nac)) AS DECIMAL(28,10))/365 ")
            loConsulta.AppendLine("					- DATEDIFF(day,fec_ini,DATEADD(year,YEAR(GETDATE())-YEAR(fec_nac),fec_nac))/365)*12)  AS VARCHAR) +' Meses' ")
            loConsulta.AppendLine("		END AS antiguedad ")
            loConsulta.AppendLine("FROM trabajadores ")
            loConsulta.AppendLine("WHERE   Trabajadores.Tip_Tra = 'Trabajador'")
            loConsulta.AppendLine("        AND Trabajadores.Cod_Tra BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("        AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Status IN ( " & lcParametro1Desde & " )")
            loConsulta.AppendLine("        AND Trabajadores.Cod_Con BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("        AND " & lcParametro2Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Dep BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("        AND " & lcParametro3Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Car BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("        AND " & lcParametro4Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Suc BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("        AND " & lcParametro5Hasta)
            loConsulta.AppendLine("ORDER BY      MONTH(fec_nac)," & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.toString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTrabajadores_Cumpleaños_Ampliado", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTrabajadores_Cumpleaños_Ampliado.ReportSource = loObjetoReporte

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
' EAG: 08/09/15: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
