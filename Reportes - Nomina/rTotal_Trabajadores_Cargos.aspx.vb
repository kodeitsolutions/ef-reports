'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTotal_Trabajadores_Cargos"
'-------------------------------------------------------------------------------------------'
Partial Class rTotal_Trabajadores_Cargos
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
            loConsulta.AppendLine("SELECT	")
            loConsulta.AppendLine("		SUM(COALESCE(Renglones_Campos_Nomina.Val_Num, 0))/COUNT(Trabajadores.cod_tra)	AS Prom_Sueldo,")
            loConsulta.AppendLine("		COUNT(Trabajadores.cod_tra)*100/CAST(SUM(COUNT(Trabajadores.cod_tra))OVER()  AS Decimal(28,10)) Porcentaje,")
            loConsulta.AppendLine("		(CASE  ")
            loConsulta.AppendLine("				WHEN (SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))/COUNT(Trabajadores.cod_tra)/365) <>0 ")
            loConsulta.AppendLine("				THEN CAST((SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))/COUNT(Trabajadores.cod_tra)/365) AS VARCHAR)+' Años ' ")
            loConsulta.AppendLine("				ELSE ''")
            loConsulta.AppendLine("		END) + ")
            loConsulta.AppendLine("		(CASE WHEN FLOOR((CAST(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE())) AS DECIMAL(28,10))/COUNT(Trabajadores.cod_tra)/365 ")
            loConsulta.AppendLine("					- SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))/COUNT(Trabajadores.cod_tra)/365)*12) = 0")
            loConsulta.AppendLine("				THEN ''")
            loConsulta.AppendLine("			  WHEN FLOOR((CAST(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE())) AS DECIMAL(28,10))/COUNT(Trabajadores.cod_tra)/365 ")
            loConsulta.AppendLine("					- SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))/COUNT(Trabajadores.cod_tra)/365)*12) <10")
            loConsulta.AppendLine("				THEN	' '+CAST(FLOOR((CAST(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE())) AS DECIMAL(28,10))/COUNT(Trabajadores.cod_tra)/365 ")
            loConsulta.AppendLine("					- SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))/COUNT(Trabajadores.cod_tra)/365)*12) AS VARCHAR)+ ' Meses'")
            loConsulta.AppendLine("			  ELSE CAST(FLOOR((CAST(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE())) AS DECIMAL(28,10))/COUNT(Trabajadores.cod_tra)/365 ")
            loConsulta.AppendLine("					- SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))/COUNT(Trabajadores.cod_tra)/365)*12) AS VARCHAR)+ ' Meses'")
            loConsulta.AppendLine("		END)	AS antiguedad,")
            loConsulta.AppendLine("	(CASE	")
            loConsulta.AppendLine("			WHEN SUM(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))) OVER()/ SUM(COUNT(Trabajadores.cod_tra)) OVER() /365 <> 0	")
            loConsulta.AppendLine("			THEN CAST( SUM(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))) OVER()/ SUM(COUNT(Trabajadores.cod_tra)) OVER()/365 AS VARCHAR) + ' Años '	")
            loConsulta.AppendLine("			ELSE ''	")
            loConsulta.AppendLine("		END) +	")
            loConsulta.AppendLine("		(CASE WHEN FLOOR(((CAST(SUM(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))) OVER() AS DECIMAL(28,10)) / SUM(COUNT(Trabajadores.cod_tra)) OVER() /365) - 	")
            loConsulta.AppendLine("					SUM(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))) OVER()/ SUM(COUNT(Trabajadores.cod_tra)) OVER() /365)*12) = 0	")
            loConsulta.AppendLine("				THEN ''	")
            loConsulta.AppendLine("				WHEN FLOOR(((CAST(SUM(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))) OVER() AS DECIMAL(28,10)) / SUM(COUNT(Trabajadores.cod_tra)) OVER() /365) - 	")
            loConsulta.AppendLine("					SUM(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))) OVER()/ SUM(COUNT(Trabajadores.cod_tra)) OVER() /365)*12) <10	")
            loConsulta.AppendLine("				THEN ' '+CAST(	FLOOR(((CAST(SUM(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))) OVER() AS DECIMAL(28,10)) / SUM(COUNT(Trabajadores.cod_tra)) OVER() /365) - 	")
            loConsulta.AppendLine("								SUM(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))) OVER()/ SUM(COUNT(Trabajadores.cod_tra)) OVER() /365)*12) AS VARCHAR) + ' Meses'	")
            loConsulta.AppendLine("				ELSE CAST(	FLOOR(((CAST(SUM(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))) OVER() AS DECIMAL(28,10)) / SUM(COUNT(Trabajadores.cod_tra)) OVER() /365) - 	")
            loConsulta.AppendLine("								SUM(SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE()))) OVER()/ SUM(COUNT(Trabajadores.cod_tra)) OVER() /365)*12) AS VARCHAR) + ' Meses'	")
            loConsulta.AppendLine("		END) AS Prom_Anti,	")
            loConsulta.AppendLine("		Cargos.cod_car,")
            loConsulta.AppendLine("		Cargos.nom_car,")
            loConsulta.AppendLine("		COUNT(Trabajadores.cod_tra) AS Trabajadores,")
            loConsulta.AppendLine("		SUM(SUM(COALESCE(Renglones_Campos_Nomina.Val_Num, 0))) OVER()/SUM(COUNT(Trabajadores.cod_tra)) OVER() Sueldo_total_promedio")
            loConsulta.AppendLine("FROM Trabajadores")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina")
            loConsulta.AppendLine("        ON  Renglones_Campos_Nomina.Cod_tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Renglones_Campos_Nomina.Cod_Cam = 'A001'")
            loConsulta.AppendLine("	JOIN Cargos ON Cargos.cod_car = Trabajadores.cod_car")
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
            loConsulta.AppendLine("Group by Cargos.cod_car,Cargos.nom_car")
            loConsulta.AppendLine("ORDER BY Cargos.cod_car")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTotal_Trabajadores_Cargos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTotal_Trabajadores_Cargos.ReportSource = loObjetoReporte

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
' EAG: 05/09/15: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
