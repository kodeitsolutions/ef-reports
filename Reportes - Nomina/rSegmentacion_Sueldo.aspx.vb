'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rSegmentacion_Sueldo"
'-------------------------------------------------------------------------------------------'
Partial Class rSegmentacion_Sueldo
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
            loConsulta.AppendLine("SELECT  Tabla.Nivel,   ")
            loConsulta.AppendLine("		COUNT(Nivel) AS Trabajadores,   ")
            loConsulta.AppendLine("		CAST(SUM(dias) AS DECIMAL(28,10))/COUNT(Nivel)/365  AS Antiguedad,   ")
            loConsulta.AppendLine("		SUM(Sueldo)/COUNT(Nivel)							AS Promedio_Sueldo,")
            loConsulta.AppendLine("		CAST(COUNT(Nivel) AS DECIMAL(28,10))*100/Tabla.Tra AS Porcentaje,   ")
            loConsulta.AppendLine("		(CASE Nivel   ")
            loConsulta.AppendLine("			WHEN 'Bajo' THEN Minimo   ")
            loConsulta.AppendLine("			WHEN 'Medio' THEN Mitad_inferior + 0.01   ")
            loConsulta.AppendLine("			WHEN 'Alto' THEN Mitad_superior +0.01   ")
            loConsulta.AppendLine("		END) AS Monto_Inicial,   ")
            loConsulta.AppendLine("		(CASE Nivel   ")
            loConsulta.AppendLine("			WHEN 'Bajo' THEN Mitad_inferior   ")
            loConsulta.AppendLine("			WHEN 'Medio' THEN Mitad_superior    ")
            loConsulta.AppendLine("			WHEN 'Alto' THEN maximo   ")
            loConsulta.AppendLine("		END) AS Monto_final   ")
            loConsulta.AppendLine("FROM   ")
            loConsulta.AppendLine("(SELECT   ")
            loConsulta.AppendLine("        Trabajadores.Fec_Ini														AS Fec_Ini, ")
            loConsulta.AppendLine("		COUNT(Trabajadores.Fec_Ini)	OVER()												AS Tra, ")
            loConsulta.AppendLine("		DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE())								AS dias, ")
            loConsulta.AppendLine("		COALESCE(Renglones_Campos_Nomina.Val_Num, 0)								AS Sueldo,")
            loConsulta.AppendLine("		MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0))							OVER() minimo, ")
            loConsulta.AppendLine("		(	MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER() -  ")
            loConsulta.AppendLine("			MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER()  ")
            loConsulta.AppendLine("		)  / 3 + MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0))					OVER() mitad_inferior, ")
            loConsulta.AppendLine("		(	MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER() -  ")
            loConsulta.AppendLine("			MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER()  ")
            loConsulta.AppendLine("		)  / 3 * 2 + MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0))				OVER() mitad_superior, ")
            loConsulta.AppendLine("		MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0))							OVER() maximo, ")
            loConsulta.AppendLine("		(CASE  ")
            loConsulta.AppendLine("			WHEN COALESCE(Renglones_Campos_Nomina.Val_Num, 0) <=  ")
            loConsulta.AppendLine("				(	MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER() -  ")
            loConsulta.AppendLine("					MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER()  ")
            loConsulta.AppendLine("				)  / 3 + MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER() ")
            loConsulta.AppendLine("			THEN 'Bajo' ")
            loConsulta.AppendLine("			WHEN COALESCE(Renglones_Campos_Nomina.Val_Num, 0) >  ")
            loConsulta.AppendLine("				(	MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER() -  ")
            loConsulta.AppendLine("					MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER()  ")
            loConsulta.AppendLine("				)  / 3 + MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER() ")
            loConsulta.AppendLine("				AND COALESCE(Renglones_Campos_Nomina.Val_Num, 0) <= ")
            loConsulta.AppendLine("				(	MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER() -  ")
            loConsulta.AppendLine("					MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER()  ")
            loConsulta.AppendLine("				)  / 3 * 2 + MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER() ")
            loConsulta.AppendLine("			THEN 'Medio' ")
            loConsulta.AppendLine("			WHEN COALESCE(Renglones_Campos_Nomina.Val_Num, 0) > ")
            loConsulta.AppendLine("				(	MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER() -  ")
            loConsulta.AppendLine("					MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER()  ")
            loConsulta.AppendLine("				)  / 3 * 2 + MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER() ")
            loConsulta.AppendLine("			THEN 'Alto' ")
            loConsulta.AppendLine("		END)																			AS Nivel ")
            loConsulta.AppendLine("FROM	   Trabajadores ")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina")
            loConsulta.AppendLine("        ON  Renglones_Campos_Nomina.Cod_tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Renglones_Campos_Nomina.Cod_Cam = 'A001'")
            loConsulta.AppendLine("    LEFT JOIN Cargos")
            loConsulta.AppendLine("        ON  Cargos.Cod_car = Trabajadores.Cod_car")
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
            loConsulta.AppendLine(") AS Tabla")
            loConsulta.AppendLine("Group by nivel, Minimo,Mitad_inferior,Mitad_Superior,maximo,tabla.tra")
            loConsulta.AppendLine("ORDER BY Monto_Inicial ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.toString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rSegmentacion_Sueldo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrSegmentacion_Sueldo.ReportSource = loObjetoReporte

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
' EAG: 03/09/15: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
