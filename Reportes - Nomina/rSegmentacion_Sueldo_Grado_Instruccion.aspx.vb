'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rSegmentacion_Sueldo_Grado_Instruccion"
'-------------------------------------------------------------------------------------------'
Partial Class rSegmentacion_Sueldo_Grado_Instruccion
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
            loConsulta.AppendLine("SELECT	nom_gra, ")
            loConsulta.AppendLine("		cod_gra, ")
            loConsulta.AppendLine("		nivel, ")
            loConsulta.AppendLine("		Tra, ")
            loConsulta.AppendLine("		COUNT(NIVEL) AS Trabajadores, ")
            loConsulta.AppendLine("		(CASE ")
            loConsulta.AppendLine("				WHEN SUM(dias)/COUNT(NIVEL)/365 <> 0  ")
            loConsulta.AppendLine("				THEN CAST( SUM(dias)/COUNT(NIVEL)/365 AS VARCHAR)+ ' Años ' ")
            loConsulta.AppendLine("				ELSE '' ")
            loConsulta.AppendLine("		END) + ")
            loConsulta.AppendLine("		(CASE ")
            loConsulta.AppendLine("				WHEN FLOOR((CAST(SUM(dias)AS DECIMAL(28,10))/COUNT(NIVEL)/365 - SUM(dias)/COUNT(NIVEL)/365)*12) ")
            loConsulta.AppendLine("						= 0 ")
            loConsulta.AppendLine("				THEN '' ")
            loConsulta.AppendLine("				WHEN FLOOR((CAST(SUM(dias)AS DECIMAL(28,10))/COUNT(NIVEL)/365 - SUM(dias)/COUNT(NIVEL)/365)*12) < 10 ")
            loConsulta.AppendLine("				THEN ' ' + CAST(FLOOR((CAST(SUM(dias)AS DECIMAL(28,10))/COUNT(NIVEL)/365 - SUM(dias)/COUNT(NIVEL)/365)*12) AS VARCHAR) + ' Meses' ")
            loConsulta.AppendLine("				ELSE CAST(FLOOR((CAST(SUM(dias)AS DECIMAL(28,10))/COUNT(NIVEL)/365 - SUM(dias)/COUNT(NIVEL)/365)*12) AS VARCHAR) + ' Meses' ")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("		END)	AS Antiguedad, ")
            loConsulta.AppendLine("		SUM(Sueldo)/COUNT(Nivel)							AS Promedio_Sueldo, ")
            loConsulta.AppendLine("		CAST(COUNT(Nivel) AS DECIMAL(28,10))*100/Tabla.Tra AS Porcentaje, ")
            loConsulta.AppendLine("		(CASE Nivel    ")
            loConsulta.AppendLine("			WHEN 'Bajo' THEN Minimo    ")
            loConsulta.AppendLine("			WHEN 'Medio' THEN Mitad_inferior + 0.01    ")
            loConsulta.AppendLine("			WHEN 'Alto' THEN Mitad_superior +0.01    ")
            loConsulta.AppendLine("		END) AS Monto_Inicial,    ")
            loConsulta.AppendLine("		(CASE Nivel    ")
            loConsulta.AppendLine("			WHEN 'Bajo' THEN Mitad_inferior    ")
            loConsulta.AppendLine("			WHEN 'Medio' THEN Mitad_superior     ")
            loConsulta.AppendLine("			WHEN 'Alto' THEN maximo    ")
            loConsulta.AppendLine("		END) AS Monto_final,   ")
            loConsulta.AppendLine("		promedio_grupo, ")
            loConsulta.AppendLine("		SUM(SUM(promedio_grupo*tra)) OVER() / SUM(COUNT(NIVEL)) OVER() AS promedio_global, ")
            loConsulta.AppendLine("		(CASE  ")
            loConsulta.AppendLine("				WHEN Dias_grupo/Tra/365 <> 0   ")
            loConsulta.AppendLine("				THEN CAST( Dias_grupo/Tra/365 AS VARCHAR)+ ' Años '  ")
            loConsulta.AppendLine("				ELSE ''  ")
            loConsulta.AppendLine("		END) +  ")
            loConsulta.AppendLine("		(CASE  ")
            loConsulta.AppendLine("				WHEN FLOOR((CAST(Dias_grupo AS DECIMAL(28,10))/Tra/365 - Dias_grupo/Tra/365)*12)  ")
            loConsulta.AppendLine("						= 0  ")
            loConsulta.AppendLine("				THEN ''  ")
            loConsulta.AppendLine("				WHEN FLOOR((CAST(Dias_grupo AS DECIMAL(28,10))/Tra/365 - Dias_grupo/Tra/365)*12) < 10  ")
            loConsulta.AppendLine("				THEN ' ' + CAST(FLOOR((CAST(Dias_grupo AS DECIMAL(28,10))/Tra/365 - Dias_grupo/Tra/365)*12) AS VARCHAR) + ' Meses'  ")
            loConsulta.AppendLine("				ELSE CAST(FLOOR((CAST(Dias_grupo AS DECIMAL(28,10))/Tra/365 - Dias_grupo/Tra/365)*12) AS VARCHAR) + ' Meses'  ")
            loConsulta.AppendLine("		END)	AS Antiguedad_grupal, ")
            loConsulta.AppendLine("		(CASE ")
            loConsulta.AppendLine("				WHEN SUM(SUM(dias)) OVER() /SUM(COUNT(NIVEL)) OVER()/365 <> 0   ")
            loConsulta.AppendLine("				THEN CAST( SUM(SUM(dias)) OVER() /SUM(COUNT(NIVEL)) OVER()/365 AS VARCHAR)+ ' Años '  ")
            loConsulta.AppendLine("				ELSE ''  ")
            loConsulta.AppendLine("		END) +  ")
            loConsulta.AppendLine("		(CASE  ")
            loConsulta.AppendLine("				WHEN FLOOR((CAST(SUM(SUM(dias)) OVER()  AS DECIMAL(28,10))/SUM(COUNT(NIVEL)) OVER()/365 - SUM(SUM(dias)) OVER() /SUM(COUNT(NIVEL)) OVER()/365)*12)  ")
            loConsulta.AppendLine("						= 0  ")
            loConsulta.AppendLine("				THEN ''  ")
            loConsulta.AppendLine("				WHEN FLOOR((CAST(SUM(SUM(dias)) OVER()  AS DECIMAL(28,10))/SUM(COUNT(NIVEL)) OVER()/365 - SUM(SUM(dias)) OVER() /SUM(COUNT(NIVEL)) OVER()/365)*12) < 10  ")
            loConsulta.AppendLine("				THEN ' ' + CAST(FLOOR((CAST(SUM(SUM(dias)) OVER()  AS DECIMAL(28,10))/SUM(COUNT(NIVEL)) OVER()/365 - SUM(SUM(dias)) OVER() /SUM(COUNT(NIVEL)) OVER()/365)*12) AS VARCHAR) + ' Meses'  ")
            loConsulta.AppendLine("				ELSE CAST(FLOOR((CAST(SUM(SUM(dias)) OVER()  AS DECIMAL(28,10))/SUM(COUNT(NIVEL)) OVER()/365 - SUM(SUM(dias)) OVER() /SUM(COUNT(NIVEL)) OVER()/365)*12) AS VARCHAR) + ' Meses'  ")
            loConsulta.AppendLine("		END) AS Antiguedad_Total ")
            loConsulta.AppendLine("FROM ")
            loConsulta.AppendLine("(SELECT   ")
            loConsulta.AppendLine("        Grados_Instruccion.nom_gra,    ")
            loConsulta.AppendLine("        Grados_Instruccion.cod_gra,    ")
            loConsulta.AppendLine("		COUNT(Trabajadores.Fec_Ini) OVER(PARTITION BY Grados_Instruccion.nom_gra)	 Tra,    ")
            loConsulta.AppendLine("		DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE())								AS dias,    ")
            loConsulta.AppendLine("		COALESCE(Renglones_Campos_Nomina.Val_Num, 0)								AS Sueldo,   ")
            loConsulta.AppendLine("		MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0))OVER(PARTITION BY Grados_Instruccion.nom_gra) minimo,   ")
            loConsulta.AppendLine("		(	MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra) -     ")
            loConsulta.AppendLine("			MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra)     ")
            loConsulta.AppendLine("		)  / 3 + MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0))OVER(PARTITION BY Grados_Instruccion.nom_gra) mitad_inferior,    ")
            loConsulta.AppendLine("		(	MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra) -     ")
            loConsulta.AppendLine("			MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra)     ")
            loConsulta.AppendLine("		)  / 3 * 2 + MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0))				OVER(PARTITION BY Grados_Instruccion.nom_gra) mitad_superior,    ")
            loConsulta.AppendLine("		MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0))							OVER(PARTITION BY Grados_Instruccion.nom_gra) maximo,   ")
            loConsulta.AppendLine("				(CASE     ")
            loConsulta.AppendLine("			WHEN COALESCE(Renglones_Campos_Nomina.Val_Num, 0) <=     ")
            loConsulta.AppendLine("				(	MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra) -     ")
            loConsulta.AppendLine("					MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra)     ")
            loConsulta.AppendLine("				)  / 3 + MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra)    ")
            loConsulta.AppendLine("			THEN 'Bajo'    ")
            loConsulta.AppendLine("			WHEN COALESCE(Renglones_Campos_Nomina.Val_Num, 0) >     ")
            loConsulta.AppendLine("				(	MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra) -     ")
            loConsulta.AppendLine("					MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra)     ")
            loConsulta.AppendLine("				)  / 3 + MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra)    ")
            loConsulta.AppendLine("				AND COALESCE(Renglones_Campos_Nomina.Val_Num, 0) <=    ")
            loConsulta.AppendLine("				(	MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra) -     ")
            loConsulta.AppendLine("					MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra)     ")
            loConsulta.AppendLine("				)  / 3 * 2 + MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra)    ")
            loConsulta.AppendLine("			THEN 'Medio'    ")
            loConsulta.AppendLine("			WHEN COALESCE(Renglones_Campos_Nomina.Val_Num, 0) >    ")
            loConsulta.AppendLine("				(	MAX(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra) -     ")
            loConsulta.AppendLine("					MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra)     ")
            loConsulta.AppendLine("				)  / 3 * 2 + MIN(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(PARTITION BY Grados_Instruccion.nom_gra)    ")
            loConsulta.AppendLine("			THEN 'Alto'    ")
            loConsulta.AppendLine("		END)																			AS Nivel,   ")
            loConsulta.AppendLine("		SUM(COALESCE(Renglones_Campos_Nomina.Val_Num, 0)) OVER(	PARTITION BY Grados_Instruccion.nom_gra) / COUNT(Trabajadores.Fec_Ini) OVER(PARTITION BY Grados_Instruccion.nom_gra) promedio_grupo,")
            loConsulta.AppendLine("		SUM(DATEDIFF(day,Trabajadores.Fec_Ini,GETDATE())) OVER(	PARTITION BY Grados_Instruccion.nom_gra) Dias_grupo	   ")
            loConsulta.AppendLine("FROM	   Trabajadores ")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina")
            loConsulta.AppendLine("        ON  Renglones_Campos_Nomina.Cod_tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Renglones_Campos_Nomina.Cod_Cam = 'A001'")
            loConsulta.AppendLine("    LEFT JOIN Cargos")
            loConsulta.AppendLine("        ON  Cargos.Cod_car = Trabajadores.Cod_car")
            loConsulta.AppendLine("	    JOIN Grados_Instruccion ON Grados_Instruccion.cod_gra = Trabajadores.cod_gra ")
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
            loConsulta.AppendLine("Group by nom_gra,cod_gra,nivel,tra,minimo,mitad_inferior, mitad_superior,minimo, maximo,promedio_grupo,dias_grupo")
            loConsulta.AppendLine("ORDER BY Tabla.Nom_gra,Monto_Inicial ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rSegmentacion_Sueldo_Grado_Instruccion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrSegmentacion_Sueldo_Grado_Instruccion.ReportSource = loObjetoReporte

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
' EAG: 07/09/15: Se particionó la información por grado de instrucción.                                                            '
'-------------------------------------------------------------------------------------------'
