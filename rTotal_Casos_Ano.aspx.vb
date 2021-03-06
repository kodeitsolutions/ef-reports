﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTotal_Casos_Ano"
'-------------------------------------------------------------------------------------------'
Partial Class rTotal_Casos_Ano
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
		   
		    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
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
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loConsulta As New StringBuilder()
			 

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      YEAR(COALESCE(Renglones_Casos.Fec_Ini, Casos.Fec_Ini))                                                   AS Ano,")
            loConsulta.AppendLine("            COUNT(DISTINCT Casos.Cod_Reg)                                                                            AS Clientes,")
            loConsulta.AppendLine("            COUNT(DISTINCT Casos.Cod_Eje)                                                                            AS Ejecutores,")
            loConsulta.AppendLine("            COUNT(DISTINCT CAST(COALESCE(Renglones_Casos.Fec_Ini, Casos.Fec_Ini) AS DATE))                           AS Dias,")
            loConsulta.AppendLine("            COUNT(DISTINCT Casos.Documento)                                                                          AS Casos,")
            loConsulta.AppendLine("            COUNT(*)                                                                                                 AS Actividades,")
            loConsulta.AppendLine("            SUM(CASE  WHEN Renglones_Casos.facturable = 1 ")
            loConsulta.AppendLine("                        THEN Renglones_Casos.duracion")
            loConsulta.AppendLine("                        ELSE 0 END)                                                                                  AS Horas_Fact,")
            loConsulta.AppendLine("            (CASE WHEN COALESCE(SUM(Renglones_Casos.duracion), 0) = 0 ")
            loConsulta.AppendLine("                THEN 0 ")
            loConsulta.AppendLine("                ELSE SUM(CASE  WHEN Renglones_Casos.facturable = 1  ")
            loConsulta.AppendLine("                                THEN Renglones_Casos.duracion ")
            loConsulta.AppendLine("                                ELSE 0 END       ")
            loConsulta.AppendLine("                     )/COALESCE(SUM(Renglones_Casos.duracion), 0)")
            loConsulta.AppendLine("            END)*100.0                                                                                               AS Por_Fact,")
            loConsulta.AppendLine("            SUM(CASE  WHEN Renglones_Casos.facturable = 0")
            loConsulta.AppendLine("                        THEN Renglones_Casos.duracion")
            loConsulta.AppendLine("                        ELSE 0 END)                                                                                  AS Horas_No_Fact,")
            loConsulta.AppendLine("            COALESCE(SUM(Renglones_Casos.duracion), 0)                                                               AS Horas_Totales,")
            loConsulta.AppendLine("            SUM(CASE  WHEN Renglones_Casos.facturable = 1 ")
            loConsulta.AppendLine("                        THEN Renglones_Casos.duracion")
            loConsulta.AppendLine("                        ELSE 0 END)/COUNT(DISTINCT CAST(COALESCE(Renglones_Casos.Fec_Ini, Casos.Fec_Ini) AS DATE))   AS Horas_Fact_Dia")
            loConsulta.AppendLine("FROM        Casos")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Casos ON Renglones_Casos.Documento = Casos.Documento")
            loConsulta.AppendLine("    JOIN    Vendedores AS Ejecutores")
            loConsulta.AppendLine("        ON  Ejecutores.Cod_Ven = COALESCE(Renglones_Casos.Cod_Eje, Casos.Cod_Eje)")
            loConsulta.AppendLine("WHERE      Casos.Documento BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 	    AND COALESCE(Renglones_Casos.Fec_Ini, Casos.Fec_Ini)	BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro1Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Status	IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine(" 	    AND Casos.Cod_Reg	BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Cod_Coo	BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro4Hasta)
            loConsulta.AppendLine(" 	    AND COALESCE(Renglones_Casos.Cod_Eje, Casos.Cod_Eje)   BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro5Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Cod_Suc	BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro6Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Cod_Rev	BETWEEN " & lcParametro7Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro7Hasta)
            loConsulta.AppendLine(" 	    AND Ejecutores.Cod_Tip	BETWEEN " & lcParametro8Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro8Hasta)
            loConsulta.AppendLine("GROUP BY    YEAR(COALESCE(Renglones_Casos.Fec_Ini, Casos.Fec_Ini))")
            loConsulta.AppendLine("ORDER BY    " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTotal_Casos_Ano", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTotal_Casos_Ano.ReportSource = loObjetoReporte


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
' RJG: 07/07/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 08/05/15: Se ajustó el filtro de Revisión y se agregó un filtro de Tipo de Ejeuctor. '
'                Se ajustó el calculo de horas promedio.                                    '
'-------------------------------------------------------------------------------------------'
' RJG: 05/09/15: Se agregó el total de registros.                                           '
'-------------------------------------------------------------------------------------------'
