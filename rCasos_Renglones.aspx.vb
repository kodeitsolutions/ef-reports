'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCasos_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rCasos_Renglones
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loConsulta As New StringBuilder()
			 

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  Clientes.Cod_Cli                                AS Cod_Reg, ")
            loConsulta.AppendLine("        Clientes.Nom_Cli                                AS Nom_Reg, ")
            loConsulta.AppendLine("        Casos.Documento                                 AS Documento, ")
            loConsulta.AppendLine("        Casos.Status                                    AS Status, ")
            loConsulta.AppendLine("        Casos.Fec_Ini                                   AS Fec_Ini, ")
            loConsulta.AppendLine("        Casos.Fec_Fin                                   AS Fec_Fin, ")
            loConsulta.AppendLine("        (CASE Casos.Status ")
            loConsulta.AppendLine("            WHEN 'Pendiente'  ")
            loConsulta.AppendLine("            THEN DATEDIFF(DAY,Casos.Fec_ini, GETDATE()) ")
            loConsulta.AppendLine("            ELSE DATEDIFF(DAY,Casos.Fec_ini, Casos.Fec_fin) ")
            loConsulta.AppendLine("        END)                                            AS Dias, ")
            loConsulta.AppendLine("        Casos.Asunto                                    AS Asunto, ")
            loConsulta.AppendLine("        Casos.Prioridad                                 AS Prioridad, ")
            loConsulta.AppendLine("        Casos.Clase                                     AS Clase, ")
            loConsulta.AppendLine("        Casos.Grupo                                     AS Grupo, ")
            loConsulta.AppendLine("        Casos.Interno                                   AS Interno, ")
            loConsulta.AppendLine("        Casos.Cod_Coo                                   AS Cod_Coo, ")
            loConsulta.AppendLine("        (CASE WHEN Casos.Cod_Coo<>''")
            loConsulta.AppendLine("        THEN(RTRIM(Casos.Cod_Coo) + ' - ' ")
            loConsulta.AppendLine("            + COALESCE(Coord.Nom_Ven,'[No Válido]') ) ")
            loConsulta.AppendLine("        ELSE '[No Asignado]'")
            loConsulta.AppendLine("        END)                                            AS Nom_Coo,")
            loConsulta.AppendLine("        Casos.Cod_Eje                                   AS Cod_Eje, ")
            loConsulta.AppendLine("        (CASE WHEN Casos.Cod_Eje<>''")
            loConsulta.AppendLine("        THEN(RTRIM(Casos.Cod_Eje) + ' - ' ")
            loConsulta.AppendLine("            + COALESCE(Ejec.Nom_Ven,'[No Válido]') ) ")
            loConsulta.AppendLine("        ELSE '[No Asignado]'")
            loConsulta.AppendLine("        END)                                            AS Nom_Eje, ")
            loConsulta.AppendLine("        Casos.Etapa                                     AS Etapa, ")
            loConsulta.AppendLine("        Casos.Por_Ava                                   AS Por_Ava, ")
            loConsulta.AppendLine("        Casos.Departamento                              AS Departamento, ")
            loConsulta.AppendLine("        Casos.Origen                                    AS Origen, ")
            loConsulta.AppendLine("        Casos.Principal                                 AS Principal,")
            loConsulta.AppendLine("        COALESCE(Renglones_Casos.Renglon,0)             AS Renglon,")
            loConsulta.AppendLine("        Renglones_Casos.Fec_Ini                         AS Renglon_Fec_Ini,")
            loConsulta.AppendLine("        Renglones_Casos.Hor_Ini                         AS Renglon_Hor_Ini,")
            loConsulta.AppendLine("        Renglones_Casos.Hor_Fin                         AS Renglon_Hor_Fin,")
            loConsulta.AppendLine("        COALESCE(Renglones_Casos.Duracion,0)            AS Renglon_Duracion,")
            loConsulta.AppendLine("        COALESCE(Renglones_Casos.Actividad,'')          AS Renglon_Actividad,")
            loConsulta.AppendLine("        COALESCE(Renglones_Casos.Control,'')            AS Renglon_Control,")
            loConsulta.AppendLine("        COALESCE(Renglones_Casos.Cod_Eje,'')            AS Renglon_Cod_Eje,")
            loConsulta.AppendLine("        CAST(COALESCE(Renglones_Casos.Facturable,0)AS BIT) AS Renglon_Facturable")
            loConsulta.AppendLine("FROM    Casos")
            loConsulta.AppendLine("    JOIN Clientes ON Clientes.Cod_Cli = Casos.Cod_Reg")
            loConsulta.AppendLine("    LEFT JOIN Vendedores Coord ON Coord.Cod_Ven = Casos.Cod_Coo")
            loConsulta.AppendLine("    LEFT JOIN Vendedores Ejec ON Ejec.Cod_Ven = Casos.Cod_Eje")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Casos ON Renglones_Casos.Documento = Casos.Documento ")
            loConsulta.AppendLine("WHERE      Casos.Documento BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Fec_Ini	BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro1Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Status	IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine(" 	    AND Casos.Cod_Reg	BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Cod_Coo	BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro4Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Cod_Eje   BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro5Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Cod_Suc	BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro6Hasta)
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento & ", COALESCE(Renglones_Casos.Renglon,0) ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCasos_Renglones", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCasos_Renglones.ReportSource = loObjetoReporte


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
' RJG: 26/06/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 19/06/15: Se corrigió el total acumulado por documento.                              '
'-------------------------------------------------------------------------------------------'
