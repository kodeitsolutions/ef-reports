'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "KDE_rCasos_Renglones_MONO"
'-------------------------------------------------------------------------------------------'
Partial Class KDE_rCasos_Renglones_MONO
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
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            'Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("DECLARE @lcDcto_Desde AS VARCHAR(10) = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @lcDcto_Hasta AS VARCHAR(10) = " & lcParametro0Hasta)
            loConsulta.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro1Desde)
            loConsulta.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro1Hasta)
            loConsulta.AppendLine("DECLARE @lcCodCli_Desde AS VARCHAR(12) = " & lcParametro3Desde)
            loConsulta.AppendLine("DECLARE @lcCodCli_Hasta AS VARCHAR(12) = " & lcParametro3Hasta)
            loConsulta.AppendLine("DECLARE @lcCodCoo_Desde AS VARCHAR(12) = " & lcParametro4Desde)
            loConsulta.AppendLine("DECLARE @lcCodCoo_Hasta AS VARCHAR(12) = " & lcParametro4Hasta)
            loConsulta.AppendLine("DECLARE @lcCodEje_Desde AS VARCHAR(12) = " & lcParametro5Desde)
            loConsulta.AppendLine("DECLARE @lcCodEje_Hasta AS VARCHAR(12) = " & lcParametro5Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   Clientes.Cod_Cli                                AS Cod_Reg, ")
            loConsulta.AppendLine("         Clientes.Nom_Cli                                AS Nom_Reg, ")
            loConsulta.AppendLine("         Casos.Documento                                 AS Documento, ")
            loConsulta.AppendLine("         Casos.Status                                    AS Status, ")
            loConsulta.AppendLine("         Casos.Fec_Ini                                   AS Fec_Ini, ")
            loConsulta.AppendLine("         Casos.Fec_Fin                                   AS Fec_Fin, ")
            loConsulta.AppendLine("         Casos.Asunto                                    AS Asunto, ")
            loConsulta.AppendLine("         Casos.Cod_Coo                                   AS Cod_Coo, ")
            loConsulta.AppendLine("         Casos.Cod_Coo                                   AS Nom_Coo,")
            loConsulta.AppendLine("         Renglones_Casos.Cod_Eje                         AS Nom_Eje,")
            loConsulta.AppendLine("         CASE WHEN Casos.Departamento = 'Con/Par/Pro'")
            loConsulta.AppendLine("              THEN 'Consultoría/Parametrización/Programación'")
            loConsulta.AppendLine("              ELSE Casos.Departamento")
            loConsulta.AppendLine("         END                                             AS Departamento, ")
            loConsulta.AppendLine("         SUBSTRING(Casos.Comentario,CHARINDEX(':',Casos.Comentario,0)+1,LEN(RTRIM(Casos.Comentario)))    AS Comentario, ")
            loConsulta.AppendLine("         Renglones_Casos.Fec_Ini                         AS Renglon_Fec_Ini,")
            loConsulta.AppendLine("         Renglones_Casos.Hor_Ini                         AS Renglon_Hor_Ini,")
            loConsulta.AppendLine("         Renglones_Casos.Hor_Fin                         AS Renglon_Hor_Fin,")
            loConsulta.AppendLine("         COALESCE(Renglones_Casos.Duracion,0)            AS Renglon_Duracion,")
            loConsulta.AppendLine("         COALESCE(Renglones_Casos.Actividad,'')          AS Renglon_Actividad,")
            loConsulta.AppendLine("		    CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
            loConsulta.AppendLine("         CONCAT('Cliente: ' , COALESCE((SELECT RTRIM(Nom_Cli) FROM Clientes WHERE Cod_Cli = @lcCodCli_Desde),'')")
            loConsulta.AppendLine("         ,', Coordinador: ' , COALESCE((SELECT RTRIM(Nom_Ven) FROM Vendedores WHERE Cod_Cli = @lcCodCoo_Desde),'')")
            loConsulta.AppendLine("         ,', Ejecutor: ' , COALESCE((SELECT RTRIM(Nom_Ven) FROM Vendedores WHERE Cod_Cli = @lcCodEje_Desde),'')")
            loConsulta.AppendLine("         ,', Departamento: ')		AS Parametros,")
            If lcParametro6Desde.Trim() = "'Con/Par/Pro','Formación','Help-Desk'" Then
                loConsulta.AppendLine("     'Consultoría/Parametrización/Programación, Formación, Help-Desk.'   AS Par_Departamento ")
            ElseIf lcParametro6Desde.Trim() = "'Con/Par/Pro'" Then
                loConsulta.AppendLine("     'Consultoría/Parametrización/Programación.'   AS Par_Departamento ")
            Else
                loConsulta.AppendLine("      CONCAT(" & lcParametro6Desde & ",'.')          AS Par_Departamento")
            End If
            loConsulta.AppendLine("         ")
            loConsulta.AppendLine("FROM Casos")
            loConsulta.AppendLine("    JOIN Clientes ON Clientes.Cod_Cli = Casos.Cod_Reg")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Casos ON Renglones_Casos.Documento = Casos.Documento ")
            loConsulta.AppendLine("WHERE Casos.Documento BETWEEN @lcDcto_Desde AND @lcDcto_Hasta")
            loConsulta.AppendLine(" 	AND Renglones_Casos.Fec_Ini	BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loConsulta.AppendLine(" 	AND Casos.Status	IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine(" 	AND Casos.Cod_Reg	BETWEEN @lcCodCli_Desde AND @lcCodCli_Hasta")
            loConsulta.AppendLine(" 	AND Casos.Cod_Coo	BETWEEN @lcCodCoo_Desde AND @lcCodCoo_Hasta")
            loConsulta.AppendLine(" 	AND Casos.Cod_Eje   BETWEEN @lcCodEje_Desde AND @lcCodEje_Hasta")
            loConsulta.AppendLine(" 	AND Casos.Departamento	IN (" & lcParametro6Desde & ")")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento & ", COALESCE(Renglones_Casos.Renglon,0) ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("KDE_rCasos_Renglones_MONO", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvKDE_rCasos_Renglones_MONO.ReportSource = loObjetoReporte


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
