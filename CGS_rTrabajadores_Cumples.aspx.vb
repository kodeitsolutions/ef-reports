'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rTrabajadores_Cumples"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rTrabajadores_Cumples
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try
            loConsulta.AppendLine("DECLARE @lcCodTra_Desde	AS VARCHAR(10) = ''")
            loConsulta.AppendLine("DECLARE @lcCodTra_Hasta	AS VARCHAR(10) = 'zzzzzz'")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   Trabajadores.Cod_Tra, ")
            loConsulta.AppendLine("		    Trabajadores.Nom_Tra,")
            loConsulta.AppendLine("		    Trabajadores.Cedula,")
            loConsulta.AppendLine("		    Trabajadores.Fec_Nac,")
            loConsulta.AppendLine("         CASE MONTH(Trabajadores.Fec_Nac) ")
            loConsulta.AppendLine("			    WHEN '1' THEN 'Enero' ")
            loConsulta.AppendLine("			    WHEN '2' THEN 'Febrero' ")
            loConsulta.AppendLine("			    WHEN '3' THEN 'Marzo' ")
            loConsulta.AppendLine("			    WHEN '4' THEN 'Abril' ")
            loConsulta.AppendLine("			    WHEN '5' THEN 'Mayo' ")
            loConsulta.AppendLine("			    WHEN '6' THEN 'Junio' ")
            loConsulta.AppendLine("			    WHEN '7' THEN 'Julio' ")
            loConsulta.AppendLine("			    WHEN '8' THEN 'Agosto' ")
            loConsulta.AppendLine("			    WHEN '9' THEN 'Septiembre' ")
            loConsulta.AppendLine("			    WHEN '10' THEN 'Octubre' ")
            loConsulta.AppendLine("			    WHEN '11' THEN 'Noviembre' ")
            loConsulta.AppendLine("			    ELSE 'Diciembre' ")
            loConsulta.AppendLine("		    END							AS Mes,")
            loConsulta.AppendLine("		    MONTH(Trabajadores.Fec_Nac)	AS Mes_Aux,")
            loConsulta.AppendLine("		    (CAST(DATEDIFF(dd,Trabajadores.Fec_Nac,GETDATE()) / 365.25 AS INT)) AS Edad")
            loConsulta.AppendLine("FROM Trabajadores ")
            loConsulta.AppendLine("WHERE Trabajadores.Cod_Tra BETWEEN @lcCodTra_Desde AND @lcCodTra_Hasta")
            loConsulta.AppendLine(" AND Trabajadores.Status IN ( " & lcParametro1Desde & " )")
            loConsulta.AppendLine("ORDER BY MONTH(Trabajadores.Fec_Nac),Trabajadores.Cod_Tra")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rTrabajadores_Cumples", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rTrabajadores_Cumples.ReportSource = loObjetoReporte

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
