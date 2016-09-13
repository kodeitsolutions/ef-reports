'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rComprobantes_fPeriodo"
'-------------------------------------------------------------------------------------------'
Partial Class rComprobantes_fPeriodo
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Try

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()


			loConsulta.AppendLine("DECLARE @lcParametro1Desde VARCHAR(100)	")
			loConsulta.AppendLine("DECLARE @lcParametro1Hasta VARCHAR(100)	")
			loConsulta.AppendLine("DECLARE @lcParametro2Desde DATETIME	")
			loConsulta.AppendLine("DECLARE @lcParametro2Hasta DATETIME	")
			loConsulta.AppendLine("DECLARE @lcParametro3Desde VARCHAR(100)	")
			loConsulta.AppendLine("DECLARE @lcParametro3Hasta VARCHAR(100)	")
			loConsulta.AppendLine("DECLARE @lcParametro4Desde VARCHAR(100)	")
			loConsulta.AppendLine("DECLARE @lcParametro4Hasta VARCHAR(100)	")
			loConsulta.AppendLine("DECLARE @lcParametro5Desde VARCHAR(100)	")
			loConsulta.AppendLine("DECLARE @lcParametro5Hasta VARCHAR(100)	")
			loConsulta.AppendLine("DECLARE @lcParametro6Hasta VARCHAR(100)	")
			loConsulta.AppendLine("DECLARE @lcParametro6Desde VARCHAR(100)	")
			loConsulta.AppendLine("DECLARE @lcParametro7Desde VARCHAR(100)	")
			loConsulta.AppendLine("DECLARE @lcParametro7Hasta VARCHAR(100)	")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SET		@lcParametro1Desde = " & lcParametro0Desde)
			loConsulta.AppendLine("SET		@lcParametro1Hasta = " & lcParametro0Hasta)
			loConsulta.AppendLine("SET		@lcParametro2Desde = " & lcParametro1Desde)
			loConsulta.AppendLine("SET		@lcParametro2Hasta = " & lcParametro1Hasta)
			loConsulta.AppendLine("SET		@lcParametro3Desde = " & lcParametro2Desde)
			loConsulta.AppendLine("SET		@lcParametro3Hasta = " & lcParametro2Hasta)
			loConsulta.AppendLine("SET		@lcParametro4Desde = " & lcParametro3Desde)
			loConsulta.AppendLine("SET		@lcParametro4Hasta = " & lcParametro3Hasta)
			loConsulta.AppendLine("SET		@lcParametro5Desde = " & lcParametro4Desde)
			loConsulta.AppendLine("SET		@lcParametro5Hasta = " & lcParametro4Hasta)
			loConsulta.AppendLine("SET		@lcParametro6Desde = " & lcParametro5Desde)
			loConsulta.AppendLine("SET		@lcParametro6Hasta = " & lcParametro5Hasta)
			loConsulta.AppendLine("SET		@lcParametro7Desde = " & lcParametro6Desde)
			loConsulta.AppendLine("SET		@lcParametro7Hasta = " & lcParametro6Hasta)
			loConsulta.AppendLine("")
			loConsulta.AppendLine("DECLARE @lnCero DECIMAL(28, 10)")
			loConsulta.AppendLine("SET	@lnCero = 0")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT  RTRIM(Renglones_Comprobantes.Documento)    AS Documento,")
			loConsulta.AppendLine("        RTRIM(Renglones_Comprobantes.Adicional)    AS Adicional,")
			loConsulta.AppendLine("        Renglones_Comprobantes.Renglon             AS Renglon,")
			loConsulta.AppendLine("        Renglones_Comprobantes.Cod_Cue             AS Cod_Cue,")
			loConsulta.AppendLine("        Comprobantes.fec_ini                       AS Fec_Doc,")
			loConsulta.AppendLine("        Renglones_Comprobantes.fec_ini             AS fec_ini,")
			loConsulta.AppendLine("        Renglones_Comprobantes.fec_ori             AS fec_ori,")
			loConsulta.AppendLine("        Renglones_Comprobantes.mon_deb             AS mon_deb,")
			loConsulta.AppendLine("        Renglones_Comprobantes.Mon_hab             AS Mon_hab")
			loConsulta.AppendLine("FROM    Renglones_Comprobantes")
			loConsulta.AppendLine("    JOIN comprobantes")
			loConsulta.AppendLine("        ON comprobantes.documento = Renglones_Comprobantes.documento")
			loConsulta.AppendLine("        AND comprobantes.Adicional = Renglones_Comprobantes.Adicional")
			loConsulta.AppendLine("        AND comprobantes.status <> 'Anulado'")
			loConsulta.AppendLine("        AND Tipo = 'Diario'")
			loConsulta.AppendLine("WHERE  (       (CONVERT(VARCHAR(6), Renglones_Comprobantes.fec_ori, 112)<> CONVERT(VARCHAR(6), Renglones_Comprobantes.fec_ini, 112))")
			loConsulta.AppendLine("            OR (CONVERT(VARCHAR(8), Renglones_Comprobantes.fec_ini, 112) <> CONVERT(VARCHAR(8), Comprobantes.fec_ini, 112) )         ")
			loConsulta.AppendLine("        )")
            loConsulta.AppendLine("		AND	Renglones_Comprobantes.Documento BETWEEN @lcParametro1Desde AND	@lcParametro1Hasta")
            loConsulta.AppendLine("		AND	Renglones_Comprobantes.Fec_Ini BETWEEN @lcParametro2Desde AND	@lcParametro2Hasta")
            loConsulta.AppendLine("		AND	Renglones_Comprobantes.Cod_Cue BETWEEN @lcParametro3Desde AND	@lcParametro3Hasta")
            loConsulta.AppendLine("		AND	Renglones_Comprobantes.Cod_Cen BETWEEN @lcParametro4Desde AND	@lcParametro4Hasta")
            loConsulta.AppendLine("		AND	Renglones_Comprobantes.Cod_Gas BETWEEN @lcParametro5Desde AND	@lcParametro5Hasta")
            loConsulta.AppendLine("		AND	Renglones_Comprobantes.Cod_Aux BETWEEN @lcParametro6Desde AND	@lcParametro6Hasta")
            loConsulta.AppendLine("		AND	Renglones_Comprobantes.Cod_Mon BETWEEN @lcParametro7Desde AND	@lcParametro7Hasta")
            loConsulta.AppendLine("ORDER BY	" & lcOrdenamiento & ", (CASE WHEN Renglones_Comprobantes.Mon_Deb>0 THEN 0 ELSE 1 END) ASC, Renglon ASC ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")


			'Me.mEscribirConsulta(loConsulta.ToString())
 
	        Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

        '--------------------------------------------------'
        ' Carga la imagen del logo en cusReportes          '
        '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

        '-------------------------------------------------------------------------------------------------------
        ' Verificando si el select (tabla nº0) trae registros
        '-------------------------------------------------------------------------------------------------------
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rComprobantes_fPeriodo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrComprobantes_fPeriodo.ReportSource = loObjetoReporte

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
' RJG: 16/12/13: Codigo inicial
'-------------------------------------------------------------------------------------------'
