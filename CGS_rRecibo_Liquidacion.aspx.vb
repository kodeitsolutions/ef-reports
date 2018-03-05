'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rRecibo_Liquidacion"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rRecibo_Liquidacion
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))

            Dim lcCodigoCampoSueldo As String = goServicios.mObtenerCampoFormatoSQL(CStr(goOpciones.mObtener("CAMSUEMEN", "C")))

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("DECLARE @lcCod_Tra AS VARCHAR(10) = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @lnAnios AS INT = " & lcParametro1Desde)
            loConsulta.AppendLine("DECLARE @lnMeses AS INT = " & lcParametro2Desde)
            loConsulta.AppendLine("DECLARE @lnDias AS INT = " & lcParametro3Desde)
            loConsulta.AppendLine("DECLARE @lnCero AS DECIMAL(28,10) = 0")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Recibos.Documento							            AS Documento,		")
            loConsulta.AppendLine("			Recibos.Fecha								            AS Fecha,			")
            loConsulta.AppendLine("			Recibos.Cod_Tra								            AS Cod_Tra,			")
            loConsulta.AppendLine("			Trabajadores.Nom_Tra						            AS Nom_Tra,			")
            loConsulta.AppendLine("			Trabajadores.Cedula							            AS Cedula,			")
            loConsulta.AppendLine("			COALESCE(Renglones_Campos_Nomina.Val_Num,@lnCero)       AS Sueldo_Mensual,")
            loConsulta.AppendLine("			Trabajadores.Fec_Ini						            AS Ingreso,			")
            loConsulta.AppendLine("			Cargos.Nom_Car								            AS Nom_Car,			")
            loConsulta.AppendLine("			Recibos.Fecha								            AS Fecha,			")
            loConsulta.AppendLine("		    @lnAnios                                                AS Año,")
            loConsulta.AppendLine("		    @lnMeses                                                AS Mes,")
            loConsulta.AppendLine("         @lnDias                                                 AS Dia,")
            loConsulta.AppendLine("			Recibos.Fec_Fin								            AS Fec_Fin,			")
            loConsulta.AppendLine("			DATEADD(DAY, 1, Recibos.Fec_Fin)			            AS Fec_Rei,			")
            loConsulta.AppendLine("			Recibos.Comentario							            AS Comentario,		")
            loConsulta.AppendLine("			Recibos.Mon_Net								            AS Mon_Net,			")
            loConsulta.AppendLine("			Renglones_Recibos.Cod_Con					            AS Cod_Con,			")
            loConsulta.AppendLine("			Renglones_Recibos.Nom_con					            AS Nom_con,			")
            loConsulta.AppendLine("			CASE When Renglones_Recibos.Tipo = 'Asignacion'")
            loConsulta.AppendLine("				  THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				  ELSE @lnCero")
            loConsulta.AppendLine("			END													    AS Mon_Asi,")
            loConsulta.AppendLine("			CASE Renglones_Recibos.Tipo ")
            loConsulta.AppendLine("				WHEN 'Retencion' THEN -Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				WHEN 'Deduccion' THEN -Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				ELSE @lnCero")
            loConsulta.AppendLine("			END													    AS Mon_Ded,")
            loConsulta.AppendLine("			Renglones_Recibos.Val_Car								AS Val_Car,")
            loConsulta.AppendLine("			COALESCE(Liquidaciones.Tipo, '')				        AS Tipo_Liquidacion")
            loConsulta.AppendLine("FROM	Recibos")
            loConsulta.AppendLine("     JOIN Renglones_Recibos ON Renglones_Recibos.documento = Recibos.Documento")
            loConsulta.AppendLine("	    JOIN Trabajadores ON Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loConsulta.AppendLine("	    JOIN Cargos ON Cargos.Cod_Car = Trabajadores.Cod_Car")
            loConsulta.AppendLine("	    LEFT JOIN Liquidaciones ON Liquidaciones.Documento = Renglones_Recibos.Doc_Ori")
            loConsulta.AppendLine("		    AND Renglones_Recibos.Tip_Ori = 'Liquidaciones'")
            loConsulta.AppendLine("     LEFT JOIN Renglones_Campos_Nomina ON Renglones_Campos_Nomina.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("		    AND	Renglones_Campos_Nomina.Cod_Cam = " & lcCodigoCampoSueldo)
            loConsulta.AppendLine("WHERE Recibos.Cod_Tra = @lcCod_Tra")
            loConsulta.AppendLine("     AND Renglones_Recibos.Tipo <> 'Otro' ")
            loConsulta.AppendLine("     AND Recibos.Cod_Con = '92' ")
            loConsulta.AppendLine("ORDER BY	(CASE Renglones_Recibos.Tipo ")
            loConsulta.AppendLine("				WHEN 'Asignacion' THEN 0")
            loConsulta.AppendLine("				WHEN 'Retencion' THEN 2")
            loConsulta.AppendLine("				ELSE 3")
            loConsulta.AppendLine("			END) ASC, ")
            loConsulta.AppendLine("			Renglones_Recibos.renglon")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rRecibo_Liquidacion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rRecibo_Liquidacion.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
