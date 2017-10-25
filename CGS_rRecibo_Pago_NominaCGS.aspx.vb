'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rRecibo_Pago_NominaCGS"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rRecibo_Pago_NominaCGS
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))

            Dim lcParametro7Desde As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(7)).Trim().ToUpper

            Dim lcCodigoCampoSueldo As String
            lcCodigoCampoSueldo = goServicios.mObtenerCampoFormatoSQL(CStr(goOpciones.mObtener("CAMSUEMEN", "C")))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("DECLARE @lcDocumento_Desde AS VARCHAR(10) = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @lcDocumento_Hasta AS VARCHAR(10) = " & lcParametro0Hasta)
            loConsulta.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro1Desde)
            loConsulta.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro1Hasta)
            loConsulta.AppendLine("DECLARE @lcCodTra_Desde AS VARCHAR(10) = " & lcParametro2Desde)
            loConsulta.AppendLine("DECLARE @lcCodTra_Hasta AS VARCHAR(10) = " & lcParametro2Hasta)
            loConsulta.AppendLine("DECLARE @lcCodCon_Desde AS VARCHAR(2) = " & lcParametro3Desde)
            loConsulta.AppendLine("DECLARE @lcCodCon_Hasta AS VARCHAR(2) = " & lcParametro3Hasta)
            loConsulta.AppendLine("DECLARE @lcCodDep_Desde AS VARCHAR(30) = " & lcParametro4Desde)
            loConsulta.AppendLine("DECLARE @lcCodDep_Hasta AS VARCHAR(30) = " & lcParametro4Hasta)
            loConsulta.AppendLine("DECLARE @lcCodSuc_Desde AS VARCHAR(15) = " & lcParametro5Desde)
            loConsulta.AppendLine("DECLARE @lcCodSuc_Hasta AS VARCHAR(15) = " & lcParametro5Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Recibos.Documento							            AS Documento,		")
            loConsulta.AppendLine("			Contratos.Tipo								            AS Tipo_Contrato,	")
            loConsulta.AppendLine("			Recibos.Cod_Tra								            AS Cod_Tra,			")
            loConsulta.AppendLine("			Trabajadores.Nom_Tra						            AS Nom_Tra,			")
            loConsulta.AppendLine("			Trabajadores.Fec_Ini						            AS Fecha_Ingreso,	")
            loConsulta.AppendLine("			Trabajadores.Cedula							            AS Cedula,			")
            loConsulta.AppendLine("			Trabajadores.Tip_Pag						            AS Tip_Pag,			")
            loConsulta.AppendLine("			Trabajadores.Num_Cue						            AS Num_Cue,			")
            loConsulta.AppendLine("			Trabajadores.Cod_Ban						            AS Cod_Ban,			")
            loConsulta.AppendLine("			Departamentos_Nomina.Nom_Dep				            AS Nom_Dep,			")
            loConsulta.AppendLine("			Cargos.Nom_Car						                    AS Nom_Car,			")
            loConsulta.AppendLine("			CAST(COALESCE(Renglones_Campos_Nomina.val_num			        ")
            loConsulta.AppendLine("			         *(CASE WHEN Contratos.Tipo = 'Semanal'			        ")
            loConsulta.AppendLine("			           THEN 1.0/30.0 ELSE 1.0 END),0)")
            loConsulta.AppendLine("				AS DECIMAL(28,4))				                    AS Sueldo_Mensual,	")
            loConsulta.AppendLine("			Recibos.Fec_Ini								            AS Fec_Ini,			")
            loConsulta.AppendLine("			Recibos.Fec_Fin								            AS Fec_Fin,			")
            loConsulta.AppendLine("			Recibos.Comentario							            AS Comentario,		")
            loConsulta.AppendLine("			Recibos.Mon_Net								            AS Mon_Net,			")
            loConsulta.AppendLine("			Renglones_Recibos.Cod_Con					            AS Cod_Con,			")
            loConsulta.AppendLine("			Renglones_Recibos.Nom_con					            AS Nom_con,			")
            loConsulta.AppendLine("			(CASE Renglones_Recibos.Tipo ")
            loConsulta.AppendLine("				WHEN 'Asignacion' THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				WHEN 'Otro' THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				ELSE CAST(0 AS DECIMAL(28,10))")
            loConsulta.AppendLine("			END)													AS Mon_Asi,")
            loConsulta.AppendLine("			(CASE Renglones_Recibos.Tipo ")
            loConsulta.AppendLine("				WHEN 'Retencion' THEN -Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				WHEN 'Deduccion' THEN -Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				ELSE CAST(0 AS DECIMAL(28,10))")
            loConsulta.AppendLine("			END)													AS Mon_Ded,")
            loConsulta.AppendLine("			Renglones_Recibos.Val_Car								AS Val_Car")
            loConsulta.AppendLine("FROM Recibos  ")
            loConsulta.AppendLine("	JOIN Renglones_Recibos ON Renglones_Recibos.documento = Recibos.Documento")
            loConsulta.AppendLine("		AND	Renglones_Recibos.Tipo <> 'Otro'")
            loConsulta.AppendLine("	JOIN Contratos ON Contratos.Cod_Con = Recibos.Cod_Con ")
            If (lcParametro7Desde = "SI") Then
                loConsulta.AppendLine("		AND Contratos.Gen_Lib = 0")
            End If
            loConsulta.AppendLine("	JOIN Trabajadores ON Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loConsulta.AppendLine("	JOIN Cargos ON Cargos.Cod_Car = Trabajadores.Cod_Car ")
            loConsulta.AppendLine("	JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep")
            loConsulta.AppendLine("	LEFT JOIN Renglones_Campos_Nomina ON Renglones_Campos_Nomina.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("		AND	Renglones_Campos_Nomina.Cod_Cam = " & lcCodigoCampoSueldo)
            loConsulta.AppendLine("WHERE Recibos.Documento BETWEEN @lcDocumento_Desde AND @lcDocumento_Hasta")
            loConsulta.AppendLine(" 	AND Recibos.Fecha BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loConsulta.AppendLine(" 	AND Recibos.Cod_Tra BETWEEN @lcCodTra_Desde AND @lcCodTra_Hasta")
            loConsulta.AppendLine(" 	AND Recibos.Cod_Con	BETWEEN @lcCodCon_Desde AND @lcCodCon_Hasta")
            loConsulta.AppendLine(" 	AND Trabajadores.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loConsulta.AppendLine(" 	AND Trabajadores.Cod_Suc BETWEEN @lcCodSuc_Desde AND @lcCodSuc_Hasta")
            loConsulta.AppendLine(" 	AND Recibos.Status IN (" & lcParametro6Desde & ")")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento & ", ")
            loConsulta.AppendLine("			(CASE Renglones_Recibos.Tipo ")
            loConsulta.AppendLine("				WHEN 'Asignacion' THEN 0")
            loConsulta.AppendLine("				WHEN 'Otro' THEN 1")
            loConsulta.AppendLine("				WHEN 'Retencion' THEN 2")
            loConsulta.AppendLine("				ELSE 3")
            loConsulta.AppendLine("			END) ASC, ")
            loConsulta.AppendLine("			Renglones_Recibos.renglon")
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loConsulta.ToString())
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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rRecibo_Pago_NominaCGS", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rRecibo_Pago_NominaCGS.ReportSource = loObjetoReporte


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
' RJG: 16/03/15: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
