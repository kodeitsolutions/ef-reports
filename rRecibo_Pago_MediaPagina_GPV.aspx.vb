﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRecibo_Pago_MediaPagina_GPV"
'-------------------------------------------------------------------------------------------'
Partial Class rRecibo_Pago_MediaPagina_GPV
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
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcCodigoCampoSueldo As String
            lcCodigoCampoSueldo = goServicios.mObtenerCampoFormatoSQL(CStr(goOpciones.mObtener("CAMSUEMEN", "C")))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Recibos.Documento							AS Documento,		")
            loConsulta.AppendLine("			Recibos.fecha								AS Fecha,			")
            loConsulta.AppendLine("			Recibos.[Status]							AS Estatus,			")
            loConsulta.AppendLine("			Recibos.Cod_Con								AS Cod_Contrato,	")
            loConsulta.AppendLine("			Recibos.Cod_Tra								AS Cod_Tra,			")
            loConsulta.AppendLine("			Trabajadores.Nom_Tra						AS Nom_Tra,			")
            loConsulta.AppendLine("			Trabajadores.Cedula							AS Cedula,			")
            loConsulta.AppendLine("			COALESCE(Renglones_Campos_Nomina.val_num,						")
            loConsulta.AppendLine("				CAST(0 AS DECIMAL(28,10)))				AS Sueldo_Mensual,	")
            loConsulta.AppendLine("			Recibos.Fec_Ini								AS Fec_Ini,			")
            loConsulta.AppendLine("			Recibos.Fec_Fin								AS Fec_Fin,			")
            loConsulta.AppendLine("			Recibos.Cod_Rev								AS Cod_Rev,			")
            loConsulta.AppendLine("			Recibos.Comentario							AS Comentario,		")
            loConsulta.AppendLine("			Recibos.Mon_Net- ")
            loConsulta.AppendLine("			SUM((CASE Renglones_Recibos.Cod_Con ")
            loConsulta.AppendLine("					WHEN 'A013' THEN (CASE Renglones_Recibos.Tipo   ")
            loConsulta.AppendLine("										WHEN 'Asignacion' THEN Renglones_Recibos.Mon_Net  ")
            loConsulta.AppendLine("										WHEN 'Otro' THEN Renglones_Recibos.Mon_Net  ")
            loConsulta.AppendLine("										ELSE CAST(0 AS DECIMAL(28,10))  ")
            loConsulta.AppendLine("									END)  ")
            loConsulta.AppendLine("					ELSE 0  ")
            loConsulta.AppendLine("				END))							OVER(partition  by  trabajadores.cod_tra, recibos.documento)		AS Mon_Net,	 ")
            loConsulta.AppendLine("			Renglones_Recibos.Cod_Con					AS Cod_Con,			")
            loConsulta.AppendLine("			Renglones_Recibos.Nom_con					AS Nom_con,			")
            loConsulta.AppendLine("			Renglones_Recibos.tipo						AS Tipo,			")
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
            loConsulta.AppendLine("			Renglones_Recibos.Val_Num								AS Val_Num,")
            loConsulta.AppendLine("			Renglones_Recibos.Val_Car								AS Val_Car,")
            loConsulta.AppendLine("            'del ' + RIGHT('00'+CAST(DAY(Recibos.fec_ini) AS VARCHAR),2)+'/' ")
            loConsulta.AppendLine("			+ RIGHT('00'+CAST(MONTH(Recibos.fec_ini) AS VARCHAR),2)+'/' ")
            loConsulta.AppendLine("			+ CAST(YEAR(Recibos.fec_ini) AS VARCHAR) +' al ' ")
            loConsulta.AppendLine("			+ RIGHT('00'+CAST(DAY(Recibos.fec_fin) AS VARCHAR),2)+'/' ")
            loConsulta.AppendLine("			+ RIGHT('00'+CAST(MONTH(Recibos.fec_fin) AS VARCHAR),2)+'/' ")
            loConsulta.AppendLine("			+ CAST(YEAR(Recibos.fec_fin) AS VARCHAR)						AS Periodo, ")
            loConsulta.AppendLine("			Trabajadores.num_cue											AS Cuenta, ")
            loConsulta.AppendLine("			COALESCE(Bancos.nom_ban,'')										AS Banco, ")
            loConsulta.AppendLine("			Trabajadores.fec_ini											AS Ingreso, ")
            loConsulta.AppendLine("			Cargos.nom_car													AS Cargo, ")
            loConsulta.AppendLine("			SUM((CASE Renglones_Recibos.Cod_Con ")
            loConsulta.AppendLine("					WHEN 'A013' THEN (CASE Renglones_Recibos.Tipo  ")
            loConsulta.AppendLine("										WHEN 'Asignacion' THEN Renglones_Recibos.Mon_Net ")
            loConsulta.AppendLine("										WHEN 'Otro' THEN Renglones_Recibos.Mon_Net ")
            loConsulta.AppendLine("										ELSE CAST(0 AS DECIMAL(28,10)) ")
            loConsulta.AppendLine("									END) ")
            loConsulta.AppendLine("					ELSE 0 ")
            loConsulta.AppendLine("				END))							OVER(partition  by trabajadores.cod_tra, recibos.documento) Alimentacion, ")
            loConsulta.AppendLine("			COALESCE(Prestamo.debe,0)  AS Prestamo ")
            loConsulta.AppendLine("FROM		Recibos  ")
            loConsulta.AppendLine("	JOIN	Renglones_Recibos   ")
            loConsulta.AppendLine("		ON	Renglones_Recibos.documento = Recibos.Documento")
            loConsulta.AppendLine("		AND	Renglones_Recibos.Tipo <> 'Otro'")
            loConsulta.AppendLine("	JOIN	Trabajadores   ")
            loConsulta.AppendLine("		ON	Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loConsulta.AppendLine("	LEFT JOIN Renglones_Campos_Nomina")
            loConsulta.AppendLine("		ON	Renglones_Campos_Nomina.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("		AND	Renglones_Campos_Nomina.Cod_Cam = " & lcCodigoCampoSueldo)
            loConsulta.AppendLine("	LEFT JOIN Bancos ON Bancos.cod_ban = Trabajadores.cod_ban ")
            loConsulta.AppendLine("		JOIN	Cargos ON Cargos.cod_car = Trabajadores.cod_car ")
            loConsulta.AppendLine("	LEFT JOIN(	select sum(mon_sal)	As debe,cod_tra  ")
            loConsulta.AppendLine("            from prestamos ")
            loConsulta.AppendLine("				WHERE Prestamos.status='Confirmado' ")
            loConsulta.AppendLine("				group by cod_tra) AS Prestamo ON Prestamo.cod_tra = trabajadores.cod_tra ")
            loConsulta.AppendLine("WHERE	Recibos.Documento				BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("		AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 	AND Recibos.Fecha	            	BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine(" 		AND " & lcParametro1Hasta)
            loConsulta.AppendLine(" 	AND Recibos.Cod_Tra				    BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine(" 		AND " & lcParametro2Hasta)
            loConsulta.AppendLine(" 	AND Recibos.Cod_Con				    BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine(" 		AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 	AND Trabajadores.Cod_Dep			BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine(" 		AND " & lcParametro4Hasta)
            loConsulta.AppendLine(" 	AND Trabajadores.Cod_Suc       		BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine(" 		AND " & lcParametro5Hasta)
            loConsulta.AppendLine(" 	AND Recibos.Status				    IN (" & lcParametro6Desde & ")")
            loConsulta.AppendLine(" 	AND Recibos.Cod_Rev				    BETWEEN " & lcParametro7Desde)
            loConsulta.AppendLine(" 		AND " & lcParametro7Hasta)
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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRecibo_Pago_MediaPagina_GPV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)


            'Dim myPrintOptions As CrystalDecisions.CrystalReports.Engine.PrintOptions = loObjetoReporte.PrintOptions



            'Dim ps As New System.Drawing.Printing.PrinterSettings()
            'Dim loTamaño As New Drawing.Printing.PaperSize("MediaCarta", 850, 550) '8,5x5,5 pulgadas 

            'ps.DefaultPageSettings.PaperSize = loTamaño
            'Dim pags As New System.Drawing.Printing.PageSettings(ps)

            ''ps.DefaultPageSettings.PaperSize.Width = 8.5 * 100
            ''ps.DefaultPageSettings.PaperSize.Height = 5.5 * 100

            'loObjetoReporte.PrintOptions.CopyFrom(ps, pags)
            'loObjetoReporte.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize '(ps, ps.DefaultPageSettings)


            Me.crvrRecibo_Pago_MediaPagina_GPV.ReportSource = loObjetoReporte


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
' EAG: 09/09/15: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
' EAG: 02/10/15: Se corrigió bug en la asignacion de los prestamos y bonos alimenticios.    '
'-------------------------------------------------------------------------------------------'
