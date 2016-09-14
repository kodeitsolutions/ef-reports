﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------------------'
' Inicio de clase "IGE_fRecibo_Pago"
'-------------------------------------------------------------------------------------------'
Partial Class IGE_fRecibo_Pago
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcCodigoCampoSueldo As String
            lcCodigoCampoSueldo = goServicios.mObtenerCampoFormatoSQL(CStr(goOpciones.mObtener("CAMSUEMEN", "C")))

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Recibos.Documento							AS Documento,		")
            loConsulta.AppendLine("			Recibos.fecha								AS Fecha,			")
            loConsulta.AppendLine("			Recibos.[Status]							AS Estatus,			")
            loConsulta.AppendLine("			Recibos.Cod_Con								AS Cod_Contrato,	")
            loConsulta.AppendLine("			Recibos.Cod_Tra								AS Cod_Tra,			")

            loConsulta.AppendLine("			Trabajadores.Nom_Tra						AS Nom_Tra,			")
            loConsulta.AppendLine("			Trabajadores.Cedula							AS Cedula,			")
            loConsulta.AppendLine("			Trabajadores.Num_Cue						AS Num_Cue,			")
            loConsulta.AppendLine("			Trabajadores.Fec_Ini						AS Fec_Ing,			")

            loConsulta.AppendLine("			Bancos.Nom_Ban	        					AS Nom_Ban,			")

            loConsulta.AppendLine("			Cargos.Nom_Car						        AS Nom_Car,			")
            
            loConsulta.AppendLine("			COALESCE(Renglones_Campos_Nomina.val_num,						")
            loConsulta.AppendLine("				CAST(0 AS DECIMAL(28,10)))				AS Sueldo_Mensual,	")
            loConsulta.AppendLine("			COALESCE(Renglones_Campos_Nomina.val_num/30,					")
            loConsulta.AppendLine("				CAST(0 AS DECIMAL(28,10)))				AS Sueldo_Diario,	")
            loConsulta.AppendLine("			Recibos.Fec_Ini								AS Fec_Ini,			")
            loConsulta.AppendLine("			Recibos.Fec_Fin								AS Fec_Fin,			")
            loConsulta.AppendLine("			Recibos.Cod_Rev								AS Cod_Rev,			")
            loConsulta.AppendLine("			Recibos.Comentario							AS Comentario,		")
            loConsulta.AppendLine("			Recibos.Mon_Net								AS Mon_Net,			")
            loConsulta.AppendLine("			Renglones_Recibos.Cod_Con					AS Cod_Con,			")
            loConsulta.AppendLine("			Renglones_Recibos.Nom_con					AS Nom_con,			")
            loConsulta.AppendLine("			Renglones_Recibos.tipo						AS Tipo,			")
            loConsulta.AppendLine("			(CASE Renglones_Recibos.Tipo ")
            loConsulta.AppendLine("				WHEN 'Asignacion' THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				ELSE CAST(0 AS DECIMAL(28,10))")
            loConsulta.AppendLine("			END)													AS Mon_Asi,")
            loConsulta.AppendLine("			(CASE Renglones_Recibos.Tipo ")
            loConsulta.AppendLine("				WHEN 'Retencion' THEN -Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				WHEN 'Deduccion' THEN -Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				ELSE CAST(0 AS DECIMAL(28,10))")
            loConsulta.AppendLine("			END)													AS Mon_Ded,")
            loConsulta.AppendLine("			Renglones_Recibos.Val_Num								AS Val_Num,")
            loConsulta.AppendLine("			Renglones_Recibos.Val_Car								AS Val_Car")
            loConsulta.AppendLine("FROM		Recibos  ")
            loConsulta.AppendLine("	JOIN	Renglones_Recibos   ")
            loConsulta.AppendLine("		ON	Renglones_Recibos.documento = Recibos.Documento")
            loConsulta.AppendLine("	JOIN	Trabajadores   ")
            loConsulta.AppendLine("		ON	Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loConsulta.AppendLine("	JOIN	Bancos   ")
            loConsulta.AppendLine("		ON	Trabajadores.Cod_Ban = Bancos.Cod_Ban ")
            loConsulta.AppendLine("	JOIN	Cargos   ")
            loConsulta.AppendLine("		ON	Cargos.Cod_Car = Trabajadores.Cod_Car ")
            loConsulta.AppendLine("	LEFT JOIN Renglones_Campos_Nomina")
            loConsulta.AppendLine("		ON	Renglones_Campos_Nomina.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("		AND	Renglones_Campos_Nomina.Cod_Cam = " & lcCodigoCampoSueldo)
            loConsulta.AppendLine("WHERE	" & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("     AND Renglones_Recibos.Tipo <> 'Otro' ")
            loConsulta.AppendLine("ORDER BY	(CASE Renglones_Recibos.Tipo ")
            loConsulta.AppendLine("				WHEN 'Asignacion' THEN 0")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("IGE_fRecibo_Pago", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvIGE_fRecibo_Pago.ReportSource = loObjetoReporte

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
' RJG: 20/02/13: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'