'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_fRecibo_Liquidacion"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_fRecibo_Liquidacion
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcCodigoCampoSueldo As String
            lcCodigoCampoSueldo = goServicios.mObtenerCampoFormatoSQL(CStr(goOpciones.mObtener("CAMSUEMEN", "C")))

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Recibos.Documento							AS Documento,		")
            loConsulta.AppendLine("			Recibos.Fecha								AS Fecha,			")
            loConsulta.AppendLine("			Recibos.Cod_Tra								AS Cod_Tra,			")
            loConsulta.AppendLine("			Trabajadores.Nom_Tra						AS Nom_Tra,			")
            loConsulta.AppendLine("			Trabajadores.Cedula							AS Cedula,			")
            loConsulta.AppendLine("			COALESCE(Renglones_Campos_Nomina.Val_Num,						")
            loConsulta.AppendLine("				CAST(0 AS DECIMAL(28,10)))				AS Sueldo_Mensual,	")
            loConsulta.AppendLine("			Trabajadores.Fec_Ini						AS Ingreso,			")
            loConsulta.AppendLine("			Cargos.Nom_Car								AS Nom_Car,			")
            loConsulta.AppendLine("			Recibos.Fecha								AS Fecha,			")
            'loConsulta.AppendLine("			(CASE ")
            'loConsulta.AppendLine("				WHEN	(CASE WHEN DAY(Trabajadores.Fec_Ini) - DAY(Recibos.Fec_Fin) > 0")
            'loConsulta.AppendLine("							  THEN DATEDIFF(MONTH, Trabajadores.Fec_Ini,Recibos.Fec_Fin) - 1")
            'loConsulta.AppendLine("							  ELSE DATEDIFF(MONTH, Trabajadores.Fec_Ini,Recibos.Fec_Fin)")
            'loConsulta.AppendLine("						END)%12 = 0 ")
            'loConsulta.AppendLine("					THEN  DATEDIFF(YEAR, Trabajadores.Fec_Ini, Recibos.Fec_Fin)")
            'loConsulta.AppendLine("				WHEN DATEDIFF(YEAR, Trabajadores.Fec_Ini, Recibos.Fec_Fin) > 0")
            'loConsulta.AppendLine("					THEN DATEDIFF(YEAR, Trabajadores.Fec_Ini, Recibos.Fec_Fin) - 1")
            'loConsulta.AppendLine("					ELSE DATEDIFF(YEAR, Trabajadores.Fec_Ini, Recibos.Fec_Fin)")
            'loConsulta.AppendLine("			END)										AS Año,")
            'loConsulta.AppendLine("		    DATEDIFF(YEAR,Trabajadores.Fec_Ini, Recibos.Fec_Fin) AS Año,")
            loConsulta.AppendLine("         DATEDIFF(YEAR, Trabajadores.Fec_Ini, Recibos.Fec_Fin) AS Año,")
            'loConsulta.AppendLine("			(CASE WHEN DAY(Trabajadores.Fec_Ini) - Day(Recibos.Fec_Fin) > 0 ")
            'loConsulta.AppendLine("				Then DATEDIFF(MONTH, Trabajadores.Fec_Ini, Recibos.Fec_Fin)%12 - 1")
            'loConsulta.AppendLine("				Else DATEDIFF(MONTH, Trabajadores.Fec_Ini,Recibos.Fec_Fin)%12")
            'loConsulta.AppendLine("			End)										As Mes,")
            loConsulta.AppendLine("         ABS(DATEDIFF(mm, DATEADD(yy,DATEDIFF(YEAR, Trabajadores.Fec_Ini, Recibos.Fec_Fin), Trabajadores.Fec_Ini), Recibos.Fec_Fin)) AS Mes,")
            'loConsulta.AppendLine("			DATEDIFF(DAY,")
            'loConsulta.AppendLine("					DATEADD(MONTH,(Case When DAY(Trabajadores.Fec_Ini) - DAY(Recibos.Fec_Fin) >0 ")
            'loConsulta.AppendLine("										Then DATEDIFF(MONTH, Trabajadores.Fec_Ini, Recibos.Fec_Fin)%12 - 1")
            'loConsulta.AppendLine("										Else DATEDIFF(MONTH, Trabajadores.Fec_Ini, Recibos.Fec_Fin)%12")
            'loConsulta.AppendLine("									End),")
            'loConsulta.AppendLine("					DATEADD(YEAR, (Case ")
            'loConsulta.AppendLine("										When(Case When DAY(Trabajadores.Fec_Ini) - DAY(Recibos.Fec_Fin) >0 ")
            'loConsulta.AppendLine("												  Then DATEDIFF(MONTH, Trabajadores.Fec_Ini, Recibos.Fec_Fin) - 1")
            'loConsulta.AppendLine("												  Else DATEDIFF(MONTH, Trabajadores.Fec_Ini, Recibos.Fec_Fin)")
            'loConsulta.AppendLine("											 End)%12 = 0 ")
            'loConsulta.AppendLine("											Then  DATEDIFF(YEAR, Trabajadores.Fec_Ini, Recibos.Fec_Fin)")
            'loConsulta.AppendLine("										When DATEDIFF(YEAR, Trabajadores.Fec_Ini, Recibos.Fec_Fin) > 0")
            'loConsulta.AppendLine("											Then  DATEDIFF(YEAR, Trabajadores.Fec_Ini, Recibos.Fec_Fin) - 1")
            'loConsulta.AppendLine("											Else DATEDIFF(YEAR, Trabajadores.Fec_Ini, Recibos.Fec_Fin)")
            'loConsulta.AppendLine("									End),Trabajadores.Fec_Ini)), ")
            'loConsulta.AppendLine("					Recibos.Fec_Fin)					As Dia,")
            loConsulta.AppendLine("         ABS(DATEDIFF(dd, DATEADD(mm, DATEDIFF(mm, DATEADD(yy, DATEDIFF(Year, Trabajadores.Fec_Ini, Recibos.Fec_Fin), Trabajadores.Fec_Ini), Recibos.Fec_Fin), DATEADD(yy, DATEDIFF(Year, Trabajadores.Fec_Ini, Recibos.Fec_Fin), Trabajadores.Fec_Ini)), Recibos.Fec_Fin)) As Dia,")
            loConsulta.AppendLine("			Recibos.Fec_Fin								As Fec_Fin,			")
            loConsulta.AppendLine("			DATEADD(DAY, 1, Recibos.Fec_Fin)			As Fec_Rei,			")
            loConsulta.AppendLine("			Recibos.Comentario							As Comentario,		")
            loConsulta.AppendLine("			Recibos.Mon_Net								As Mon_Net,			")
            loConsulta.AppendLine("			Renglones_Recibos.Cod_Con					As Cod_Con,			")
            loConsulta.AppendLine("			Renglones_Recibos.Nom_con					As Nom_con,			")
            loConsulta.AppendLine("			(Case When Renglones_Recibos.Tipo = 'Asignacion'")
            loConsulta.AppendLine("				  THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				  ELSE CAST(0 AS DECIMAL(28,10))")
            loConsulta.AppendLine("			END)													AS Mon_Asi,")
            loConsulta.AppendLine("			(CASE Renglones_Recibos.Tipo ")
            loConsulta.AppendLine("				WHEN 'Retencion' THEN -Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				WHEN 'Deduccion' THEN -Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("				ELSE CAST(0 AS DECIMAL(28,10))")
            loConsulta.AppendLine("			END)													AS Mon_Ded,")
            loConsulta.AppendLine("			Renglones_Recibos.Val_Car								AS Val_Car,")
            loConsulta.AppendLine("			COALESCE(Liquidaciones.Tipo, '')				        AS Tipo_Liquidacion")
            loConsulta.AppendLine("FROM	Recibos  ")
            loConsulta.AppendLine("     JOIN Renglones_Recibos ON Renglones_Recibos.documento = Recibos.Documento")
            loConsulta.AppendLine("	    JOIN Trabajadores ON Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loConsulta.AppendLine("	    JOIN Cargos ON Cargos.Cod_Car = Trabajadores.Cod_Car")
            loConsulta.AppendLine("	    LEFT JOIN Liquidaciones ON Liquidaciones.Documento = Renglones_Recibos.Doc_Ori")
            loConsulta.AppendLine("		    AND Renglones_Recibos.Tip_Ori = 'Liquidaciones'")
            loConsulta.AppendLine("     LEFT JOIN Renglones_Campos_Nomina ON Renglones_Campos_Nomina.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("		    AND	Renglones_Campos_Nomina.Cod_Cam = " & lcCodigoCampoSueldo)
            loConsulta.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("     AND Renglones_Recibos.Tipo <> 'Otro' ")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fRecibo_Liquidacion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_fRecibo_Liquidacion.ReportSource = loObjetoReporte

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
' EAG: 24/08/15: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
