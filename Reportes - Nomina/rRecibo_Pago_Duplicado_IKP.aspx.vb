'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRecibo_Pago_Duplicado_IKP"
'-------------------------------------------------------------------------------------------'
Partial Class rRecibo_Pago_Duplicado_IKP
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
            loConsulta.AppendLine("			Trabajadores.Fec_Ini						AS Ingreso,			")
            loConsulta.AppendLine("			Trabajadores.Cod_Dep        				AS Cod_Dep,			")
            loConsulta.AppendLine("			COALESCE(Departamentos_Nomina.Nom_Dep,		                    ")
            loConsulta.AppendLine("			    Trabajadores.Cod_Dep)   				AS Nom_Dep,			")
            loConsulta.AppendLine("			Trabajadores.Cod_Car        				AS Cod_Car,			")
            loConsulta.AppendLine("			COALESCE(Cargos.Nom_Car,		                                ")
            loConsulta.AppendLine("			    Trabajadores.Cod_Car)   				AS Nom_Car,			")
            loConsulta.AppendLine("			COALESCE(Renglones_Campos_Nomina.val_num,		                ")
            loConsulta.AppendLine("				CAST(0 AS DECIMAL(28,10)))				AS Sueldo_Mensual,	")
            loConsulta.AppendLine("			Trabajadores.Cod_Suc						AS Cod_Suc,			")
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
            loConsulta.AppendLine("			(CASE WHEN Renglones_Recibos.Tipo <> 'Otro'				        ")
            loConsulta.AppendLine("			    THEN 'Trabajador' ELSE 'Empresa' END)				AS Grupo_Renglones, ")
            loConsulta.AppendLine("         Renglones_Recibos.Renglon                               AS Renglon")
            loConsulta.AppendLine("INTO     #tmpRecibos  ")
            loConsulta.AppendLine("FROM		Recibos  ")
            loConsulta.AppendLine("	JOIN	Renglones_Recibos   ")
            loConsulta.AppendLine("		ON	Renglones_Recibos.documento = Recibos.Documento")
            'loConsulta.AppendLine("		AND	Renglones_Recibos.Tipo <> 'Otro'")
            loConsulta.AppendLine("	JOIN	Trabajadores   ")
            loConsulta.AppendLine("		ON	Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loConsulta.AppendLine("	LEFT JOIN Departamentos_Nomina")
            loConsulta.AppendLine("		ON	Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep ")
            loConsulta.AppendLine("	LEFT JOIN Cargos")
            loConsulta.AppendLine("		ON	Cargos.Cod_Car = Trabajadores.Cod_Car ")
            loConsulta.AppendLine("	LEFT JOIN Renglones_Campos_Nomina")
            loConsulta.AppendLine("		ON	Renglones_Campos_Nomina.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("		AND	Renglones_Campos_Nomina.Cod_Cam = " & lcCodigoCampoSueldo)
            loConsulta.AppendLine("WHERE	Renglones_Recibos.Cod_Con NOT IN ('A011', 'A013')")
            loConsulta.AppendLine("		AND Recibos.Documento				BETWEEN " & lcParametro0Desde)
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
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   #tmpRecibos.*, Sueldos.Monto_Periodo, ")
            loConsulta.AppendLine("         Sucursales.nom_suc                       AS Nombre_Empresa_Cliente,  ")
            loConsulta.AppendLine("         COALESCE(campos_propiedades.val_car, '') AS Rif_Empresa_Cliente,  ")
            loConsulta.AppendLine("         Sucursales.direccion                     AS Direccion_Empresa_Cliente,  ")
            loConsulta.AppendLine("         Sucursales.telefonos                     AS Telefono_Empresa_Cliente  ")
            loConsulta.AppendLine("FROM     #tmpRecibos")
            loConsulta.AppendLine("    JOIN(")
            loConsulta.AppendLine("        SELECT  SUM(#tmpRecibos.mon_asi) AS Monto_Periodo, documento")
            loConsulta.AppendLine("        FROM    #tmpRecibos")
            loConsulta.AppendLine("        WHERE   #tmpRecibos.Tipo <> 'Otro'")
            loConsulta.AppendLine("        GROUP BY Documento")
            loConsulta.AppendLine("    ) AS Sueldos ON  Sueldos.Documento = #tmpRecibos.Documento")
            loConsulta.AppendLine("  JOIN     Sucursales ON Sucursales.cod_suc = #tmpRecibos.cod_suc")
            loConsulta.AppendLine("  LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("        AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("        AND campos_propiedades.cod_pro = 'SUC-RIF'")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento & ", ")
            loConsulta.AppendLine("			(CASE #tmpRecibos.Tipo ")
            loConsulta.AppendLine("				WHEN 'Asignacion' THEN 0")
            loConsulta.AppendLine("				WHEN 'Deduccion' THEN 2")
            loConsulta.AppendLine("				WHEN 'Retencion' THEN 3")
            loConsulta.AppendLine("				ELSE 4")
            loConsulta.AppendLine("			END) ASC, ")
            loConsulta.AppendLine("			renglon")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("DROP TABLE #tmpRecibos")
            loConsulta.AppendLine("")
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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRecibo_Pago_Duplicado_IKP", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRecibo_Pago_Duplicado_IKP.ReportSource = loObjetoReporte


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
' RJG: 01/08/13: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
