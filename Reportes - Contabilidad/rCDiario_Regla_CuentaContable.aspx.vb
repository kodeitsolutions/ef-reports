'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCDiario_Regla_CuentaContable"
'-------------------------------------------------------------------------------------------'
Partial Class rCDiario_Regla_CuentaContable
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

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine("DECLARE @lcParametro1Desde VARCHAR(100)	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro1Hasta VARCHAR(100)	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro2Desde DATETIME	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro2Hasta DATETIME	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro3Desde VARCHAR(100)	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro3Hasta VARCHAR(100)	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro4Desde VARCHAR(100)	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro4Hasta VARCHAR(100)	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro5Desde VARCHAR(100)	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro5Hasta VARCHAR(100)	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro6Hasta VARCHAR(100)	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro6Desde VARCHAR(100)	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro7Desde VARCHAR(100)	")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro7Hasta VARCHAR(100)	")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SET		@lcParametro1Desde = " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro1Hasta = " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("SET		@lcParametro2Desde = " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro2Hasta = " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("SET		@lcParametro3Desde = " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro3Hasta = " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("SET		@lcParametro4Desde = " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro4Hasta = " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("SET		@lcParametro5Desde = " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro5Hasta = " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("SET		@lcParametro6Desde = " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro6Hasta = " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("SET		@lcParametro7Desde = " & lcParametro6Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro7Hasta = " & lcParametro6Hasta)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10)")
			loComandoSeleccionar.AppendLine("DECLARE @llFalso BIT")
			loComandoSeleccionar.AppendLine("DECLARE @llVerdadero BIT")		 
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SET	@lnCero = 0")
			loComandoSeleccionar.AppendLine("SET	@llFalso = 0")
			loComandoSeleccionar.AppendLine("SET	@llVerdadero = 1")
			loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	SUM(RC.Mon_Deb)									AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		SUM(RC.Mon_Hab)									AS Mon_Hab,")
			loComandoSeleccionar.AppendLine("		(CASE WHEN RC.Cod_Reg = '' THEN 'Sin Regla' ELSE RC.Cod_Reg END)	AS Cod_Reg,")
			loComandoSeleccionar.AppendLine("		ISNULL(Reglas_Integracion.Nom_Reg, 'Sin Regla')	As Nom_Reg,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Cod_Cue 						As Cod_Cue,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Nom_Cue 						AS Nom_Cue")
            loComandoSeleccionar.AppendLine("FROM	Comprobantes ")
            loComandoSeleccionar.AppendLine("	JOIN (")
            loComandoSeleccionar.AppendLine("			SELECT	Renglones_Comprobantes.Documento	AS Documento,")
            loComandoSeleccionar.AppendLine("					Renglones_Comprobantes.Cod_Cue		AS Cod_Cue	,")
            loComandoSeleccionar.AppendLine("					Renglones_Comprobantes.Cod_Reg		AS Cod_Reg	,")
            loComandoSeleccionar.AppendLine("					SUM(Renglones_Comprobantes.Mon_Deb) AS Mon_Deb	,")
            loComandoSeleccionar.AppendLine("					SUM(Renglones_Comprobantes.Mon_Hab) AS Mon_Hab	")
            loComandoSeleccionar.AppendLine("			FROM	Renglones_Comprobantes")
            loComandoSeleccionar.AppendLine("			WHERE	Renglones_Comprobantes.Documento BETWEEN @lcParametro1Desde AND	@lcParametro1Hasta")
            loComandoSeleccionar.AppendLine("				AND	Renglones_Comprobantes.Fec_Ini BETWEEN @lcParametro2Desde AND	@lcParametro2Hasta")
            loComandoSeleccionar.AppendLine("				AND	Renglones_Comprobantes.Cod_Cue BETWEEN @lcParametro3Desde AND	@lcParametro3Hasta")
            loComandoSeleccionar.AppendLine("				AND	Renglones_Comprobantes.Cod_Cen BETWEEN @lcParametro4Desde AND	@lcParametro4Hasta")
            loComandoSeleccionar.AppendLine("				AND	Renglones_Comprobantes.Cod_Gas BETWEEN @lcParametro5Desde AND	@lcParametro5Hasta")
            loComandoSeleccionar.AppendLine("				AND	Renglones_Comprobantes.Cod_Aux BETWEEN @lcParametro6Desde AND	@lcParametro6Hasta")
            loComandoSeleccionar.AppendLine("				AND	Renglones_Comprobantes.Cod_Mon BETWEEN @lcParametro7Desde AND	@lcParametro7Hasta")
            loComandoSeleccionar.AppendLine("			GROUP BY Documento, Cod_Cue, Cod_Reg ")
            loComandoSeleccionar.AppendLine("		) AS RC ")
            loComandoSeleccionar.AppendLine("		ON	RC.Documento = Comprobantes.Documento")
            loComandoSeleccionar.AppendLine("		AND	Comprobantes.Status = 'Pendiente'")
            loComandoSeleccionar.AppendLine("	JOIN	Cuentas_Contables ON Cuentas_Contables.Cod_Cue = RC.Cod_Cue   ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Reglas_Integracion ON Reglas_Integracion.Cod_Reg = RC.Cod_Reg")
            loComandoSeleccionar.AppendLine("GROUP BY	RC.Cod_Reg, ISNULL(Reglas_Integracion.Nom_Reg, 'Sin Regla'), ")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Cod_Cue, Cuentas_Contables.Nom_Cue")
            loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


			'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
 

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCDiario_Regla_CuentaContable", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCDiario_Regla_CuentaContable.ReportSource = loObjetoReporte

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
' RJG: 18/11/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
