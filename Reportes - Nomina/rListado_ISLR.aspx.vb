'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rListado_ISLR"
'-------------------------------------------------------------------------------------------'
Partial Class rListado_ISLR
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            Dim lcCampoSueldo As String

            lcCampoSueldo = goServicios.mObtenerCampoFormatoSQL(CStr(goOpciones.mObtener("CAMSUEMEN", "")))

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Recibos.documento                                                       AS documento, ")
            loComandoSeleccionar.AppendLine("		Recibos.cod_tra                                                         AS cod_tra, ")
            loComandoSeleccionar.AppendLine("		Trabajadores.nom_tra                                                    AS nom_tra, ")
            loComandoSeleccionar.AppendLine("		Recibos.fecha                                                           AS fecha, ")
            loComandoSeleccionar.AppendLine("		Renglones_recibos.val_num                                               AS val_num, ")
            loComandoSeleccionar.AppendLine("       Case Renglones_recibos.val_num ")
            loComandoSeleccionar.AppendLine("			WHEN 0 THEN 0 ")
            loComandoSeleccionar.AppendLine("			ELSE Renglones_Recibos.mon_net*100/Renglones_recibos.val_num  ")
            loComandoSeleccionar.AppendLine("		END                                                                     AS Porcentaje, ")
            loComandoSeleccionar.AppendLine("       Renglones_Recibos.mon_net                                               AS mon_net")
            loComandoSeleccionar.AppendLine("FROM Recibos ")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Recibos ON Renglones_Recibos.documento = Recibos.Documento ")
            loComandoSeleccionar.AppendLine("			AND Renglones_Recibos.cod_con IN ('R005','R405') ")
            loComandoSeleccionar.AppendLine("	JOIN Trabajadores ON Trabajadores.cod_tra = Recibos.cod_tra ")
            loComandoSeleccionar.AppendLine("WHERE		Recibos.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND	Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND	Recibos.Cod_Tra BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND	Recibos.Cod_Con BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND	Trabajadores.Cod_Dep BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("		AND	Trabajadores.Cod_Suc BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND	Recibos.Status IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("		AND	Recibos.Cod_Rev BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY	Trabajadores.nom_tra,Recibos.cod_tra,Recibos.fecha DESC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rListado_ISLR", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrListado_ISLR.ReportSource = loObjetoReporte

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
' EAG: 10/10/15: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
