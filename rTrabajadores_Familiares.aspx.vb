'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTrabajadores_Familiares"
'-------------------------------------------------------------------------------------------'
Partial Class rTrabajadores_Familiares
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT      Trabajadores.cod_tra   AS cod_tra,")
            loComandoSeleccionar.AppendLine("            Trabajadores.nom_tra   AS nom_tra,")
            loComandoSeleccionar.AppendLine("            Trabajadores.cedula    AS cedula,")
            loComandoSeleccionar.AppendLine("            Trabajadores.fec_nac   AS Fec_Nac_Tra,")
            loComandoSeleccionar.AppendLine("            Trabajadores.sexo      AS sexo_tra,")
            loComandoSeleccionar.AppendLine("            Trabajadores.cod_est   AS cod_est,")
            loComandoSeleccionar.AppendLine("            Familiares.nom_fam     AS nom_fam,")
            loComandoSeleccionar.AppendLine("            Familiares.ape_fam     AS ape_fam,")
            loComandoSeleccionar.AppendLine("            Familiares.parentesco  AS parentesco,")
            loComandoSeleccionar.AppendLine("            Familiares.ced_fam     AS ced_fam,")
            loComandoSeleccionar.AppendLine("            Familiares.fec_nac     AS Fec_Nac_Fam,")
            loComandoSeleccionar.AppendLine("            Familiares.sexo        AS sexo_fam")
            loComandoSeleccionar.AppendLine("FROM        Trabajadores")
            loComandoSeleccionar.AppendLine("    JOIN    Familiares ")
            loComandoSeleccionar.AppendLine("        ON  Familiares.cod_tra = Trabajadores.cod_tra")
            loComandoSeleccionar.AppendLine("        AND Familiares.Adicional = 'Trabajadores.Trabajador'")
            loComandoSeleccionar.AppendLine("        AND Trabajadores.tip_tra = 'Trabajador'")
            loComandoSeleccionar.AppendLine("WHERE	     Trabajadores.Cod_Tra BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND	 Trabajadores.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("		AND  Trabajadores.Cod_Con BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND  Trabajadores.Cod_Dep BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND  Trabajadores.Cod_Suc BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTrabajadores_Familiares", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTrabajadores_Familiares.ReportSource = loObjetoReporte

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
' RJG: 10/05/13: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
