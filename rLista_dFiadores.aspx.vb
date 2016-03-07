Imports System.Data
Partial Class rLista_dFiadores
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

            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	fiadores.origen, ")
            loComandoSeleccionar.AppendLine("		fiadores.cod_reg, ")
            loComandoSeleccionar.AppendLine("		fiadores.ced_fia, ")
            loComandoSeleccionar.AppendLine("		fiadores.nom_fia, ")
            loComandoSeleccionar.AppendLine("		fiadores.clase, ")
            loComandoSeleccionar.AppendLine("		fiadores.tipo, ")
            loComandoSeleccionar.AppendLine("		fiadores.status, ")
            loComandoSeleccionar.AppendLine("		clientes.nom_cli ")
            loComandoSeleccionar.AppendLine(" FROM	fiadores, ")
            loComandoSeleccionar.AppendLine("		clientes ")
            loComandoSeleccionar.AppendLine(" WHERE	fiadores.cod_reg = clientes.cod_cli ")
            loComandoSeleccionar.AppendLine(" AND 	fiadores.origen IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" AND 	fiadores.cod_reg between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" AND 	" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" AND 	fiadores.nom_fia between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" AND 	" & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" AND 	fiadores.clase between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" AND 	" & lcParametro3Hasta)

            loComandoSeleccionar.AppendLine(" AND 	fiadores.tipo between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" AND 	" & lcParametro4Hasta)

            loComandoSeleccionar.AppendLine(" AND 	fiadores.status IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine(" ORDER BY fiadores." & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLista_dFiadores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLista_dFiadores.ReportSource = loObjetoReporte


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
' YJP: 23/04/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 15/08/09: Se incluyo el orden de los registros
'-------------------------------------------------------------------------------------------'
' CMS: 28/05/10: Validacion de registro cero
'-------------------------------------------------------------------------------------------'