'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMovimientosBancos_TCuenta"
'-------------------------------------------------------------------------------------------'
Partial Class rMovimientosBancos_TCuenta
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosFinales(5)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("   SELECT   ")
            loComandoSeleccionar.AppendLine("   		Movimientos_Cuentas.Cod_Cue,   ")
            loComandoSeleccionar.AppendLine("   		Cuentas_Bancarias.	Nom_Cue,   ")
            loComandoSeleccionar.AppendLine("   		sum(Movimientos_Cuentas.Mon_Deb) AS Mon_Deb,   ")
            loComandoSeleccionar.AppendLine("   		sum(Movimientos_Cuentas.Mon_Hab) AS Mon_Hab,   ")
            loComandoSeleccionar.AppendLine("   		sum((Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)) AS Diferencia   ")
            loComandoSeleccionar.AppendLine("   FROM    ")
            loComandoSeleccionar.AppendLine("   		Movimientos_Cuentas,   ")
            loComandoSeleccionar.AppendLine("   		Cuentas_Bancarias   ")
            loComandoSeleccionar.AppendLine("   WHERE   ")
            loComandoSeleccionar.AppendLine("   		Movimientos_Cuentas.Cod_cue = Cuentas_Bancarias.Cod_Cue   ")
            loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Fec_Ini between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_Cue between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_Mon between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Status IN (" & lcParametro3Desde & ")")

            If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_Rev between " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_Rev NOT between " & lcParametro4Desde)
            End If

            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("   GROUP BY Movimientos_Cuentas.Cod_Cue, Cuentas_Bancarias.Nom_Cue ")
            loComandoSeleccionar.AppendLine("   ORDER BY      " & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMovimientosBancos_TCuenta", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrMovimientosBancos_TCuenta.ReportSource = loObjetoReporte


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
' CMS: 22/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 31/07/09: Filtro "Revision:", Verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS:  03/08/09: Filtro “Tipo Revisión:”
'-------------------------------------------------------------------------------------------'