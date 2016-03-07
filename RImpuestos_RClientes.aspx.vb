'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "RImpuestos_RClientes"
'-------------------------------------------------------------------------------------------'
Partial Class RImpuestos_RClientes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("     SELECT  ")
            loComandoSeleccionar.AppendLine("              Cuentas_Cobrar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("              Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("              Cuentas_Cobrar.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("              Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("              Cuentas_Cobrar.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("              Cuentas_Cobrar.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("              Cuentas_Cobrar.Mon_Net ")
            loComandoSeleccionar.AppendLine("      FROM Cuentas_Cobrar, Clientes, Vendedores    ")
            loComandoSeleccionar.AppendLine("      WHERE (Cuentas_Cobrar.Cod_tip = 'Retiva' OR Cuentas_Cobrar.Cod_tip = 'Riva')")
            loComandoSeleccionar.AppendLine("      					AND Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("      					AND Vendedores.Cod_Ven = Cuentas_Cobrar.Cod_Ven   ")

            loComandoSeleccionar.AppendLine("             			AND Cuentas_Cobrar.Fec_Ini between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("             			AND Cuentas_Cobrar.Cod_Cli between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("             			AND Cuentas_Cobrar.Cod_Ven between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("         				AND Cuentas_Cobrar.status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("             			AND Cuentas_Cobrar.Cod_Mon between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("         				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("             			AND Clientes.Cod_Zon between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("         				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("             			AND Cuentas_Cobrar.Cod_Rev between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("         				AND " & lcParametro6Hasta)

            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("RImpuestos_RClientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvRImpuestos_RClientes.ReportSource = loObjetoReporte


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
' CMS:  20/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
