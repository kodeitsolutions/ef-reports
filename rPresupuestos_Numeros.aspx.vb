'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPresupuestos_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rPresupuestos_Numeros

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
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("  SELECT		Presupuestos.Documento, ")
            loComandoSeleccionar.AppendLine(" 				Presupuestos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 				Presupuestos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine(" 				Presupuestos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine(" 				Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine(" 				Presupuestos.Mon_Bru, ")
            loComandoSeleccionar.AppendLine(" 				Presupuestos.Cod_Tra, ")
            loComandoSeleccionar.AppendLine(" 				Presupuestos.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine(" 				Presupuestos.Status, ")
            loComandoSeleccionar.AppendLine(" 				Presupuestos.Control, ")
            loComandoSeleccionar.AppendLine(" 				Presupuestos.Mon_Net, ")
            loComandoSeleccionar.AppendLine(" 				Presupuestos.Mon_Imp1 ")
            loComandoSeleccionar.AppendLine(" FROM      Presupuestos, ")
            loComandoSeleccionar.AppendLine("           Proveedores  ")
            loComandoSeleccionar.AppendLine(" WHERE     Presupuestos.Cod_Pro        =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           And Presupuestos.Documento  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Presupuestos.Fec_Ini    Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Presupuestos.Cod_Pro    Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Presupuestos.Status     IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           And Presupuestos.Cod_rev    Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Presupuestos.Documento ")


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPresupuestos_Numeros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPresupuestos_Numeros.ReportSource = loObjetoReporte

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
' YYG: 02/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 27/12/08: Ajustes al codigo.
'-------------------------------------------------------------------------------------------'
' CMS: 20/07/09: Se elimino el campp Mon_Sal y se agrego el campo Mon_Imp1, verificacion de 
'                registros
'-------------------------------------------------------------------------------------------'
