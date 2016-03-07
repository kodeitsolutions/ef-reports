Imports System.Data
Partial Class rSeguimientos_Proveedores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Seguimientos.Cod_Reg    As  Cod_Reg, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro     As  Nom_Reg, ")
            loComandoSeleccionar.AppendLine("           Seguimientos.Cod_Ven    As  Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven      As  Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Seguimientos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Seguimientos.Hor_Ini, ")
            loComandoSeleccionar.AppendLine("           Seguimientos.Status, ")
            loComandoSeleccionar.AppendLine("           Seguimientos.Contacto, ")
            loComandoSeleccionar.AppendLine("           Seguimientos.Lugar, ")
            loComandoSeleccionar.AppendLine("           Seguimientos.Accion, ")
            loComandoSeleccionar.AppendLine("           Seguimientos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Seguimientos.Notas ")
            loComandoSeleccionar.AppendLine(" FROM      Seguimientos, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Vendedores ")
            loComandoSeleccionar.AppendLine(" WHERE     Seguimientos.Cod_Reg        =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           And Seguimientos.Cod_Ven    =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Seguimientos.Tipo       =   'Proveedores' ")
            loComandoSeleccionar.AppendLine("           And Seguimientos.Cod_Reg    Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Seguimientos.Fec_Ini    Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Seguimientos.Status     IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine("           And Seguimientos.Cod_Ven    Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Seguimientos.Cod_Reg, ")
            'loComandoSeleccionar.AppendLine("           Seguimientos.Fec_Ini ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rSeguimientos_Proveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrSeguimientos_Proveedores.ReportSource = loObjetoReporte

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
' JJD: 21/03/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 13/04/10: Verificacion de registro Cero
'-------------------------------------------------------------------------------------------'