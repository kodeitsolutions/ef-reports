'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLSolicitudes_Requerimientos_Vendedor"
'-------------------------------------------------------------------------------------------'
Partial Class rLSolicitudes_Requerimientos_Vendedor

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		Solicitudes.Documento, ")
            loComandoSeleccionar.AppendLine(" 		Solicitudes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 		Solicitudes.Requerimiento, ")
            loComandoSeleccionar.AppendLine(" 		Solicitudes.Asunto, ")
            loComandoSeleccionar.AppendLine(" 		Solicitudes.Comentario, ")
            loComandoSeleccionar.AppendLine(" 		Solicitudes.Etapa, ")
            loComandoSeleccionar.AppendLine(" 		Solicitudes.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 		Solicitudes.Status, ")
            loComandoSeleccionar.AppendLine(" 		Factory_Global.dbo.Usuarios.Nom_Usu, ")
            loComandoSeleccionar.AppendLine(" 		Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine(" 		Clientes.Nom_cli, ")
            loComandoSeleccionar.AppendLine(" 		Clientes.Cod_cli")
            loComandoSeleccionar.AppendLine(" FROM Solicitudes")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores ON Vendedores.Cod_Ven = Solicitudes.Cod_Ven")
            loComandoSeleccionar.AppendLine(" JOIN Factory_Global.dbo.Usuarios ON Factory_Global.dbo.Usuarios.Cod_usu collate database_default= Solicitudes.Cod_usu collate database_default")
            loComandoSeleccionar.AppendLine(" JOIN Clientes ON Clientes.Cod_Cli = Solicitudes.Cod_reg")
            loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Solicitudes.Documento                  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Solicitudes.Fec_Ini                 Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Solicitudes.Status IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Solicitudes.Cod_Usu                  Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Solicitudes.cod_ven                  Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLSolicitudes_Requerimientos_Vendedor", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrLSolicitudes_Requerimientos_Vendedor.ReportSource = loObjetoReporte

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
' CMS: 14/09/09: Codigo inicial.
'-------------------------------------------------------------------------------------------'
' MAT: 18/04/11: Mejora en la vista de Diseño
'-------------------------------------------------------------------------------------------'
' EAG: 31/07/15: Se agregaron campos faltantes al .rpt
'-------------------------------------------------------------------------------------------'