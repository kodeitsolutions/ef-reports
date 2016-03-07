'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOportunidades_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class rOportunidades_Clientes

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

            loComandoSeleccionar.AppendLine(" 	SELECT 		")
            loComandoSeleccionar.AppendLine(" 			Clientes.Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 			Oportunidades.Cod_Opo,")
            loComandoSeleccionar.AppendLine(" 			Oportunidades.Nom_Opo,")
            loComandoSeleccionar.AppendLine(" 			Oportunidades.Procedencia,")
            loComandoSeleccionar.AppendLine(" 			Oportunidades.Fec_Ini,")
            loComandoSeleccionar.AppendLine(" 			Oportunidades.Fec_Fin,")
            loComandoSeleccionar.AppendLine(" 			Oportunidades.Status,")
            loComandoSeleccionar.AppendLine(" 			Oportunidades.Cod_Ven,")
            loComandoSeleccionar.AppendLine(" 			Vendedores.Nom_Ven,")
            loComandoSeleccionar.AppendLine(" 			Oportunidades.Mon_Net,")
            loComandoSeleccionar.AppendLine(" 			Oportunidades.Etapa,")
            loComandoSeleccionar.AppendLine(" 			Oportunidades.Tip_Pro,")
            loComandoSeleccionar.AppendLine(" 			Oportunidades.Tip_Opo,")
            loComandoSeleccionar.AppendLine(" 			Oportunidades.Comentario")
            loComandoSeleccionar.AppendLine(" FROM	Oportunidades,Clientes,Vendedores")
            loComandoSeleccionar.AppendLine(" WHERE Oportunidades.Cod_Reg  = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("           AND Oportunidades.Cod_Ven  = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("           AND Oportunidades.Origen       =   'Clientes' ")
            loComandoSeleccionar.AppendLine("           AND Oportunidades.Cod_Reg    BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Oportunidades.Fec_Ini    BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Oportunidades.Status     IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Oportunidades.Cod_Ven    BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOportunidades_Clientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOportunidades_Clientes.ReportSource = loObjetoReporte

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
' MAT: 11/10/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
