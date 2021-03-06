﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rFaltantes"
'-------------------------------------------------------------------------------------------'
Partial Class rFaltantes

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

			
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Documento,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Factura,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Control,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Status,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 		Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Fec_ini,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Fec_Fin,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Cod_Ven,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Cod_Tra,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Cod_Suc,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Cod_Mon,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Cod_For,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Mon_Net,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Mon_Imp1,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Mon_Des1,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Mon_Rec1,")
            loComandoSeleccionar.AppendLine(" 		Faltantes.Mon_Sal,")
            loComandoSeleccionar.AppendLine(" 		CONVERT(nchar(30), Faltantes.Fec_Ini,112) AS Fecha2")
            loComandoSeleccionar.AppendLine(" FROM Faltantes")
            loComandoSeleccionar.AppendLine(" JOIN Clientes ON Clientes.Cod_Cli = Faltantes.Cod_Cli")
            loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Faltantes.Documento Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Faltantes.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)

'me.mEscribirConsulta(loComandoSeleccionar.ToString())
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rFaltantes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrFaltantes.ReportSource = loObjetoReporte

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
' CMS: 24/09/08: Codigo inicial.
'-------------------------------------------------------------------------------------------'