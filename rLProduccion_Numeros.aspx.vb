﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLProduccion_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rLProduccion_Numeros

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Libres_Produccion.Status, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Documento, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           (Libres_Produccion.Mon_Imp1 + Libres_Produccion.Mon_Imp2 + Libres_Produccion.Mon_Imp3)  AS Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Mon_Sal, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Transportes.Nom_Tra ")
            loComandoSeleccionar.AppendLine(" FROM      Libres_Produccion, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Transportes, ")
            loComandoSeleccionar.AppendLine("           Vendedores ")
            loComandoSeleccionar.AppendLine(" WHERE     Libres_Produccion.Cod_Cli			=   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Cod_Ven		=   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Cod_Tra		=   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Documento		Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Fec_Ini		Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Cod_Cli       Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Cod_Ven       Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Status        IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Cod_Tra       Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Cod_Mon       Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Libres_Produccion.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLProduccion_Numeros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLProduccion_Numeros.ReportSource = loObjetoReporte

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
' JJD: 09/01/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
