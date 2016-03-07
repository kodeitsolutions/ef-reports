'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCCobrar_ClientesSucursales"
'-------------------------------------------------------------------------------------------'
Partial Class rCCobrar_ClientesSucursales
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            'Dim lcParametro12Desde As String = cusAplicacion.goReportes.paParametrosIniciales(12)
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))

            Dim loComandoSeleccionar As New StringBuilder()
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            loComandoSeleccionar.AppendLine(" SELECT    Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Control, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Suc, ")
            loComandoSeleccionar.AppendLine("           Sucursales.Nom_Suc, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Bru *(-1) ELSE Cuentas_Cobrar.Mon_Bru END) As Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Net *(-1) ELSE Cuentas_Cobrar.Mon_Net END) As Mon_Net, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal *(-1) ELSE Cuentas_Cobrar.Mon_Sal END) As Mon_Sal,  ")

            'loComandoSeleccionar.AppendLine("  		    CASE WHEN (DATALENGTH(Cuentas_Cobrar.Comentario) > 1) AND (DATALENGTH(Cuentas_Cobrar.Notas) > 1) THEN '- '+CAST(Cuentas_Cobrar.comentario AS  VARCHAR(1000))+CHAR(13)+'- '+CAST(Cuentas_Cobrar.Notas AS  VARCHAR(1000))   ")
            'loComandoSeleccionar.AppendLine("  		         WHEN (DATALENGTH(Cuentas_Cobrar.Comentario) > 1) AND (DATALENGTH(Cuentas_Cobrar.Notas) <= 1) THEN '- '+CAST(Cuentas_Cobrar.comentario AS  VARCHAR(1000))   ")
            'loComandoSeleccionar.AppendLine("  		         WHEN (DATALENGTH(Cuentas_Cobrar.Comentario) <= 1) AND (DATALENGTH(Cuentas_Cobrar.Notas) > 1) THEN '- '+CAST(Cuentas_Cobrar.Notas AS  VARCHAR(1000))   ")
            'loComandoSeleccionar.AppendLine("  		         ELSE '' ")
            'loComandoSeleccionar.AppendLine("  		    END AS Comentario ")

            loComandoSeleccionar.AppendLine("  		    SUBSTRING(Cuentas_Cobrar.Comentario,1,150)  AS Comentario ")

            loComandoSeleccionar.AppendLine(" FROM      Clientes, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Transportes, ")
            loComandoSeleccionar.AppendLine("           Sucursales, ")
            loComandoSeleccionar.AppendLine("           Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Cod_Cli          =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Ven      =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Tra      =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Mon      =   Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Suc      =   Sucursales.Cod_Suc ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Ven      BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Zon            BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Status       IN ( " & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Tip            BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Cla            BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Tra      BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Mon      BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("		    AND ((" & lcParametro11Desde & " = 'Si' AND Cuentas_Cobrar.Mon_Sal > 0)")
            loComandoSeleccionar.AppendLine("			OR  (" & lcParametro11Desde & " <> 'Si' AND (Cuentas_Cobrar.Mon_Sal >= 0 or Cuentas_Cobrar.Mon_Sal < 0)))")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Rev      BETWEEN " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento & ", Cuentas_Cobrar.Fec_Ini ")


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCCobrar_ClientesSucursales", laDatosReporte)

            '-------------------------------------------------------------------------------------------------------
            ' Agregando el comentario al reporte completo
            '-------------------------------------------------------------------------------------------------------
            'If lcParametro12Desde.ToString = "Si" Then
            '    loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 1
            '    CType(loObjetoReporte.ReportDefinition.ReportObjects("Line3"), CrystalDecisions.CrystalReports.Engine.LineObject).Right = 0
            '    CType(loObjetoReporte.ReportDefinition.ReportObjects("Line3"), CrystalDecisions.CrystalReports.Engine.LineObject).Left = 0
            'Else
            '    loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("text1").Height = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("text14").Height = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Height = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("text1").Top = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("text14").Top = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Top = 0
            '    loObjetoReporte.ReportDefinition.Sections(3).Height = 200
            '    loObjetoReporte.ReportDefinition.Sections("DetailSection2").Height = 10
            'End If

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCCobrar_ClientesSucursales.ReportSource = loObjetoReporte

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
' JJD: 07/05/10: Programacion inicial
'-------------------------------------------------------------------------------------------'
