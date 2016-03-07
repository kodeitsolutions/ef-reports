'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEfectividad_Cobranzas"
'-------------------------------------------------------------------------------------------'
Partial Class rEfectividad_Cobranzas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load



        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" ROW_NUMBER() OVER(PARTITION BY Cobros.Cod_Cli ORDER BY Cobros.Cod_Cli, " & lcOrdenamiento & " ) AS Fila,")
            loComandoSeleccionar.AppendLine(" 			Cobros.Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 			Cobros.Documento,")
            loComandoSeleccionar.AppendLine(" 			Renglones_Cobros.Cod_Tip,")
            loComandoSeleccionar.AppendLine(" 			Renglones_Cobros.Doc_Ori,")
            loComandoSeleccionar.AppendLine(" 			Renglones_Cobros.Mon_Net,")
            loComandoSeleccionar.AppendLine(" 			Renglones_Cobros.Fec_Ini AS Fec_Ini_Ren,")
            loComandoSeleccionar.AppendLine(" 			Cobros.Fec_Ini AS Fec_Ini_Cob,")
            loComandoSeleccionar.AppendLine(" 			DATEDIFF(d, Renglones_Cobros.Fec_Ini, Cobros.Fec_Ini) AS Dias")
            loComandoSeleccionar.AppendLine(" FROM		Cobros")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Cobros ON Renglones_Cobros.Documento = Cobros.Documento")
            'loComandoSeleccionar.AppendLine(" JOIN Cuentas_Cobrar ON Cuentas_Cobrar.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine(" JOIN Cuentas_Cobrar ON Cuentas_Cobrar.Documento = Renglones_Cobros.Doc_Ori AND Cuentas_Cobrar.Cod_Tip = Renglones_Cobros.Cod_Tip")
            loComandoSeleccionar.AppendLine(" JOIN Clientes ON Clientes.Cod_Cli = Cobros.Cod_Cli")
            loComandoSeleccionar.AppendLine(" WHERE	")
            loComandoSeleccionar.AppendLine("           Cobros.Status <> 'Anulado'	")
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Cobros.Cod_Tip between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Fec_Ini between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Cli between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Ven between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Mon between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Zon between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Tip between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Rev between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Suc between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Ven between " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY   Cobros.Cod_Cli, " & lcOrdenamiento)
            
            
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEfectividad_Cobranzas", laDatosReporte)

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


            If lcParametro8Desde.ToString = "Si" Then
                loObjetoReporte.ReportDefinition.ReportObjects("Documento1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Documento1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("CodTip1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("CodTip1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("DocOri1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("DocOri1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("FecIniRen1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("FecIniRen1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("FecIniCob1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("FecIniCob1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Dias1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Dias1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("MonNet1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("MonNet1").Top = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line1"), CrystalDecisions.CrystalReports.Engine.LineObject).Right = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line1"), CrystalDecisions.CrystalReports.Engine.LineObject).Left = 0
                loObjetoReporte.ReportDefinition.Sections("Section3").Height = 10
            End If

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrEfectividad_Cobranzas.ReportSource = loObjetoReporte

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
' CMS: 22/09/09: Programacion inicial
'-------------------------------------------------------------------------------------------'

