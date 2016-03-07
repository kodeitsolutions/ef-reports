'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDocumentosCXC_aComentario"
'-------------------------------------------------------------------------------------------'
Partial Class rDocumentosCXC_aComentario
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
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosIniciales(6)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))

            Dim lcParametro13Hasta As String = cusAplicacion.goReportes.paParametrosFinales(13)
            Dim Fecha As String

            If lcParametro10Hasta = "Vencimiento" Then
                Fecha = "Fec_Fin"
            Else
                Fecha = "Fec_Ini"
            End If

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Documento,")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Tip,")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Fec_Ini,")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Fec_Fin,")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Ven,")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Mon_Imp1,")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Mon_Net,")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Mon_Sal,")
            loComandoSeleccionar.AppendLine(" 			UPPER(SubString(Cuentas_Cobrar.Comentario, 1,4)) AS Comentario")
            loComandoSeleccionar.AppendLine(" FROM		Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine(" JOIN Clientes ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Tip between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Fec_Ini between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Cli between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Ven between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Status IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Mon between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Zon between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Tip between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Rev between " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Suc between " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro12Hasta)

            Select Case lcParametro6Desde
                Case "Con Saldo"
                    loComandoSeleccionar.AppendLine("    	    AND Cuentas_Cobrar.Mon_Sal <> 0 AND Cuentas_Cobrar.Status <> 'Anulado'")
                Case "Cancelados"
                    loComandoSeleccionar.AppendLine("    	    AND Cuentas_Cobrar.Mon_Sal = 0 AND Cuentas_Cobrar.Status <> 'Anulado'")
                Case "Por Vencer"
                    loComandoSeleccionar.AppendLine("    	    AND Cuentas_Cobrar." & Fecha & " < " & lcParametro2Desde & " AND Cuentas_Cobrar.Status <> 'Anulado' AND Cuentas_Cobrar.Mon_Sal <> 0 ")
                Case "Vencidos"
                    loComandoSeleccionar.AppendLine("    	    AND Cuentas_Cobrar." & Fecha & " > " & lcParametro2Desde & " AND Cuentas_Cobrar.Status <> 'Anulado' AND Cuentas_Cobrar.Mon_Sal <> 0  ")
                Case "0-30"
                    loComandoSeleccionar.AppendLine("    	    AND DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro2Desde & ") >= 0 AND DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro2Desde & ") <= 30 AND Cuentas_Cobrar.Status <> 'Anulado' AND Cuentas_Cobrar.Mon_Sal <> 0 ")
                Case "31-60"
                    loComandoSeleccionar.AppendLine("    	    AND DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro2Desde & ") >= 31 AND DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro2Desde & ") <= 60 AND Cuentas_Cobrar.Status <> 'Anulado' AND Cuentas_Cobrar.Mon_Sal <> 0 ")
                Case "61-90"
                    loComandoSeleccionar.AppendLine("    	    AND DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro2Desde & ") >= 61 AND DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro2Desde & ") <= 90 AND Cuentas_Cobrar.Status <> 'Anulado' AND Cuentas_Cobrar.Mon_Sal <> 0 ")
                Case "91-120"
                    loComandoSeleccionar.AppendLine("    	    AND DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro2Desde & ") >= 91 AND DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro2Desde & ") <= 120 AND Cuentas_Cobrar.Status <> 'Anulado' AND Cuentas_Cobrar.Mon_Sal <> 0 ")
                Case "+120"
                    loComandoSeleccionar.AppendLine("    	    AND DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro2Desde & ") >= 121 AND Cuentas_Cobrar.Status <> 'Anulado' AND Cuentas_Cobrar.Mon_Sal <> 0 ")
            End Select

            loComandoSeleccionar.AppendLine("ORDER BY   UPPER(SubString(Cuentas_Cobrar.Comentario, 1,4)), " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDocumentosCXC_aComentario", laDatosReporte)

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
                loObjetoReporte.ReportDefinition.ReportObjects("FecIni1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("FecIni1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("FecFin1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("FecFin1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("CodVen1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("CodVen1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("MonImp11").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("MonImp11").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("MonNet1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("MonNet1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("MonSal1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("MonSal1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("CodCli1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("CodCli1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("NomCli1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("NomCli1").Top = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects("NomCli1"), CrystalDecisions.CrystalReports.Engine.FieldObject).ObjectFormat.EnableCanGrow = False
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line1"), CrystalDecisions.CrystalReports.Engine.LineObject).Right = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line1"), CrystalDecisions.CrystalReports.Engine.LineObject).Left = 0
                loObjetoReporte.ReportDefinition.Sections("Section3").Height = 10
            End If

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrDocumentosCXC_aComentario.ReportSource = loObjetoReporte

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
' CMS: 23/09/09: Programacion inicial
'-------------------------------------------------------------------------------------------'

