'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEfectividad_Pagos"
'-------------------------------------------------------------------------------------------'
Partial Class rEfectividad_Pagos
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
           
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" ROW_NUMBER() OVER(PARTITION BY Pagos.Cod_Pro ORDER BY Pagos.Cod_Pro, " & lcOrdenamiento & " ) AS Fila,")
            If lcParametro8Desde.ToString = "Si" Then
                loComandoSeleccionar.AppendLine(" 'Si' As Resumen,")
            Else
                loComandoSeleccionar.AppendLine(" 'No' As Resumen,")
            End If
            loComandoSeleccionar.AppendLine(" 			Pagos.Cod_Pro,")
            loComandoSeleccionar.AppendLine(" 			Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine(" 			Pagos.Documento,")
            loComandoSeleccionar.AppendLine(" 			Renglones_Pagos.Cod_Tip,")
            loComandoSeleccionar.AppendLine(" 			Renglones_Pagos.Doc_Ori,")
            loComandoSeleccionar.AppendLine(" 			Renglones_Pagos.Fec_Ini AS Fec_Ini_Ren,")
            loComandoSeleccionar.AppendLine(" 			Pagos.Fec_Ini AS Fec_Ini_Pag,")
            loComandoSeleccionar.AppendLine(" 			Renglones_Pagos.Mon_Net,")
            loComandoSeleccionar.AppendLine(" 			DATEDIFF(d, Renglones_Pagos.Fec_Ini, Pagos.Fec_Ini) AS Dias")
            loComandoSeleccionar.AppendLine(" FROM		Pagos")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine(" JOIN Cuentas_Pagar ON Cuentas_Pagar.Documento = Renglones_Pagos.Doc_Ori AND Cuentas_Pagar.Cod_Tip = Renglones_Pagos.Cod_Tip")
            loComandoSeleccionar.AppendLine(" JOIN Proveedores ON Proveedores.Cod_Pro = Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine(" WHERE	")
            loComandoSeleccionar.AppendLine("           Pagos.Status <> 'Anulado'	")
            loComandoSeleccionar.AppendLine(" 			AND Pagos.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Pagos.Cod_Tip between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Pagos.Fec_Ini between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Pagos.Cod_Pro between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Pagos.Cod_Ven between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Mon between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Proveedores.Cod_Zon between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Tip between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Rev between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Suc between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY   Pagos.Cod_Pro, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEfectividad_Pagos", laDatosReporte)

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


            'If lcParametro8Desde.ToString = "Si" Then
            '    loObjetoReporte.ReportDefinition.ReportObjects("Documento1").Height = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("Documento1").Top = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("CodTip1").Height = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("CodTip1").Top = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("DocOri1").Height = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("DocOri1").Top = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("FecIniRen1").Height = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("FecIniRen1").Top = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("FecIniPag1").Height = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("FecIniPag1").Top = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("Dias1").Height = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("Dias1").Top = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("MonNet1").Height = 0
            '    loObjetoReporte.ReportDefinition.ReportObjects("MonNet1").Top = 0
            '    CType(loObjetoReporte.ReportDefinition.ReportObjects("Line1"), CrystalDecisions.CrystalReports.Engine.LineObject).Right = 0
            '    CType(loObjetoReporte.ReportDefinition.ReportObjects("Line1"), CrystalDecisions.CrystalReports.Engine.LineObject).Left = 0
            '    loObjetoReporte.ReportDefinition.Sections("Section3").Height = 10
            'End If

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrEfectividad_Pagos.ReportSource = loObjetoReporte

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
' MAT: 28/02/11: Eliminación del Parámetro Vendedor
'-------------------------------------------------------------------------------------------'

