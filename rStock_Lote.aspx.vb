'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rStock_Lote"
'-------------------------------------------------------------------------------------------'
Partial Class rStock_Lote
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
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosIniciales(7)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine("             Lotes.Cod_Art, ")
            loComandoSeleccionar.AppendLine("             Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("             Articulos.Cod_uni1, ")
            loComandoSeleccionar.AppendLine("             Lotes.Cod_Lot, ")
            loComandoSeleccionar.AppendLine("             Lotes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("             Lotes.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("             CONVERT(nchar(30), Lotes.Fec_Ini,112) AS Fec_Ini2, ")
            loComandoSeleccionar.AppendLine("             CONVERT(nchar(30), Lotes.Fec_Fin,112) AS Fec_Fin2, ")
            loComandoSeleccionar.AppendLine("             Lotes.Exi_Act1 ")
            loComandoSeleccionar.AppendLine(" FROM Lotes ")
            loComandoSeleccionar.AppendLine(" JOIN Articulos ON Articulos.Cod_Art = Lotes.Cod_Art ")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Usa_Lot = 1 ")

            loComandoSeleccionar.AppendLine("               And Lotes.Cod_Lot       Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Art   Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Dep   Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Sec   Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Tip   Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Cla   Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Mar   Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro6Hasta)
            
            Select Case lcParametro7Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("")
                Case "Mayor"
                    loComandoSeleccionar.AppendLine("               And Lotes.Exi_Act1 > 0")
                Case "Menor"
                    loComandoSeleccionar.AppendLine("               And Lotes.Exi_Act1 < 0")
                Case "Igual"
                    loComandoSeleccionar.AppendLine("               And Lotes.Exi_Act1 = 0")
            End Select

            loComandoSeleccionar.AppendLine("ORDER BY  Lotes.Cod_Art, " & lcOrdenamiento)
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rStock_Lote", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrStock_Lote.ReportSource = loObjetoReporte

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
' CMS: 28/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'