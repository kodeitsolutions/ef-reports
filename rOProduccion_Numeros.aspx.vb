'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOProduccion_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rOProduccion_Numeros
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = cusAplicacion.goReportes.paParametrosIniciales(4)

            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT         ")
            loComandoSeleccionar.AppendLine("			Documento,    ")
            loComandoSeleccionar.AppendLine("			Fec_Ini,      ")
            loComandoSeleccionar.AppendLine("			Cod_Art,      ")
            loComandoSeleccionar.AppendLine("			Nom_Art,      ")
            loComandoSeleccionar.AppendLine("			Status,       ")
            loComandoSeleccionar.AppendLine("			Modelo,       ")
            loComandoSeleccionar.AppendLine("			Cod_Uni,      ")
            loComandoSeleccionar.AppendLine("			Can_Art1,     ")
            loComandoSeleccionar.AppendLine("			Comentario    ")
            loComandoSeleccionar.AppendLine("FROM		Ordenes_Produccion   ")
            loComandoSeleccionar.AppendLine("WHERE   ")
            loComandoSeleccionar.AppendLine(" 			 Ordenes_Produccion.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			 AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			 AND Ordenes_Produccion.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			 AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			 AND Ordenes_Produccion.Cod_Art between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			 AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			 AND Ordenes_Produccion.status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("ORDER BY    " & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOProduccion_Numeros", laDatosReporte)


            If lcParametro4Desde.ToString = "Si" Then
                loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 1
                CType(loObjetoReporte.ReportDefinition.ReportObjects(26), CrystalDecisions.CrystalReports.Engine.LineObject).Right = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects(26), CrystalDecisions.CrystalReports.Engine.LineObject).Left = 0
            Else
                loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text11").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text12").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text11").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text12").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Top = 0
                loObjetoReporte.ReportDefinition.Sections(2).Height = 400
            End If

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOProduccion_Numeros.ReportSource = loObjetoReporte


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
' CMS:  06/07/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
