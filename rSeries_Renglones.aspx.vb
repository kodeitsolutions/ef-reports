'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rSeries_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rSeries_Renglones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("  SELECT ")
            loComandoSeleccionar.AppendLine("  			Series.Cod_Ser, ")
            loComandoSeleccionar.AppendLine("  			Series.Nom_Ser, ")
            loComandoSeleccionar.AppendLine("  			Series.Tipo,  ")
            loComandoSeleccionar.AppendLine("  			CASE ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Status = 'A'  THEN 'Activo' ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Status = 'I'  THEN 'Inactivo' ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Status = 'S'  THEN 'Suspendido' ")
            loComandoSeleccionar.AppendLine("  			END AS Status, ")
            loComandoSeleccionar.AppendLine("  			Series.Resultado, ")
            loComandoSeleccionar.AppendLine("  			Series.Comentario, ")
            loComandoSeleccionar.AppendLine("  			Renglones_Series.Renglon, ")
            loComandoSeleccionar.AppendLine("  			CASE ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Tipo = 'Caracter' THEN CAST(Renglones_Series.Car_Ini  AS VARCHAR(20)) ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Tipo = 'Numerico' THEN CAST(Renglones_Series.Num_Ini  AS VARCHAR(20)) ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Tipo = 'Fecha' THEN CONVERT(nchar(30), Renglones_Series.Fec_Ini,103) ")
            loComandoSeleccionar.AppendLine("  			END AS Inicio,  ")
            loComandoSeleccionar.AppendLine("  			CASE ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Tipo = 'Caracter' THEN CAST(Renglones_Series.Car_Fin  AS VARCHAR(20)) ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Tipo = 'Numerico' THEN CAST(Renglones_Series.Num_Fin  AS VARCHAR(20)) ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Tipo = 'Fecha' THEN CONVERT(nchar(30), Renglones_Series.Fec_Fin,103)   ")
            loComandoSeleccionar.AppendLine("  			END AS Fin, ")
            loComandoSeleccionar.AppendLine("  			CASE ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Resultado = 'Numerico' THEN CAST(Renglones_Series.Val_Num  AS VARCHAR(20)) ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Resultado = 'Logico' THEN CAST(Renglones_Series.Val_Log AS VARCHAR(20)) ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Resultado = 'Caracter' THEN CAST(Renglones_Series.Val_Car AS VARCHAR(20)) ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Resultado = 'Fecha' THEN CONVERT(nchar(30), Renglones_Series.Val_Fec,103) ")
            loComandoSeleccionar.AppendLine("  				WHEN Series.Resultado = 'Memo' THEN CAST(Renglones_Series.Val_Mem AS VARCHAR(1000)) ")
            loComandoSeleccionar.AppendLine("  			END AS Valor ")
            loComandoSeleccionar.AppendLine("  FROM		Series ")
            loComandoSeleccionar.AppendLine("  JOIN Renglones_Series ON Renglones_Series.Cod_Ser = Series.Cod_Ser ")
            loComandoSeleccionar.AppendLine("WHERE ")
            loComandoSeleccionar.AppendLine(" 				Series.Cod_Ser between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Series.status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento & ", Renglones_Series.Renglon")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rSeries_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrSeries_Renglones.ReportSource = loObjetoReporte


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
' CMS:  15/07/09: Codigo inicial
'-------------------------------------------------------------------------------------------'