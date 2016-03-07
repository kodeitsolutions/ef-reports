'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTablas"
'-------------------------------------------------------------------------------------------'
Partial Class rTablas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            
            'Creando Tabla Temporal
            loComandoSeleccionar.AppendLine(" CREATE TABLE #tblRESULTADOS")
            loComandoSeleccionar.AppendLine(" 			(")
            loComandoSeleccionar.AppendLine(" 				[name]			nvarchar(100),")
            loComandoSeleccionar.AppendLine(" 				[rows]			int,")
            loComandoSeleccionar.AppendLine(" 				[reserved]		varchar(50),")
            loComandoSeleccionar.AppendLine(" 				[data]			varchar(50),")
            loComandoSeleccionar.AppendLine(" 				[index_size]	varchar(50),")
            loComandoSeleccionar.AppendLine(" 				[unused]		varchar(50),")
            loComandoSeleccionar.AppendLine(" 			)")

            'Llenando la Tabbla Temporal
            loComandoSeleccionar.AppendLine(" EXEC sp_MSforeachtable @command1=")
            loComandoSeleccionar.AppendLine(" 'INSERT INTO #tblRESULTADOS")
            loComandoSeleccionar.AppendLine(" ([name],[rows],[reserved],[data],[index_size],[unused])")
            loComandoSeleccionar.AppendLine(" EXEC sp_spaceused ''?'''")

            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine("  			name AS Tabla,")
            loComandoSeleccionar.AppendLine("  			rows As Filas,")
            loComandoSeleccionar.AppendLine("  			CAST(left(reserved, LEN(reserved)-3) AS INT) AS Espacio_Reservado,")
            loComandoSeleccionar.AppendLine("  			CAST(left(data, LEN(data)-3) AS INT) AS Espacio_Usado_Dato,")
            loComandoSeleccionar.AppendLine("  			CAST(left(index_size, LEN(index_size)-3) AS INT) AS Espacio_Usado_Indice,")
            loComandoSeleccionar.AppendLine("  			CAST(left(unused, LEN(unused)-3) AS INT) AS Espacio_No_Utilizado")
            loComandoSeleccionar.AppendLine(" FROM #tblRESULTADOS")
            loComandoSeleccionar.AppendLine(" WHERE     ")
            loComandoSeleccionar.AppendLine(" 			name BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY      " & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTablas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTablas.ReportSource = loObjetoReporte

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
' CMS: 19/08/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' MAT: 09/08/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
