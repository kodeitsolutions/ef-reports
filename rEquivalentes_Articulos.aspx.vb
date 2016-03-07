'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEquivalentes_Articulos"
'-------------------------------------------------------------------------------------------'

Partial Class rEquivalentes_Articulos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcComandoSeleccionar As New StringBuilder()

            lcComandoSeleccionar.AppendLine(" SELECT")
            lcComandoSeleccionar.AppendLine(" 			Cod_Art,")
            lcComandoSeleccionar.AppendLine(" 			Nom_Art,")
            lcComandoSeleccionar.AppendLine(" 			Cod_Uni1 ")
            lcComandoSeleccionar.AppendLine(" INTO #tempARTICULOPRINCIPAL")
            lcComandoSeleccionar.AppendLine(" FROM Articulos")
            lcComandoSeleccionar.AppendLine(" WHERE     Cod_Art        Between " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Status     IN ( " & lcParametro1Desde & ")")
            
            lcComandoSeleccionar.AppendLine(" SELECT")
            lcComandoSeleccionar.AppendLine(" 			Equivalentes_Articulos.Cod_Art,")
            lcComandoSeleccionar.AppendLine(" 			Equivalentes_Articulos.Cod_Equ AS Cod_Art_Equivalente,")
            lcComandoSeleccionar.AppendLine(" 			Equivalentes_Articulos.Nom_Equ AS Nom_Art_Equivalente,")
            lcComandoSeleccionar.AppendLine(" 			Articulos.Cod_Uni1  AS Cod_Uni1_Equivalente")
            lcComandoSeleccionar.AppendLine(" INTO #tempARTICULOSEQUIVALENTE")
            lcComandoSeleccionar.AppendLine(" FROM Equivalentes_Articulos")
            lcComandoSeleccionar.AppendLine(" JOIN Articulos ON Articulos.Cod_Art = Equivalentes_Articulos.Cod_Art")
            lcComandoSeleccionar.AppendLine(" WHERE     Equivalentes_Articulos.Cod_Art        Between " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Status     IN ( " & lcParametro1Desde & ")")

            lcComandoSeleccionar.AppendLine(" SELECT * FROM ")
            lcComandoSeleccionar.AppendLine(" #tempARTICULOPRINCIPAL, #tempARTICULOSEQUIVALENTE")
            lcComandoSeleccionar.AppendLine(" WHERE #tempARTICULOPRINCIPAL.cod_art = #tempARTICULOSEQUIVALENTE.Cod_Art")
            lcComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEquivalentes_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrEquivalentes_Articulos.ReportSource = loObjetoReporte

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
' CMS: 24/08/09: Codigo inicial
'-------------------------------------------------------------------------------------------'