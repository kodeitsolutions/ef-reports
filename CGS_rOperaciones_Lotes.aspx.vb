'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rOperaciones_Lotes"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rOperaciones_Lotes
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


            Dim loComandoSeleccionar As New StringBuilder()



            loComandoSeleccionar.AppendLine("DECLARE @lcLote_Desde	 AS VARCHAR(30) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcLote_Hasta	 AS VARCHAR(30) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(8) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(8) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Desde AS VARCHAR(10) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Hasta AS VARCHAR(10) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT cod_alm, cod_art, cod_lot, cantidad,num_doc, renglon, tip_doc, tip_ope")
            loComandoSeleccionar.AppendLine("FROM operaciones_lotes")
            loComandoSeleccionar.AppendLine("WHERE Cod_Lot BETWEEN @lcLote_Desde AND @lcLote_Hasta")
            loComandoSeleccionar.AppendLine("	AND cod_art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("	AND cod_alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rOperaciones_Lotes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rOperaciones_Lotes.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' MVP: 09/07/08: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' MVP: 01/08/08: Cambios para multi idioma, mensaje de error y clase padre.					'
'-------------------------------------------------------------------------------------------'
' JFP: 09/10/08: Limitacion de la longitud del nombre y ajustes de forma de presentacion	'
'-------------------------------------------------------------------------------------------'
' JJD: 02/02/09: Inclusion del Status. Normalizacion del Codigo								'
'-------------------------------------------------------------------------------------------'
' CMS: 05/05/09: Ordenamiento																'
'-------------------------------------------------------------------------------------------'
' CMS: 11/08/09: Verificacion de registros y se agregaro el filtro_ Ubicación				'
'-------------------------------------------------------------------------------------------'
' RJG: 19/08/09: Eliminación de uniones adicionales (departamentos, secciones....)			'
'-------------------------------------------------------------------------------------------'
' CMS: 21/09/09: Se agregaro los filtros tipo de existencia, nivel de existencia y proveedor'
'-------------------------------------------------------------------------------------------'