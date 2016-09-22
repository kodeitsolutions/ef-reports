'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rLista_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rLista_Articulos
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


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcArtDesde VARCHAR(8) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcArtHasta VARCHAR(8) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcDepDesde VARCHAR(2) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcDepHasta VARCHAR(2) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcSecDesde VARCHAR(2) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcSecHasta VARCHAR(2) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Articulos.Cod_Art		AS Codigo_Articulo, ")
            loComandoSeleccionar.AppendLine("        Articulos.Nom_Art		AS Descripcion, ")
            loComandoSeleccionar.AppendLine("        Departamentos.Nom_Dep	AS Departamento, ")
            loComandoSeleccionar.AppendLine("        Secciones.Nom_Sec		AS Seccion,	")
            loComandoSeleccionar.AppendLine("        ISNULL(CAST(Articulos.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')		AS CC_Art,")
            loComandoSeleccionar.AppendLine("		ISNULL((SELECT Nom_Cue FROM Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("				WHERE Cod_Cue = ISNULL(CAST(Articulos.Contable AS XML)")
            loComandoSeleccionar.AppendLine("				.value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')),'')									AS Nom_CC_Art,")
            loComandoSeleccionar.AppendLine("		ISNULL(CAST(Secciones.Contable AS XML) .value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')		AS CC_Sec,")
            loComandoSeleccionar.AppendLine("		ISNULL((SELECT Nom_Cue FROM Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("				WHERE Cod_Cue = ISNULL(CAST(Secciones.Contable AS XML)")
            loComandoSeleccionar.AppendLine("				.value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')),'')									AS Nom_CC_Sec")
            loComandoSeleccionar.AppendLine("FROM	Articulos ")
            loComandoSeleccionar.AppendLine("    JOIN Departamentos ON Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("    JOIN Secciones ON Articulos.Cod_Sec = Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine("        AND Secciones.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("WHERE	Articulos.Cod_Art BETWEEN @lcArtDesde AND @lcArtHasta")
            loComandoSeleccionar.AppendLine("        And Departamentos.Cod_Dep BETWEEN @lcDepDesde AND @lcDepHasta")
            loComandoSeleccionar.AppendLine("        And Secciones.Cod_Sec BETWEEN @lcSecDesde AND @lcSecHasta")
            loComandoSeleccionar.AppendLine("ORDER BY Departamentos.Nom_Dep, Secciones.Nom_Sec ")
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rLista_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvTRF_rLista_Articulos.ReportSource = loObjetoReporte

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
' MJP: 16/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GS:  14/03/16: Cambio a Listado de Artículos.
'-------------------------------------------------------------------------------------------'

