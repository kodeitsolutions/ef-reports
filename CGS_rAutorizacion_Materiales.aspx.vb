Imports System.Data
Partial Class CGS_rAutorizacion_Materiales
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
        Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))

        Try

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("DECLARE @lcCod_Pro AS VARCHAR(10) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha AS DATETIME = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt1 AS VARCHAR(8) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lnCanArt1 AS DECIMAL(28,3) = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt2 AS VARCHAR(8) = " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lnCanArt2 AS DECIMAL(28,3) = " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt3 AS VARCHAR(8) = " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lnCanArt3 AS DECIMAL(28,3) = " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt4 AS VARCHAR(8) = " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lnCanArt4 AS DECIMAL(28,3) = " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Proveedores.Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine("		@ldFecha	AS Fecha,")
            loComandoSeleccionar.AppendLine("		@lcCodArt1	AS Cod_Art1,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt1 <> '' THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt1) ELSE '' END AS Nom_Art1,	")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt1 <> '' THEN (SELECT Cod_Uni1 FROM Articulos WHERE Cod_Art = @lcCodArt1) ELSE '' END AS Cod_Uni1,")
            loComandoSeleccionar.AppendLine("		@lnCanArt1	AS Can_Art1,")
            loComandoSeleccionar.AppendLine("		@lcCodArt2	AS Cod_Art2,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt2 <> '' THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt2) ELSE '' END AS Nom_Art2,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt2 <> '' THEN (SELECT Cod_Uni1 FROM Articulos WHERE Cod_Art = @lcCodArt2) ELSE '' END AS Cod_Uni2,")
            loComandoSeleccionar.AppendLine("		@lnCanArt2	AS Can_Art2,")
            loComandoSeleccionar.AppendLine("		@lcCodArt3	AS Cod_Art3,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt3 <> '' THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt3) ELSE '' END AS Nom_Art3,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt3 <> '' THEN (SELECT Cod_Uni1 FROM Articulos WHERE Cod_Art = @lcCodArt3) ELSE '' END AS Cod_Uni3,")
            loComandoSeleccionar.AppendLine("		@lnCanArt3	AS Can_Art3,")
            loComandoSeleccionar.AppendLine("		@lcCodArt4	AS Cod_Art4,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt4 <> '' THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt4) ELSE '' END AS Nom_Art4,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt4 <> '' THEN (SELECT Cod_Uni1 FROM Articulos WHERE Cod_Art = @lcCodArt4) ELSE '' END AS Cod_Uni4,")
            loComandoSeleccionar.AppendLine("		@lnCanArt4	AS Can_Art4,")
            loComandoSeleccionar.AppendLine("		CAST(YEAR(@ldFecha) AS VARCHAR(4))		AS Anio,")
            loComandoSeleccionar.AppendLine("		CASE WHEN MONTH(@ldFecha) < 10 THEN '0' + CAST(MONTH(@ldFecha) AS VARCHAR(2)) ELSE CAST(MONTH(@ldFecha) AS VARCHAR(2)) END AS Mes,")
            loComandoSeleccionar.AppendLine("		CASE WHEN DAY(@ldFecha) < 10 THEN '0' + CAST(DAY(@ldFecha) AS VARCHAR(2)) ELSE CAST(DAY(@ldFecha) AS VARCHAR(2)) END AS Dia")
            loComandoSeleccionar.AppendLine("FROM Proveedores")
            loComandoSeleccionar.AppendLine("WHERE Proveedores.Cod_Pro = @lcCod_Pro")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rAutorizacion_Materiales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rAutorizacion_Materiales.ReportSource = loObjetoReporte

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
' JJD: 08/11/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 10/09/09: Se ajusto el nombre del articulo para los casos de aquellos articulos gen.
'-------------------------------------------------------------------------------------------'
' JJD: 09/01/10: Se cambio para que leyera datos de genericos de la Cotizacion cuando aplique
'-------------------------------------------------------------------------------------------'
' CMS: 17/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT: 02/09/11: Adición de Comentario en Renglones
'-------------------------------------------------------------------------------------------'
