Imports System.Data
Partial Class CGS_rResumenCostos_ArtProd

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try

            lcComandoSeleccionar.AppendLine("DECLARE @ldFechaIni		AS DATETIME = " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @ldFechaFin		AS DATETIME = " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcProduccionIni	AS VARCHAR(10) = " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcProduccionFin	AS VARCHAR(10) = " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcArticuloIni     AS VARCHAR(8) = " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcArticuloFin     AS VARCHAR(8) = " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcLoteIni		    AS VARCHAR(15)= " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcLoteFin		    AS VARCHAR(15)= " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT Formulas.Documento          AS Documento,")
            lcComandoSeleccionar.AppendLine("       Formulas.Fec_Ini            AS Fecha,")
            lcComandoSeleccionar.AppendLine("       Formulas.Cod_Art            AS Cod_Art,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN LEN(Articulos.Nom_Art) > 50")
            lcComandoSeleccionar.AppendLine("			 THEN CONCAT(SUBSTRING(Articulos.Nom_Art, 0,50), '...')")
            lcComandoSeleccionar.AppendLine("			 ELSE Articulos.Nom_Art")
            lcComandoSeleccionar.AppendLine("		END					        AS Nom_Art,")
            lcComandoSeleccionar.AppendLine("       Formulas.Referencia         AS Referencia,")
            lcComandoSeleccionar.AppendLine("       Formulas.Cos_Otr1           AS Costo_Act,")
            lcComandoSeleccionar.AppendLine("       Formulas.Cos_Otr2           AS Costo_Std,")
            lcComandoSeleccionar.AppendLine("       Formulas.Caracter1          AS Lote,")
            lcComandoSeleccionar.AppendLine("       Formulas.Numerico1          AS Total_Act,")
            lcComandoSeleccionar.AppendLine("       Formulas.Numerico2          AS Total_Est,")
            lcComandoSeleccionar.AppendLine("       Formulas.Can_Ren            AS Obtenido    ")
            lcComandoSeleccionar.AppendLine("FROM   Formulas ")
            lcComandoSeleccionar.AppendLine("   JOIN Articulos ON Formulas.Cod_Art = Articulos.Cod_Art")
            lcComandoSeleccionar.AppendLine("	JOIN Proyectos ON Formulas.Referencia = Proyectos.Cod_Pro")
            lcComandoSeleccionar.AppendLine("WHERE Formulas.Origen = ''")
            lcComandoSeleccionar.AppendLine("	AND Proyectos.Fec_Pro BETWEEN @ldFechaIni AND @ldFechaFin")
            lcComandoSeleccionar.AppendLine("	AND Formulas.Referencia BETWEEN @lcProduccionIni AND @lcProduccionFin")
            lcComandoSeleccionar.AppendLine("	AND Formulas.Cod_Art BETWEEN @lcArticuloIni AND @lcArticuloFin")
            lcComandoSeleccionar.AppendLine("	AND Formulas.Caracter1 BETWEEN @lcLoteIni AND @lcLoteFin")

            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rResumenCostos_ArtProd", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rResumenCostos_ArtProd.ReportSource = loObjetoReporte

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
' JJD: 06/12/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' CMS: 22/06/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
