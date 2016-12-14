'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rStock_LoteAlmacen"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rStock_LoteAlmacen
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcAlm_Desde	AS VARCHAR(15) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcAlm_Hasta	AS VARCHAR(15) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcArt_Desde	AS VARCHAR(8) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcArt_Hasta	AS VARCHAR(8) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcDep_Desde	AS VARCHAR(2) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcDep_Hasta	AS VARCHAR(2) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcSec_Desde	AS VARCHAR(2) = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcSec_Hasta	AS VARCHAR(2) = " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT DISTINCT")
            loComandoSeleccionar.AppendLine("       Almacenes.Nom_Alm							AS Nom_Alm,")
            loComandoSeleccionar.AppendLine("		Renglones_Lotes.Cod_Alm						AS Cod_Alm,	")
            loComandoSeleccionar.AppendLine("		Renglones_Lotes.Cod_Art						AS Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Nom_Art							AS Nom_Art,")
            loComandoSeleccionar.AppendLine("		CASE WHEN LEN(Articulos.Nom_Art) > 50")
            loComandoSeleccionar.AppendLine("			 THEN CONCAT(SUBSTRING(Articulos.Nom_Art, 0, 50), '...')")
            loComandoSeleccionar.AppendLine("			 ELSE Articulos.Nom_Art")
            loComandoSeleccionar.AppendLine("		END											AS Descripcion, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Cod_Uni1							AS Cod_Uni,")
            loComandoSeleccionar.AppendLine("		Departamentos.Nom_Dep						AS Nom_Dep,")
            loComandoSeleccionar.AppendLine("		Secciones.Nom_Sec							AS Nom_Sec,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Lotes.Exi_Act1					AS Existencia, ")
            loComandoSeleccionar.AppendLine(" 		Renglones_Lotes.Cod_Lot						AS Lote,")
            'loComandoSeleccionar.AppendLine("		COALESCE(Mediciones.Documento, '')			AS Medicion,")
            loComandoSeleccionar.AppendLine("		COALESCE(Renglones_Mediciones.Res_Num, 0)	AS Piezas,")
            loComandoSeleccionar.AppendLine("		COALESCE(Mediciones.Origen, '')				AS Origen 		")
            loComandoSeleccionar.AppendLine("FROM Lotes")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Lotes ON Renglones_Lotes.Cod_Lot = Lotes.Cod_Lot ")
            loComandoSeleccionar.AppendLine("	JOIN Almacenes ON Renglones_Lotes.Cod_Alm = Almacenes.Cod_Alm ")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Renglones_Lotes.Cod_Art ")
            loComandoSeleccionar.AppendLine("	JOIN Departamentos ON Articulos.Cod_Dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine("	JOIN Secciones ON Secciones.Cod_Sec = Articulos.Cod_Sec")
            loComandoSeleccionar.AppendLine("	    AND Secciones.Cod_Dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen IN ('Ajustes_Inventarios', 'Recepciones', 'Traslados', 'Encabezados')")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Renglones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones ON Mediciones.Documento = Renglones_Mediciones.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_Mediciones.Cod_Var IN ('AINV-NPIEZ', 'NREC-NPIEZ')")
            loComandoSeleccionar.AppendLine("WHERE Renglones_Lotes.Exi_Act1 > 0")
            loComandoSeleccionar.AppendLine("	AND Almacenes.Cod_Alm BETWEEN @lcAlm_Desde AND @lcAlm_Hasta")
            loComandoSeleccionar.AppendLine("	AND Articulos.Cod_Art BETWEEN @lcArt_Desde AND @lcArt_Hasta")
            loComandoSeleccionar.AppendLine("	AND Articulos.Cod_Dep BETWEEN @lcDep_Desde AND @lcDep_Hasta")
            loComandoSeleccionar.AppendLine("	AND Articulos.Cod_Sec BETWEEN @lcSec_Desde AND @lcSec_Hasta")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rStock_LoteAlmacen", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rStock_LoteAlmacen.ReportSource = loObjetoReporte

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