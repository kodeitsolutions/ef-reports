'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rArticulos_Stock"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rArticulos_Stock
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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            'Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)

            'Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            'Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            'Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            'Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            'Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            'Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            'Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            'Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            'Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            'Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)
            'Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)

            'Dim lcExisiencia As String = ""

            'Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()



            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(8) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(8) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Desde AS VARCHAR(10) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Hasta AS VARCHAR(10) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Desde AS VARCHAR(2) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Hasta AS VARCHAR(2) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Desde AS VARCHAR(2) = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Hasta AS VARCHAR(2) = " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Almacenes.Cod_Alm				AS Cod_Alm,")
            loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm				AS Nom_Alm,")
            loComandoSeleccionar.AppendLine("		Articulos.Cod_Art				AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("    	Articulos.Nom_Art				AS Nom_Art, ")
            loComandoSeleccionar.AppendLine("    	Articulos.Cod_Uni1				AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("    	Articulos.Cod_Dep				AS Cod_Dep, ")
            loComandoSeleccionar.AppendLine("    	Articulos.Cod_Sec				AS Cod_Sec,		")
            loComandoSeleccionar.AppendLine("		Articulos.Exi_Act1				AS Exi_Art,")
            loComandoSeleccionar.AppendLine("		Renglones_Almacenes.Exi_Act1	AS Exi_Alm, ")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Art_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Art_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodDep_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Dep FROM Departamentos WHERE Cod_Dep = @lcCodDep_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Dep_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodDep_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Dep FROM Departamentos WHERE Cod_Dep = @lcCodDep_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Dep_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSec_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN CASE WHEN @lcCodDep_Desde = '' AND @lcCodDep_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("		               THEN (SELECT TOP 1 Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Desde)")
            loComandoSeleccionar.AppendLine("		        	   ELSE (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Desde AND Cod_Dep = @lcCodDep_Desde) END")
            loComandoSeleccionar.AppendLine("            ELSE '' END				AS Sec_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSec_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN CASE WHEN @lcCodDep_Desde = '' AND @lcCodDep_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("		               THEN (SELECT TOP 1 Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Hasta)")
            loComandoSeleccionar.AppendLine("		        	   ELSE (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Hasta AND Cod_Dep = @lcCodDep_Hasta) END")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Sec_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Alm_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm  FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Alm_Hasta")
            loComandoSeleccionar.AppendLine("FROM	Articulos")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Almacenes ON Renglones_Almacenes.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN Almacenes ON Renglones_Almacenes.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("WHERE Articulos.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("	AND Almacenes.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("	AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            loComandoSeleccionar.AppendLine("	AND Renglones_Almacenes.Exi_Act1 > 0")
            loComandoSeleccionar.AppendLine("   AND Articulos.Status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("ORDER BY Almacenes.Cod_Alm, Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("")
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rArticulos_Stock", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rArticulos_Stock.ReportSource = loObjetoReporte

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