'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rTAlmacenes_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rTAlmacenes_Renglones
    Inherits vis2Formularios.frmReporte

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
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            'Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            'Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcDoc_Desde AS VARCHAR(10) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcDoc_Hasta AS VARCHAR(10) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(8) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(8) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcAlmOri_Desde AS VARCHAR(15) = " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcAlmOri_Hasta AS VARCHAR(15) = " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcAlmDes_Desde AS VARCHAR(15) = " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcAlmDes_Hasta AS VARCHAR(15) = " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT DISTINCT")
            loComandoSeleccionar.AppendLine("		Traslados.Documento                         AS Documento, ")
            loComandoSeleccionar.AppendLine("		Traslados.Fec_Ini                           AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("		Traslados.Alm_Ori                           AS Alm_Ori, ")
            loComandoSeleccionar.AppendLine("		Origen.Nom_Alm                              AS NomAlm_Ori,")
            loComandoSeleccionar.AppendLine("		Traslados.Alm_Des                           AS Alm_Des, ")
            loComandoSeleccionar.AppendLine("		Destino.Nom_Alm                             AS NomAlm_Des,")
            loComandoSeleccionar.AppendLine("		Traslados.Status                            AS Status, ")
            loComandoSeleccionar.AppendLine("		Traslados.Comentario                        AS Comentario,")
            loComandoSeleccionar.AppendLine("		Renglones_Traslados.Renglon                 AS Renglon, ")
            loComandoSeleccionar.AppendLine("		Renglones_Traslados.Cod_Art                 AS Cod_Art,")
            loComandoSeleccionar.AppendLine("		Renglones_Traslados.Can_Art1                AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("		Renglones_Traslados.Cod_Uni                 AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("		Articulos.Nom_Art                           AS Nom_Art,")
            loComandoSeleccionar.AppendLine("		COALESCE(Operaciones_Lotes.Cod_Lot ,'')		AS Cod_Lot,")
            loComandoSeleccionar.AppendLine("		COALESCE(Operaciones_Lotes.Cantidad,0)		AS Can_Lot,")
            loComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num, 0)					AS Piezas,")
            loComandoSeleccionar.AppendLine("		COALESCE(Desperdicio.Res_Num, 0)			AS Porc_Desperdicio,")
            loComandoSeleccionar.AppendLine("		COALESCE((Desperdicio.Res_Num*Operaciones_Lotes.Cantidad)/100, ")
            loComandoSeleccionar.AppendLine("		(Desperdicio.Res_Num*Renglones_Traslados.Can_Art1)/100, 0)	AS Cant_Desperdicio,")
            loComandoSeleccionar.AppendLine("		COALESCE(Longitud.Res_Num, 0)				AS Longitud,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcDoc_Desde <> '' ")
            loComandoSeleccionar.AppendLine("			 THEN CONCAT(@lcDoc_Desde, ' - ', @lcDoc_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE 'N/E'")
            loComandoSeleccionar.AppendLine("		END										    AS Docs,")
            loComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Art_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Art_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcAlmOri_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes WHERE Cod_Alm = @lcAlmOri_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				            AS AlmOri_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcAlmOri_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes WHERE Cod_Alm = @lcAlmOri_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				            AS AlmOri_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcAlmDes_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes WHERE Cod_Alm = @lcAlmDes_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				            AS AlmDes_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcAlmDes_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes WHERE Cod_Alm = @lcAlmDes_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				            AS AlmDes_Hasta")
            loComandoSeleccionar.AppendLine("FROM Traslados")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Traslados ON Traslados.Documento = Renglones_Traslados.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_Traslados.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN Almacenes AS Origen ON Traslados.Alm_Ori = Origen.Cod_Alm")
            loComandoSeleccionar.AppendLine("	JOIN Almacenes AS Destino ON Traslados.Alm_Des = Destino.Cod_Alm")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Traslados.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Renglones_Traslados.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Traslados.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Traslados'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Traslados.Documento")
            loComandoSeleccionar.AppendLine("       AND Mediciones.Origen = 'Traslados'")
            loComandoSeleccionar.AppendLine("       AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("       AND Renglones_Traslados.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'TA-NPIEZ'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Desperdicio ON Mediciones.Documento = Desperdicio.Documento")
            loComandoSeleccionar.AppendLine("		AND Desperdicio.Cod_Var = 'TA-PDESP'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Longitud ON Mediciones.Documento = Longitud.Documento")
            loComandoSeleccionar.AppendLine("		AND Longitud.Cod_Var = 'TA-LARG'")
            loComandoSeleccionar.AppendLine("WHERE Traslados.Documento BETWEEN @lcDoc_Desde AND @lcDoc_Hasta")
            loComandoSeleccionar.AppendLine("   AND Traslados.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("	AND Articulos.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("	AND Traslados.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("   AND Traslados.Alm_Ori BETWEEN @lcAlmOri_Desde AND @lcAlmOri_Hasta")
            loComandoSeleccionar.AppendLine("   AND Traslados.Alm_Des BETWEEN @lcAlmDes_Desde AND @lcAlmDes_Hasta")
            loComandoSeleccionar.AppendLine("ORDER BY Traslados.Documento, Renglones_Traslados.Renglon")
            loComandoSeleccionar.AppendLine("")


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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("CGS_rTAlmacenes_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rTAlmacenes_Renglones.ReportSource =	 loObjetoReporte	

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
' JJD: 05/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  29/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  20/07/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'
' CMS:  11/08/09: Se agregaro el filtro_ Ubicación
'-------------------------------------------------------------------------------------------'