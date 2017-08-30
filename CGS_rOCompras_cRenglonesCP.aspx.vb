'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rOCompras_cRenglonesCP"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rOCompras_cRenglonesCP
    Inherits vis2Formularios.frmReporte

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
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
        Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try
            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(8) = " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(8) = " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Desde AS VARCHAR(2) = " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Hasta AS VARCHAR(2) = " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Desde AS VARCHAR(2) = " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Hasta AS VARCHAR(2) = " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Desde AS VARCHAR(10) = " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Hasta AS VARCHAR(10) = " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Desde AS VARCHAR(10) = " & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Hasta AS VARCHAR(10) = " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodSuc_Desde AS VARCHAR(10) = " & lcParametro6Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodSuc_Hasta AS VARCHAR(10) = " & lcParametro6Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcOrden AS VARCHAR(10) = " & lcParametro8Desde)
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT	Ordenes_Compras.Documento, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Cod_Pro, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Status, ")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Renglon, ")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Cod_Uni,")
            lcComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Cod_Art,")
            lcComandoSeleccionar.AppendLine("		Articulos.Nom_Art,")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Cod_Alm, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Prioridad                                               AS Tipo, ")
            lcComandoSeleccionar.AppendLine("       Renglones_OCompras.Precio1                                              AS Precio,")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Can_Art1, ")
            lcComandoSeleccionar.AppendLine("       Renglones_OCompras.Mon_Bru                                              AS Monto,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Recepciones.Documento, '')										AS Recepcion,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Recepciones.Fec_Ini, '')										AS Fecha_R,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Renglones_Recepciones.Renglon,0)								AS Renglon_R,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Renglones_Recepciones.Cod_Art,'')								AS CodArt_R,")
            'lcComandoSeleccionar.AppendLine("		COALESCE(Renglones_Recepciones.Can_Art1,0)								AS Cantidad_R,")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("       CASE WHEN Recepciones.Fec_Ini >= '01/08/2017'")
            lcComandoSeleccionar.AppendLine("            THEN COALESCE(Renglones_Recepciones.Can_Art1,0)")
            lcComandoSeleccionar.AppendLine("       ")
            lcComandoSeleccionar.AppendLine("		     ELSE COALESCE((CASE WHEN Renglones_Recepciones.Comentario <> ''")
            lcComandoSeleccionar.AppendLine("                                THEN CONVERT(NUMERIC(18,2), REPLACE(REPLACE(Renglones_Recepciones.Comentario,'.',''), ',','.' ) )")
            lcComandoSeleccionar.AppendLine("                                ELSE 0")
            lcComandoSeleccionar.AppendLine("                           END) ,0)")
            lcComandoSeleccionar.AppendLine("       END                                                                     AS Cantidad_R,")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("")
            'lcComandoSeleccionar.AppendLine("       CONVERT(NUMERIC(18,2), REPLACE(REPLACE(Renglones_Recepciones.Comentario,'.',''), ',','.' ) )  AS Peso_R,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Operaciones_Lotes.Cod_Lot,'')									AS Lote,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Operaciones_Lotes.Cantidad,0)									AS Cantidad_Lote,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num, 0)												AS Piezas,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Desperdicio.Res_Num, 0)										AS Porc_Desperdicio,")
            lcComandoSeleccionar.AppendLine("		COALESCE((Desperdicio.Res_Num*(Operaciones_Lotes.Cantidad+Renglones_Ajustes.Can_Art1))/100,")
            lcComandoSeleccionar.AppendLine("		(Desperdicio.Res_Num*(Renglones_Recepciones.Can_Art1++Renglones_Ajustes.Can_Art1))/100, 0)	        AS Cant_Desperdicio,")
            lcComandoSeleccionar.AppendLine("       ")
            lcComandoSeleccionar.AppendLine("       ")
            lcComandoSeleccionar.AppendLine("       ")
            lcComandoSeleccionar.AppendLine("       ")
            lcComandoSeleccionar.AppendLine("       COALESCE(Renglones_Recepciones.Caracter1,'')							AS Adicional,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Ajustes.Documento,'')			                                AS Ajuste,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Lotes_Ajustes.Cod_Lot,'')		                                AS Lote_Ajuste,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Renglones_Ajustes.Can_Art1,0)	                                AS Cantidad_Ajuste,")
            lcComandoSeleccionar.AppendLine("		COALESCE(RTRIM(Tipos_Ajustes.Nom_Tip),'')	                            AS Tipo_Ajuste,")
            lcComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Art_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Art_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodDep_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Dep FROM Departamentos WHERE Cod_Dep = @lcCodDep_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Dep_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodDep_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Dep FROM Departamentos WHERE Cod_Dep = @lcCodDep_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Dep_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSec_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Sec_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSec_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Sec_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodPro_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro FROM Proveedores  WHERE Cod_Pro = @lcCodPro_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Pro_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodPro_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro  FROM Proveedores  WHERE Cod_Pro = @lcCodPro_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Pro_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes WHERE Cod_Alm = @lcCodAlm_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Alm_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes WHERE Cod_Alm = @lcCodAlm_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Alm_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSuc_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Suc FROM Sucursales WHERE Cod_Suc = @lcCodSuc_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Suc_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSuc_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Suc FROM Sucursales WHERE Cod_Suc = @lcCodSuc_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END                AS Suc_Hasta,")
            lcComandoSeleccionar.AppendLine("       " & lcParametro8Desde & "       AS Ordenes")
            lcComandoSeleccionar.AppendLine("INTO #tmpOrdenes")
            lcComandoSeleccionar.AppendLine("FROM Ordenes_Compras ")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_OCompras ON Ordenes_Compras.Documento = Renglones_OCompras.Documento ")
            lcComandoSeleccionar.AppendLine("	JOIN Proveedores ON Ordenes_Compras.Cod_Pro = Proveedores.Cod_Pro")
            lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_OCompras.Cod_Art =	Articulos.Cod_Art")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Recepciones ON Renglones_Recepciones.Doc_Ori = Ordenes_Compras.Documento")
            'lcComandoSeleccionar.AppendLine("       AND Renglones_Recepciones.Ren_Ori = Renglones_OCompras.Renglon")
            lcComandoSeleccionar.AppendLine("       AND Renglones_Recepciones.Cod_Art = Articulos.Cod_Art")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Recepciones ON Renglones_Recepciones.Documento = Recepciones.Documento")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Renglones_Recepciones.Documento")
            lcComandoSeleccionar.AppendLine("		AND Renglones_Recepciones.Cod_Art = Operaciones_Lotes.Cod_Art")
            lcComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Recepciones.Renglon")
            lcComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Recepciones' AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Recepciones.Documento")
            lcComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Recepciones'")
            lcComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            lcComandoSeleccionar.AppendLine("		AND Renglones_Recepciones.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            lcComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'NREC-NPIEZ'")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Desperdicio ON Mediciones.Documento = Desperdicio.Documento")
            lcComandoSeleccionar.AppendLine("		AND Desperdicio.Cod_Var = 'NREC-PDESP'")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Ajustes ON Ajustes.Doc_Ori = Recepciones.Documento")
            lcComandoSeleccionar.AppendLine("		AND Ajustes.Tip_Ori = 'Recepciones'")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Ajustes ON Ajustes.Documento = Renglones_Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Almacenes ON Renglones_Ajustes.Cod_Alm = Almacenes.Cod_Alm")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Tipos_Ajustes ON Tipos_Ajustes.Cod_Tip = Renglones_Ajustes.Cod_Tip")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Operaciones_Lotes AS Lotes_Ajustes ON Lotes_Ajustes.Num_Doc = Renglones_Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("       AND Renglones_Ajustes.Cod_Art = Lotes_Ajustes.Cod_Art")
            lcComandoSeleccionar.AppendLine("       AND Lotes_Ajustes.Ren_Ori = Renglones_Ajustes.Renglon")
            lcComandoSeleccionar.AppendLine("	    AND Lotes_Ajustes.Tip_Doc = 'Ajustes_Inventarios' AND Lotes_Ajustes.Tip_Ope = 'Entrada'")
            lcComandoSeleccionar.AppendLine("	    AND Lotes_Ajustes.Cod_Lot = Operaciones_Lotes.Cod_Lot")
            lcComandoSeleccionar.AppendLine("WHERE Ordenes_Compras.Prioridad = 'PDC'")
            lcComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Fec_Ini	BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_OCompras.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Proveedores.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_OCompras.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Cod_Suc	BETWEEN	@lcCodSuc_Desde AND @lcCodSuc_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Ordenes_Compras.Documento LIKE '%'+@lcOrden+'%'")

            'If lcParametro8Desde = "'No'" Then
            lcComandoSeleccionar.AppendLine("   AND Ordenes_Compras.Status IN ( " & lcParametro7Desde & ")")
            'Else
            '    lcComandoSeleccionar.AppendLine("   AND Ordenes_Compras.Status IN ('Afectado','Confirmado','Anulado')")
            '    lcComandoSeleccionar.AppendLine("   AND (Ordenes_Compras.Documento NOT IN (SELECT Doc_Ori FROM Renglones_Recepciones)")
            '    lcComandoSeleccionar.AppendLine("   OR")
            '    lcComandoSeleccionar.AppendLine("   (ISNULL((SELECT SUM(Renglones_Recepciones.Can_Art1)")
            '    lcComandoSeleccionar.AppendLine("            FROM Renglones_Recepciones")
            '    lcComandoSeleccionar.AppendLine("            WHERE Renglones_Recepciones.Doc_Ori = Renglones_OCompras.Documento")
            '    lcComandoSeleccionar.AppendLine("	             AND Renglones_Recepciones.Ren_Ori = Renglones_OCompras.Renglon),0) < Renglones_OCompras.Can_Art1))")
            'End If
            lcComandoSeleccionar.AppendLine("ORDER BY Ordenes_Compras.Documento, Renglones_OCompras.Cod_Art, Renglones_OCompras.Renglon;")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("WITH Sumatorias (CodArt_R, Lote, Recepcion, Cantidad_R, Cantidad_Lote) AS")
            lcComandoSeleccionar.AppendLine("(")
            lcComandoSeleccionar.AppendLine("   SELECT CodArt_R, Lote, Recepcion, SUM(Cantidad_R) AS Cantidad_R, SUM(Cantidad_Lote) AS Cantidad_Lote")
            lcComandoSeleccionar.AppendLine("   FROM #tmpOrdenes")
            lcComandoSeleccionar.AppendLine("   GROUP BY Recepcion,CodArt_R,Lote")
            lcComandoSeleccionar.AppendLine(")")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("UPDATE #tmpOrdenes ")
            lcComandoSeleccionar.AppendLine("SET #tmpOrdenes.Cantidad_R = Sumatorias.Cantidad_R, ")
            lcComandoSeleccionar.AppendLine("    #tmpOrdenes.Cantidad_Lote = Sumatorias.Cantidad_Lote")
            lcComandoSeleccionar.AppendLine("FROM #tmpOrdenes")
            lcComandoSeleccionar.AppendLine("	INNER JOIN Sumatorias ON Sumatorias.CodArt_R = #tmpOrdenes.CodArt_R")
            lcComandoSeleccionar.AppendLine("		AND Sumatorias.Lote = #tmpOrdenes.Lote")
            lcComandoSeleccionar.AppendLine("WHERE Adicional <> 'DECLARADO'")
            lcComandoSeleccionar.AppendLine("   AND #tmpOrdenes.Lote = Sumatorias.Lote")
            lcComandoSeleccionar.AppendLine("   AND #tmpOrdenes.CodArt_R = Sumatorias.CodArt_R")
            lcComandoSeleccionar.AppendLine("   AND #tmpOrdenes.Recepcion = Sumatorias.Recepcion")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("UPDATE #tmpOrdenes SET Cant_Desperdicio = ((Cantidad_Lote + Cantidad_Ajuste) * Porc_Desperdicio)/ 100 WHERE Adicional <> 'DECLARADO'")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT * FROM #tmpOrdenes WHERE Adicional <> 'DECLARADO'")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("DROP TABLE #tmpOrdenes")


            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rOCompras_cRenglonesCP", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rOCompras_cRenglonesCP.ReportSource = loObjetoReporte

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
