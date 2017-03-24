'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "MCL_rLProduccionMONO"
'-------------------------------------------------------------------------------------------'
Partial Class MCL_rLProduccionMONO

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFechaDesde AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaHasta AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcArtDesde AS VARCHAR(8) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcArtHasta AS VARCHAR(8) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcLoteDesde AS VARCHAR(15) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcLoteHasta AS VARCHAR(15) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Proyectos.Cod_Pro											    AS Orden_Produccion,")
            loComandoSeleccionar.AppendLine("       Renglones_Proyectos.Fec_Ini                                     AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("       Renglones_Proyectos.Fec_Fin                                     AS Fec_Fin,")
            loComandoSeleccionar.AppendLine("		Entrada_Lote.Cantidad										    AS Cantidad_Lote,")
            loComandoSeleccionar.AppendLine("		COALESCE(Desperdicio.Res_Num,0)								    AS Porc_Desperdicio,")
            loComandoSeleccionar.AppendLine("       (Entrada_Lote.Cantidad * COALESCE(Desperdicio.Res_Num,0))/100   AS Cant_Desperdicio,")
            loComandoSeleccionar.AppendLine("		Orden_Trabajo.Documento										    AS Orden_Trabajo,")
            loComandoSeleccionar.AppendLine("		Consumos.Documento											    AS Consumo,")
            'loComandoSeleccionar.AppendLine("       RTRIM(Renglones_CP.Cod_Art) + ' - ' + Renglones_CP.Nom_Art  AS Art_Consumido,")
            loComandoSeleccionar.AppendLine("       CASE WHEN LEN(Renglones_CP.Nom_Art) > 40")
            loComandoSeleccionar.AppendLine("            THEN CONCAT(RTRIM(Renglones_CP.Cod_Art),'-',SUBSTRING(Renglones_CP.Nom_Art, 1, 40),'...')")
            loComandoSeleccionar.AppendLine("            ELSE CONCAT(RTRIM(Renglones_CP.Cod_Art),'-',Renglones_CP.Nom_Art)")
            loComandoSeleccionar.AppendLine("       END                                                             AS Art_Consumido,")
            loComandoSeleccionar.AppendLine("		Lote_Consumido.Cantidad										    AS Cantidad_Consumida,")
            loComandoSeleccionar.AppendLine("		COALESCE(Lote_Consumido.Cod_Lot,'')							    AS Lote_Consumido,")
            'loComandoSeleccionar.AppendLine("       CASE WHEN LEN(Formulas.Nom_Art) > 30")
            'loComandoSeleccionar.AppendLine("            THEN CONCAT(RTRIM(Formulas.Cod_Art),'-',SUBSTRING(Formulas.Nom_Art, 1, 30),'...')")
            'loComandoSeleccionar.AppendLine("            ELSE CONCAT(RTRIM(Formulas.Cod_Art),'-',Formulas.Nom_Art)")
            'loComandoSeleccionar.AppendLine("       END                                                             AS Art_Obtenido,")
            loComandoSeleccionar.AppendLine("       RTRIM(Formulas.Cod_Art) + '-' + Formulas.Nom_Art              AS Art_Obtenido,")
            loComandoSeleccionar.AppendLine("		Renglones_OT.Can_Art										    AS Cantidad_Obtenida,")
            loComandoSeleccionar.AppendLine("		Lote_Obtenido.Cod_Lot										    AS Lote_Obtenido,")
            loComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num,0)									    AS Piezas,")
            loComandoSeleccionar.AppendLine("       (SELECT SUM(Consumo_P.Can_Art) FROM Renglones AS Consumo_P")
            loComandoSeleccionar.AppendLine("       WHERE Consumo_P.Documento = Consumos.Documento AND Consumo_P.Origen = 'Consumos Produccion') ")
            loComandoSeleccionar.AppendLine("       - Renglones_OT.Can_Art											AS Cant_DespReal")
            loComandoSeleccionar.AppendLine("FROM Proyectos")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Proyectos ON Proyectos.Cod_Pro =  Renglones_Proyectos.Cod_Pro")
            loComandoSeleccionar.AppendLine("	JOIN Encabezados AS Orden_Trabajo ON Proyectos.Cod_Pro = Orden_Trabajo.Proyecto")
            loComandoSeleccionar.AppendLine("		AND Orden_Trabajo.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("	JOIN Renglones AS Renglones_OT ON Orden_Trabajo.Documento = Renglones_OT.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_OT.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("	JOIN Formulas ON Renglones_OT.Cod_Reg = Formulas.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes AS Lote_Obtenido ON Lote_Obtenido.Num_Doc = Orden_Trabajo.Documento")
            loComandoSeleccionar.AppendLine("		AND Lote_Obtenido.Cod_Art = Formulas.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Lote_Obtenido.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("       AND Lote_Obtenido.Ren_Ori = Renglones_OT.Renglon ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones AS Salida ON Lote_Obtenido.Num_Doc = Salida.Cod_Reg")
            loComandoSeleccionar.AppendLine("		AND Salida.Origen = 'Encabezados'")
            loComandoSeleccionar.AppendLine("		AND Salida.Cod_Art = Lote_Obtenido.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Salida.Adicional LIKE ('%'+RTRIM(Lote_Obtenido.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Salida.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var  = 'OTRA-NPIEZ'")
            loComandoSeleccionar.AppendLine("	JOIN Encabezados  AS Consumos ON Proyectos.Cod_Pro = Consumos.Proyecto")
            loComandoSeleccionar.AppendLine("		AND Consumos.Origen = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("	JOIN Renglones AS Renglones_CP ON Consumos.Documento = Renglones_CP.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_CP.Origen = 'Consumos Produccion' ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Operaciones_Lotes AS Lote_Consumido ON Lote_Consumido.Num_Doc = Consumos.Documento")
            loComandoSeleccionar.AppendLine("		AND Lote_Consumido.Cod_Art = Renglones_CP.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Lote_Consumido.Tip_Ope = 'Salida'")
            loComandoSeleccionar.AppendLine("       AND Lote_Consumido.Ren_Ori = Renglones_CP.Renglon ")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes AS Entrada_Lote ON Lote_Consumido.Cod_Lot = Entrada_Lote.Cod_Lot")
            loComandoSeleccionar.AppendLine("		AND Lote_Consumido.Cod_Art = Entrada_Lote.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Entrada_Lote.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("		AND Entrada_Lote.Tip_Doc = 'Ajustes_Inventarios'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones AS Entrada ON Entrada_Lote.Num_Doc = Entrada.Cod_Reg")
            loComandoSeleccionar.AppendLine("		AND Entrada.Origen = 'Ajustes_Inventarios'")
            loComandoSeleccionar.AppendLine("		AND Entrada.Cod_Art = Entrada_Lote.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Entrada.Adicional LIKE ('%'+RTRIM(Entrada_Lote.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Desperdicio ON Entrada.Documento = Desperdicio.Documento")
            loComandoSeleccionar.AppendLine("		AND Desperdicio.Cod_Var  = 'AINV-PDESP'")
            loComandoSeleccionar.AppendLine("WHERE Renglones_Proyectos.Fec_Ini BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("   AND Formulas.Cod_Art BETWEEN @lcArtDesde AND @lcArtHasta")
            loComandoSeleccionar.AppendLine("	AND Lote_Obtenido.Cod_Lot BETWEEN @lcLoteDesde AND @lcLoteHasta")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("MCL_rLProduccionMONO", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvMCL_rLProduccionMONO.ReportSource = loObjetoReporte

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
' JJD: 09/01/10: Codigo inicial
'-------------------------------------------------------------------------------------------'