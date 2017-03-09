'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rVerificacion_Desperdicio"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rVerificacion_Desperdicio

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

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcArt_Desde	AS VARCHAR(8) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcArt_Hasta	AS VARCHAR(8) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcAlm_Desde	AS VARCHAR(30) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcAlm_Hasta	AS VARCHAR(30) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcLot_Desde	AS VARCHAR(30) = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcLot_Hasta	AS VARCHAR(30) = " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Recepciones'														AS Origen,")
            loComandoSeleccionar.AppendLine("		Recepciones.Documento												AS Documento,")
            loComandoSeleccionar.AppendLine("		Renglones_Recepciones.Renglon										AS Renglon,")
            loComandoSeleccionar.AppendLine("		Renglones_Recepciones.Cod_Art										AS Cod_Art,")
            loComandoSeleccionar.AppendLine("		Articulos.Nom_Art													AS Nom_Art,")
            loComandoSeleccionar.AppendLine("		Renglones_Recepciones.Can_Art1										AS Cantidad_Art,")
            loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm										            AS Nom_Alm,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot											AS Lote,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cantidad											AS Cantidad_Lote,")
            loComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num, 0)											AS Piezas,")
            loComandoSeleccionar.AppendLine("		COALESCE(Desperdicio.Res_Num, 0)									AS Porc_Desperdicio,")
            loComandoSeleccionar.AppendLine("		COALESCE((Desperdicio.Res_Num * Operaciones_Lotes.Cantidad)/100, 0)	AS Cant_Desperdicio")
            loComandoSeleccionar.AppendLine("FROM Recepciones")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Recepciones ON Recepciones.Documento = Renglones_Recepciones.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_Recepciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("   JOIN Almacenes ON Renglones_Recepciones.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Recepciones.Cod_Pro")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Recepciones.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Recepciones'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Recepciones.Renglon")
            loComandoSeleccionar.AppendLine("		AND Renglones_Recepciones.Cod_Art = Operaciones_Lotes.Cod_Art")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Lotes ON Renglones_Lotes.Cod_Lot = Operaciones_Lotes.Cod_Lot")
            loComandoSeleccionar.AppendLine("       AND Renglones_Lotes.Cod_Art = Operaciones_Lotes.Cod_Art")
            loComandoSeleccionar.AppendLine("       AND Renglones_Lotes.Cod_Alm = Operaciones_Lotes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Recepciones.Documento")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Recepciones'")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("		AND CAST(Renglones_Recepciones.Renglon AS VARCHAR(1)) = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'NREC-NPIEZ'")
            loComandoSeleccionar.AppendLine("		AND Piezas.Res_Num > 0")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Desperdicio ON Mediciones.Documento = Desperdicio.Documento")
            loComandoSeleccionar.AppendLine("		AND Desperdicio.Cod_Var = 'NREC-PDESP'")
            loComandoSeleccionar.AppendLine("		AND Desperdicio.Res_Num > 0")
            loComandoSeleccionar.AppendLine("WHERE Recepciones.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("	AND Articulos.Cod_Art BETWEEN @lcArt_Desde AND @lcArt_Hasta")
            loComandoSeleccionar.AppendLine("   AND Almacenes.Cod_Alm BETWEEN @lcAlm_Desde AND @lcAlm_Hasta")
            loComandoSeleccionar.AppendLine("	AND Operaciones_Lotes.Cod_Lot BETWEEN @lcLot_Desde AND @lcLot_Hasta")
            loComandoSeleccionar.AppendLine("   AND Renglones_Lotes.Exi_Act1 > 0")
            loComandoSeleccionar.AppendLine("   AND Operaciones_Lotes.Cod_Lot NOT IN ")
            loComandoSeleccionar.AppendLine("										(SELECT Lotes_Traslados.Cod_Lot")
            loComandoSeleccionar.AppendLine("										FROM Operaciones_Lotes AS Lotes_Traslados")
            loComandoSeleccionar.AppendLine("											JOIN Traslados ON Traslados.Documento = Lotes_Traslados.Num_Doc")
            loComandoSeleccionar.AppendLine("										WHERE Lotes_Traslados.Cod_Alm = Almacenes.Cod_Alm AND Lotes_Traslados.Tip_Doc = 'Traslados' ")
            loComandoSeleccionar.AppendLine("                                           AND Lotes_Traslados.Tip_Ope = 'Salida' AND Lotes_Traslados.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("											AND Traslados.Status <> 'Pendiente')")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Ajustes de Inventario'												AS Origen,")
            loComandoSeleccionar.AppendLine("		Ajustes.Documento												    AS Documento,")
            loComandoSeleccionar.AppendLine("		Renglones_Ajustes.Renglon										    AS Renglon,")
            loComandoSeleccionar.AppendLine("		Renglones_Ajustes.Cod_Art										    AS Cod_Art,")
            loComandoSeleccionar.AppendLine("		Articulos.Nom_Art													AS Nom_Art,")
            loComandoSeleccionar.AppendLine("		Renglones_Ajustes.Can_Art1										    AS Cantidad_Art,")
            loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm										            AS Nom_Alm,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot											AS Lote,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cantidad											AS Cantidad_Lote,")
            loComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num, 0)											AS Piezas,")
            loComandoSeleccionar.AppendLine("		COALESCE(Desperdicio.Res_Num, 0)									AS Porc_Desperdicio,")
            loComandoSeleccionar.AppendLine("		COALESCE((Desperdicio.Res_Num * Operaciones_Lotes.Cantidad)/100, 0)	AS Cant_Desperdicio")
            loComandoSeleccionar.AppendLine("FROM Ajustes")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Ajustes ON Ajustes.Documento = Renglones_Ajustes.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_Ajustes.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("   JOIN Almacenes ON Renglones_Ajustes.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Ajustes.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Ajustes_Inventarios'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Ajustes.Renglon")
            loComandoSeleccionar.AppendLine("		AND Renglones_Ajustes.Cod_Art = Operaciones_Lotes.Cod_Art")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Lotes ON Renglones_Lotes.Cod_Lot = Operaciones_Lotes.Cod_Lot")
            loComandoSeleccionar.AppendLine("       AND Renglones_Lotes.Cod_Art = Operaciones_Lotes.Cod_Art")
            loComandoSeleccionar.AppendLine("       AND Renglones_Lotes.Cod_Alm = Operaciones_Lotes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Ajustes.Documento")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Ajustes_Inventarios'")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("		AND CAST(Renglones_Ajustes.Renglon AS VARCHAR(1)) = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'AINV-NPIEZ'")
            loComandoSeleccionar.AppendLine("		AND Piezas.Res_Num > 0")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Desperdicio ON Mediciones.Documento = Desperdicio.Documento")
            loComandoSeleccionar.AppendLine("		AND Desperdicio.Cod_Var = 'AINV-PDESP'")
            loComandoSeleccionar.AppendLine("		AND Desperdicio.Res_Num > 0")
            loComandoSeleccionar.AppendLine("WHERE Renglones_Ajustes.Tipo = 'Entrada'")
            loComandoSeleccionar.AppendLine("	AND Ajustes.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("	AND Articulos.Cod_Art BETWEEN @lcArt_Desde AND @lcArt_Hasta")
            loComandoSeleccionar.AppendLine("   AND Almacenes.Cod_Alm BETWEEN @lcAlm_Desde AND @lcAlm_Hasta")
            loComandoSeleccionar.AppendLine("	AND Operaciones_Lotes.Cod_Lot BETWEEN @lcLot_Desde AND @lcLot_Hasta")
            loComandoSeleccionar.AppendLine("   AND Renglones_Lotes.Exi_Act1 > 0")
            loComandoSeleccionar.AppendLine("   AND Operaciones_Lotes.Cod_Lot NOT IN ")
            loComandoSeleccionar.AppendLine("										(SELECT Lotes_Traslados.Cod_Lot")
            loComandoSeleccionar.AppendLine("										FROM Operaciones_Lotes AS Lotes_Traslados")
            loComandoSeleccionar.AppendLine("											JOIN Traslados ON Traslados.Documento = Lotes_Traslados.Num_Doc")
            loComandoSeleccionar.AppendLine("										WHERE Lotes_Traslados.Cod_Alm = Almacenes.Cod_Alm AND Lotes_Traslados.Tip_Doc = 'Traslados' ")
            loComandoSeleccionar.AppendLine("                                           AND Lotes_Traslados.Tip_Ope = 'Salida' AND Lotes_Traslados.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("											AND Traslados.Status <> 'Pendiente')")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Ordenes de Trabajo'												AS Origen,")
            loComandoSeleccionar.AppendLine("		Encabezados.Documento												AS Documento,")
            loComandoSeleccionar.AppendLine("		Renglones.Renglon   												AS Renglon,")
            loComandoSeleccionar.AppendLine("		Formulas.Cod_Art													AS Cod_Art,")
            loComandoSeleccionar.AppendLine("		Formulas.Nom_Art													AS Nom_Art,")
            loComandoSeleccionar.AppendLine("		Renglones.Can_Art													AS Cantidad_Art,")
            loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm													AS Nom_Alm,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot											AS Lote,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cantidad											AS Cantidad_Lote,")
            loComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num, 0)											AS Piezas,")
            loComandoSeleccionar.AppendLine("		COALESCE(Desperdicio.Res_Num, 0)									AS Porc_Desperdicio,")
            loComandoSeleccionar.AppendLine("		COALESCE((Desperdicio.Res_Num * Operaciones_Lotes.Cantidad)/100, 0)	AS Cant_Desperdicio")
            loComandoSeleccionar.AppendLine("FROM Encabezados")
            loComandoSeleccionar.AppendLine("	JOIN Renglones ON Encabezados.Documento = Renglones.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("	JOIN Formulas ON Renglones.Cod_Reg = Formulas.Documento")
            loComandoSeleccionar.AppendLine("		AND Formulas.Origen = 'Cedula'")
            loComandoSeleccionar.AppendLine("	JOIN Almacenes ON Encabezados.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Encabezados.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Encabezados'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Adicional = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Formulas.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Lotes ON Operaciones_Lotes.Cod_Lot = Renglones_Lotes.Cod_Lot")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Renglones_Lotes.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Alm = Renglones_Lotes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Encabezados.Documento")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Encabezados'")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("		AND CAST(Renglones.Renglon AS VARCHAR(1)) = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'OTRA-NPIEZ'")
            loComandoSeleccionar.AppendLine("		AND Piezas.Res_Num > 0")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Desperdicio ON Mediciones.Documento = Desperdicio.Documento")
            loComandoSeleccionar.AppendLine("		AND Desperdicio.Cod_Var = 'OTRA-PDESP'")
            loComandoSeleccionar.AppendLine("		AND Desperdicio.Res_Num > 0")
            loComandoSeleccionar.AppendLine("WHERE Encabezados.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("	AND Encabezados.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("	AND Formulas.Cod_Art BETWEEN @lcArt_Desde AND @lcArt_Hasta")
            loComandoSeleccionar.AppendLine("	AND Almacenes.Cod_Alm BETWEEN @lcAlm_Desde AND @lcAlm_Hasta")
            loComandoSeleccionar.AppendLine("	AND Operaciones_Lotes.Cod_Lot BETWEEN @lcLot_Desde AND @lcLot_Hasta")
            loComandoSeleccionar.AppendLine("	AND Renglones_Lotes.Exi_Act1 > 0")
            loComandoSeleccionar.AppendLine("	AND Operaciones_Lotes.Cod_Lot NOT IN ")
            loComandoSeleccionar.AppendLine("									(SELECT Lotes_Traslados.Cod_Lot")
            loComandoSeleccionar.AppendLine("									FROM Operaciones_Lotes AS Lotes_Traslados")
            loComandoSeleccionar.AppendLine("										JOIN Traslados ON Traslados.Documento = Lotes_Traslados.Num_Doc")
            loComandoSeleccionar.AppendLine("									WHERE Lotes_Traslados.Cod_Alm = Almacenes.Cod_Alm AND Lotes_Traslados.Tip_Doc = 'Traslados' ")
            loComandoSeleccionar.AppendLine("                                        AND Lotes_Traslados.Tip_Ope = 'Salida' AND Lotes_Traslados.Cod_Art = Formulas.Cod_Art")
            loComandoSeleccionar.AppendLine("										AND Traslados.Status <> 'Pendiente')")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rVerificacion_Desperdicio", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rVerificacion_Desperdicio.ReportSource = loObjetoReporte

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