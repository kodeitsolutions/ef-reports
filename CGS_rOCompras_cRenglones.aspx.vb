﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rOCompras_cRenglones"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rOCompras_cRenglones
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
        Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try
            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(8) = " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(8) = " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Desde AS VARCHAR(10) = " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Hasta AS VARCHAR(10) = " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Desde AS VARCHAR(15) = " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Hasta AS VARCHAR(15) = " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodSuc_Desde AS VARCHAR(15) = " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodSuc_Hasta AS VARCHAR(15) = " & lcParametro4Hasta)
            
            lcComandoSeleccionar.AppendLine("SELECT	Ordenes_Compras.Documento, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Cod_Pro, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Status, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Comentario, ")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Renglon, ")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Cod_Uni,")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Notas,")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Comentario   AS Com_Ren, ")
            lcComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro, ")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Cod_Art,")
            lcComandoSeleccionar.AppendLine("		Articulos.Nom_Art,")
            lcComandoSeleccionar.AppendLine("		Articulos.Generico,")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Cod_Alm, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Prioridad       AS Tipo, ")
            lcComandoSeleccionar.AppendLine("       ISNULL ((SELECT SUM(Renglones_Recepciones.Can_Art1)")
            lcComandoSeleccionar.AppendLine("                   FROM Renglones_Recepciones")
            lcComandoSeleccionar.AppendLine("                   WHERE Renglones_Recepciones.Doc_Ori = Renglones_OCompras.Documento")
            lcComandoSeleccionar.AppendLine("                   AND Renglones_Recepciones.Ren_Ori = Renglones_OCompras.Renglon),0) AS Cant_Recibida,")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Can_Art1, ")
            lcComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Art_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Art_Hasta,")
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
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Suc_Hasta")
            lcComandoSeleccionar.AppendLine("FROM	Ordenes_Compras ")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_OCompras ON Ordenes_Compras.Documento = Renglones_OCompras.Documento ")
            lcComandoSeleccionar.AppendLine("	JOIN Proveedores ON Ordenes_Compras.Cod_Pro = Proveedores.Cod_Pro")
            lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_OCompras.Cod_Art =	Articulos.Cod_Art")
            lcComandoSeleccionar.AppendLine("WHERE  Ordenes_Compras.Fec_Ini	BETWEEN	@ldFecha_Desde AND @ldFecha_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Renglones_OCompras.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Proveedores.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_OCompras.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Cod_Suc	BETWEEN	@lcCodSuc_Desde AND @lcCodSuc_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Ordenes_Compras.Status IN ( " & lcParametro5Desde & ")")

            lcComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSeleccionar.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rOCompras_cRenglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rOCompras_cRenglones.ReportSource = loObjetoReporte

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
' JJD: 14/10/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 20/04/09: se agregaron las condiciones: Ordenes_Compras.Fec_Ini, Proveedores.nom_pro y Ordenes_Compras.status
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' CMS: 18/06/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS: 22/07/09: Filtro BackOrder, lo conllevo al anexo del campo Can_Pen1,
'                 verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS:  13/08/09: Se Agrego la restricción Renglones_Pedidos.Can_Pen1 <> 0 cuando el filtro 
'                   BackOrder = BackOrder
'-------------------------------------------------------------------------------------------'
' CMS: 19/03/10: se agrego el filtro cod_art
'-------------------------------------------------------------------------------------------'