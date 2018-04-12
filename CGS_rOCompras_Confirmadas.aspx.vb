'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rOCompras_Confirmadas"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rOCompras_Confirmadas
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
            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcDcto_Desde AS VARCHAR(10) = " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcDcto_Hasta AS VARCHAR(10) = " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Desde AS VARCHAR(10) = " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Hasta AS VARCHAR(10) = " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodUsu_Desde AS VARCHAR(10) = " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodUsu_Hasta AS VARCHAR(10) = " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT	DISTINCT")
            lcComandoSeleccionar.AppendLine("   	Ordenes_Compras.Documento       AS Documento, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Cod_Pro         AS Cod_Pro, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Status          AS Status, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Comentario      AS Comentario, ")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Renglon      AS Renglon, ")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Cod_Uni      AS Cod_Uni,")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Notas        AS Notas,")
            lcComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro             AS Nom_Pro, ")
            lcComandoSeleccionar.AppendLine("		Ordenes_Compras.Fec_Ini         AS Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Cod_Art      AS Cod_Art,")
            lcComandoSeleccionar.AppendLine("		Articulos.Nom_Art               AS Nom_Art,")
            lcComandoSeleccionar.AppendLine("		Articulos.Generico              AS Generico,")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Cod_Alm      AS Cod_Alm, ")
            lcComandoSeleccionar.AppendLine("       Renglones_OCompras.Precio1      AS Precio,")
            lcComandoSeleccionar.AppendLine("		Renglones_OCompras.Can_Art1     AS Can_Art1, ")
            lcComandoSeleccionar.AppendLine("       Renglones_OCompras.Mon_Bru      AS Monto,")
            lcComandoSeleccionar.AppendLine("       Ordenes_Compras.Logico2         AS ssimanca,")
            lcComandoSeleccionar.AppendLine("       Ordenes_Compras.Logico3         AS lcarrizal,")
            lcComandoSeleccionar.AppendLine("       Ordenes_Compras.Logico4         AS yreina,")
            lcComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcDcto_Desde <> '' ")
            lcComandoSeleccionar.AppendLine("			 THEN CONCAT(@lcDcto_Desde, ' - ', @lcDcto_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE 'N/E'")
            lcComandoSeleccionar.AppendLine("		END								AS Docs,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodPro_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro FROM Proveedores  WHERE Cod_Pro = @lcCodPro_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Pro_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodPro_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro  FROM Proveedores  WHERE Cod_Pro = @lcCodPro_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Pro_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodUsu_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Usu FROM Factory_Global.dbo.Usuarios WHERE Cod_Usu = @lcCodUsu_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Usu_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodUsu_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Usu FROM Factory_Global.dbo.Usuarios WHERE Cod_Usu = @lcCodUsu_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Usu_Hasta")
            lcComandoSeleccionar.AppendLine("FROM	Ordenes_Compras ")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_OCompras ON Ordenes_Compras.Documento = Renglones_OCompras.Documento ")
            lcComandoSeleccionar.AppendLine("	JOIN Proveedores ON Ordenes_Compras.Cod_Pro = Proveedores.Cod_Pro")
            lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_OCompras.Cod_Art =	Articulos.Cod_Art")
            lcComandoSeleccionar.AppendLine("	JOIN Auditorias ON Ordenes_Compras.Documento = Auditorias.Documento")
            lcComandoSeleccionar.AppendLine("		AND Auditorias.Tabla = 'Ordenes_Compras'")
            lcComandoSeleccionar.AppendLine("		AND Auditorias.Accion = 'Confirmar'")
            lcComandoSeleccionar.AppendLine("WHERE Ordenes_Compras.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Documento BETWEEN @lcDcto_Desde AND @lcDcto_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Auditorias.Cod_Usu BETWEEN @lcCodUsu_Desde AND @lcCodUsu_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Ordenes_Compras.Prioridad <> 'PDC'")
            lcComandoSeleccionar.AppendLine("	AND (Ordenes_Compras.Status = 'Confirmado'")
            lcComandoSeleccionar.AppendLine("		OR (Ordenes_Compras.Status = 'Pendiente' AND Ordenes_Compras.Logico2 = 1 AND Ordenes_Compras.Logico3 = 0 AND Ordenes_Compras.Logico4 = 0)")
            lcComandoSeleccionar.AppendLine("		OR (Ordenes_Compras.Status = 'Pendiente' AND Ordenes_Compras.Logico2 = 0 AND Ordenes_Compras.Logico3 = 1 AND Ordenes_Compras.Logico4 = 0)")
            lcComandoSeleccionar.AppendLine("		OR (Ordenes_Compras.Status = 'Pendiente' AND Ordenes_Compras.Logico2 = 0 AND Ordenes_Compras.Logico3 = 0 AND Ordenes_Compras.Logico4 = 1)")
            lcComandoSeleccionar.AppendLine("	)")
            lcComandoSeleccionar.AppendLine("")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rOCompras_Confirmadas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rOCompras_Confirmadas.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error",
                          "No se pudo Completar el Proceso: " & loExcepcion.Message,
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error,
                           "auto",
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
