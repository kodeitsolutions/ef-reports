Imports System.Data
Partial Class CGS_rOrdenes_Compras
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
        Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
        Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcCodSuc_Desde AS VARCHAR(15) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSuc_Hasta AS VARCHAR(15) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Ordenes_Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("      Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("      Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("      Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("      Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("      Proveedores.Telefonos,")
            loComandoSeleccionar.AppendLine("      Proveedores.Correo, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Documento, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Por_Des1 AS Por_Des1_Enc, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Mon_Des1 AS Mon_Des1_Enc, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Por_Rec1 AS Por_Rec1_Enc, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Mon_Rec1 AS Mon_Rec1_Enc, ")
            loComandoSeleccionar.AppendLine("      Renglones_OCompras.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Mon_Net, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Comentario, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Status,")
            loComandoSeleccionar.AppendLine("      COALESCE((SELECT Factory_Global.dbo.Usuarios.Nom_Usu ")
            loComandoSeleccionar.AppendLine("      FROM Factory_Global.dbo.Usuarios ")
            loComandoSeleccionar.AppendLine("      WHERE Factory_Global.dbo.Usuarios.Cod_Usu COLLATE SQL_Latin1_General_CP1_CI_AS = Ordenes_Compras.Usu_Cre COLLATE SQL_Latin1_General_CP1_CI_AS),'') AS Usuario,")
            loComandoSeleccionar.AppendLine("      Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("      Renglones_OCompras.Cod_Art, ")
            loComandoSeleccionar.AppendLine("      Articulos.Nom_Art               AS Nom_Art, ")
            loComandoSeleccionar.AppendLine("      Articulos.Generico              AS Generico,")
            loComandoSeleccionar.AppendLine("      Renglones_OCompras.Notas        AS Notas,")
            loComandoSeleccionar.AppendLine("      Renglones_OCompras.Can_Art1, ")
            loComandoSeleccionar.AppendLine("      Renglones_OCompras.Precio1      As Precio1, ")
            loComandoSeleccionar.AppendLine("      Renglones_OCompras.Mon_Net      As Neto, ")
            loComandoSeleccionar.AppendLine("      Renglones_OCompras.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Registro        As Fec_Cre, ")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Fec_Aut1        As Fec_Aut, ")
            loComandoSeleccionar.AppendLine("      COALESCE((SELECT TOP 1 Auditorias.Cod_Usu")
            loComandoSeleccionar.AppendLine("                FROM Auditorias")
            loComandoSeleccionar.AppendLine("                WHERE	Ordenes_Compras.Documento = Auditorias.Documento")
            loComandoSeleccionar.AppendLine("                  AND Auditorias.Tabla = 'Ordenes_Compras'")
            loComandoSeleccionar.AppendLine("                  AND Auditorias.Accion = 'Confirmar'")
            loComandoSeleccionar.AppendLine("                ORDER BY Auditorias.Registro DESC)")
            loComandoSeleccionar.AppendLine("      ,'')	AS Usu_Aut,")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Prioridad		AS Tipo,")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Logico1          AS mgentili,")
            'loComandoSeleccionar.AppendLine("       Ordenes_Compras.Logico2          AS ssimanca,")
            loComandoSeleccionar.AppendLine("	   CASE WHEN COALESCE((SELECT TOP 1 Auditorias.Cod_Usu")
            loComandoSeleccionar.AppendLine("							FROM Auditorias")
            loComandoSeleccionar.AppendLine("							WHERE	Ordenes_Compras.Documento = Auditorias.Documento")
            loComandoSeleccionar.AppendLine("							  AND Auditorias.Tabla = 'Ordenes_Compras'")
            loComandoSeleccionar.AppendLine("							  AND Auditorias.Accion = 'Confirmar'")
            loComandoSeleccionar.AppendLine("							ORDER BY Auditorias.Registro DESC),'') = 'ssimanca'")
            loComandoSeleccionar.AppendLine("			THEN CAST(1 AS BIT)")
            loComandoSeleccionar.AppendLine("			ELSE CAST(0 AS BIT) END		AS ssimanca,")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Logico3         AS lcarrizal,")
            'loComandoSeleccionar.AppendLine("       Ordenes_Compras.Logico4          AS yreina,")
            loComandoSeleccionar.AppendLine("	   CASE WHEN COALESCE((SELECT TOP 1 Auditorias.Cod_Usu")
            loComandoSeleccionar.AppendLine("							FROM Auditorias")
            loComandoSeleccionar.AppendLine("							WHERE	Ordenes_Compras.Documento = Auditorias.Documento")
            loComandoSeleccionar.AppendLine("							  AND Auditorias.Tabla = 'Ordenes_Compras'")
            loComandoSeleccionar.AppendLine("							  AND Auditorias.Accion = 'Confirmar'")
            loComandoSeleccionar.AppendLine("							ORDER BY Auditorias.Registro DESC),'') = 'yreina'")
            loComandoSeleccionar.AppendLine("			THEN CAST(1 AS BIT)")
            loComandoSeleccionar.AppendLine("			ELSE CAST(0 AS BIT) END		AS yreina,")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Fecha1           AS Faut_mgentili,")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Fecha2           AS Faut_ssimanca,")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Fecha3           AS Faut_lcarrizal,")
            loComandoSeleccionar.AppendLine("      Ordenes_Compras.Fecha4           AS Faut_yreina   ")
            loComandoSeleccionar.AppendLine("FROM Ordenes_Compras")
            loComandoSeleccionar.AppendLine("  JOIN Renglones_OCompras ON Ordenes_Compras.Documento = Renglones_OCompras.Documento")
            loComandoSeleccionar.AppendLine("  JOIN Proveedores ON  Ordenes_Compras.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("  JOIN Formas_Pagos ON Ordenes_Compras.Cod_For = Formas_Pagos.Cod_For")
            loComandoSeleccionar.AppendLine("  JOIN Articulos ON Articulos.Cod_Art = Renglones_OCompras.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE (Ordenes_Compras.Documento LIKE '%'+" & lcParametro1Desde)
            If lcParametro2Desde <> "''" Then
                loComandoSeleccionar.AppendLine("   OR Ordenes_Compras.Documento LIKE '%'+" & lcParametro2Desde)
            End If
            If lcParametro3Desde <> "''" Then
                loComandoSeleccionar.AppendLine("   OR Ordenes_Compras.Documento LIKE '%'+" & lcParametro3Desde)
            End If
            If lcParametro4Desde <> "''" Then
                loComandoSeleccionar.AppendLine("   OR Ordenes_Compras.Documento LIKE '%'+" & lcParametro4Desde)
            End If
            If lcParametro5Desde <> "''" Then
                loComandoSeleccionar.AppendLine("   OR Ordenes_Compras.Documento LIKE '%'+" & lcParametro5Desde)
            End If
            If lcParametro6Desde <> "''" Then
                loComandoSeleccionar.AppendLine("   OR Ordenes_Compras.Documento LIKE '%'+" & lcParametro6Desde)
            End If
            If lcParametro7Desde <> "''" Then
                loComandoSeleccionar.AppendLine("   OR Ordenes_Compras.Documento LIKE '%'+" & lcParametro7Desde)
            End If
            If lcParametro8Desde <> "''" Then
                loComandoSeleccionar.AppendLine("   OR Ordenes_Compras.Documento LIKE '%'+" & lcParametro8Desde)
            End If
            If lcParametro9Desde <> "''" Then
                loComandoSeleccionar.AppendLine("   OR Ordenes_Compras.Documento LIKE '%'+" & lcParametro9Desde)
            End If
            If lcParametro10Desde <> "''" Then
                loComandoSeleccionar.AppendLine("   OR Ordenes_Compras.Documento LIKE '%'+" & lcParametro10Desde)
            End If
            loComandoSeleccionar.AppendLine(")")
            loComandoSeleccionar.AppendLine("   AND Ordenes_Compras.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("   AND Ordenes_Compras.Prioridad <> 'PDC'")
            loComandoSeleccionar.AppendLine("   AND Ordenes_Compras.Cod_Suc BETWEEN @lcCodSuc_Desde AND @lcCodSuc_Hasta")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rOrdenes_Compras", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rOrdenes_Compras.ReportSource = loObjetoReporte

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
