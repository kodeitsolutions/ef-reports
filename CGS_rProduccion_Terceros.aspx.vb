Imports System.Data
Partial Class CGS_rProduccion_Terceros

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try
            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde	    AS DATETIME = " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta	    AS DATETIME = " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde	AS VARCHAR(10) = " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta	AS VARCHAR(10) = " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodCli_Desde	AS VARCHAR(10) = " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodCli_Hasta	AS VARCHAR(10) = " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Desde	AS VARCHAR(10) = " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Hasta	AS VARCHAR(10) = " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcDocOri	        AS VARCHAR(10) = " & lcParametro6Desde)
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT Pedidos.Cod_Cli                         AS Cod_Cli, ")
            lcComandoSeleccionar.AppendLine("       Clientes.Nom_Cli                        AS Nom_Cli, ")
            lcComandoSeleccionar.AppendLine("       Clientes.Rif                            AS Rif, ")
            lcComandoSeleccionar.AppendLine("       Clientes.Dir_Fis                        AS Dir_Fis, ")
            lcComandoSeleccionar.AppendLine("       Clientes.Telefonos                      AS Telefonos, ")
            lcComandoSeleccionar.AppendLine("	    Ajustes.Documento					    AS Documento,")
            lcComandoSeleccionar.AppendLine("	    Ajustes.Status					        AS Status,")
            lcComandoSeleccionar.AppendLine("	    Ajustes.Fec_Ini						    AS Fec_Ini,")
            lcComandoSeleccionar.AppendLine("	    Ajustes.Comentario					    AS Comentario,")
            lcComandoSeleccionar.AppendLine("	    Renglones_Ajustes.Renglon			    AS Renglon,")
            lcComandoSeleccionar.AppendLine("	    Renglones_Ajustes.Cod_Art			    AS Cod_Art,")
            lcComandoSeleccionar.AppendLine("	    Articulos.Nom_Art						AS Nom_Art,")
            lcComandoSeleccionar.AppendLine("       Renglones_Ajustes.Comentario            AS Peso, ")
            lcComandoSeleccionar.AppendLine("	    Renglones_Ajustes.Can_Art1              AS Cantidad,")
            lcComandoSeleccionar.AppendLine("	    Renglones_Ajustes.Cod_Uni			    AS Cod_Uni,")
            lcComandoSeleccionar.AppendLine("	    Renglones_Ajustes.Tipo			        AS Tipo,")
            lcComandoSeleccionar.AppendLine("	    Renglones_Ajustes.Cod_Alm			    AS Cod_Alm,")
            lcComandoSeleccionar.AppendLine("	    Ajustes.Referencia			            AS Doc_Ori,")
            lcComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cod_Lot,'')	AS Lote,")
            lcComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cantidad,0)	AS Cantidad_Lote,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num, 0)				AS Piezas,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Desperdicio.Res_Num, 0)		AS Porc_Desperdicio,")
            lcComandoSeleccionar.AppendLine("		COALESCE((Desperdicio.Res_Num*Operaciones_Lotes.Cantidad)/100, ")
            lcComandoSeleccionar.AppendLine("		(Desperdicio.Res_Num*Renglones_Ajustes.Can_Art1)/100, 0)	AS Cant_Desperdicio,")
            lcComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
            lcComandoSeleccionar.AppendLine("		" & lcParametro1Desde & "               AS Operacion,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				        AS Art_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				        AS Art_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodCli_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Cli FROM Clientes  WHERE Cod_Cli = @lcCodCli_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				        AS Cli_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodCli_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Cli  FROM Clientes  WHERE Cod_Cli = @lcCodCli_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				        AS Cli_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				        AS Alm_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm  FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				        AS Alm_Hasta,")
            lcComandoSeleccionar.AppendLine("       @lcDocOri                               AS Pedido")
            lcComandoSeleccionar.AppendLine("FROM Ajustes")
            lcComandoSeleccionar.AppendLine("   JOIN Renglones_Ajustes ON Ajustes.Documento = Renglones_Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("   JOIN Pedidos ON Ajustes.Referencia = Pedidos.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Clientes ON Clientes.Cod_Cli = Pedidos.Cod_Cli")
            lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art ")
            lcComandoSeleccionar.AppendLine("   LEFT JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Renglones_Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("       AND Renglones_Ajustes.Cod_Art = Operaciones_Lotes.Cod_Art")
            lcComandoSeleccionar.AppendLine("       AND Operaciones_Lotes.Ren_Ori = Renglones_Ajustes.Renglon")
            lcComandoSeleccionar.AppendLine("       AND Operaciones_Lotes.Tip_Doc = 'Ajustes_Inventarios'")
            lcComandoSeleccionar.AppendLine("   LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("       AND Mediciones.Origen = 'Ajustes_Inventarios'")
            lcComandoSeleccionar.AppendLine("       AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            lcComandoSeleccionar.AppendLine("       AND Renglones_Ajustes.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            lcComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'AINV-NPIEZ'")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Desperdicio ON Mediciones.Documento = Desperdicio.Documento")
            lcComandoSeleccionar.AppendLine("		AND Desperdicio.Cod_Var = 'AINV-PDESP'")
            lcComandoSeleccionar.AppendLine("WHERE Ajustes.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Renglones_Ajustes.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Clientes.Cod_Cli BETWEEN @lcCodCli_Desde AND @lcCodCli_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Ajustes.Status IN ( " & lcParametro5Desde & ")")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Referencia LIKE '%'+RTRIM(@lcDocOri)+'%'")
            If lcParametro1Desde = "'Recepcion'" Then
                lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Tipo = 'Entrada'")
            End If
            If lcParametro1Desde = "'Despacho'" Then
                lcComandoSeleccionar.AppendLine("	AND Ajustes.Referencia LIKE '%'+RTRIM(@lcDocOri)+'%'")
            End If

            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rProduccion_Terceros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rProduccion_Terceros.ReportSource = loObjetoReporte

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
' GS: 02/02/17: El orden de los grupos en el rpt es por el caso que el mismo lote llegue en recepciones separadas.
'-------------------------------------------------------------------------------------------'

