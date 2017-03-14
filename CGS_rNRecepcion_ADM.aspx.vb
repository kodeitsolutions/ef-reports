Imports System.Data
Partial Class CGS_rNRecepcion_ADM

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
            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde	    AS DATETIME = " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta	    AS DATETIME = " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde	AS VARCHAR(10) = " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta	AS VARCHAR(10) = " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Desde	AS VARCHAR(10) = " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Hasta	AS VARCHAR(10) = " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Desde	AS VARCHAR(10) = " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Hasta	AS VARCHAR(10) = " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Desde	AS VARCHAR(10) = " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Hasta	AS VARCHAR(10) = " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Desde	AS VARCHAR(10) = " & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Hasta	AS VARCHAR(10) = " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodSuc_Desde	AS VARCHAR(10) = " & lcParametro6Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodSuc_Hasta	AS VARCHAR(10) = " & lcParametro6Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcDocOri	        AS VARCHAR(10) = " & lcParametro8Desde)
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT Recepciones.Documento, ")
            lcComandoSeleccionar.AppendLine("       Recepciones.Cod_Pro, ")
            lcComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro,")
            lcComandoSeleccionar.AppendLine("		Recepciones.Fec_Ini,")
            lcComandoSeleccionar.AppendLine("		Recepciones.Status,")
            lcComandoSeleccionar.AppendLine("		Recepciones.Comentario,")
            lcComandoSeleccionar.AppendLine("		Recepciones.Ord_Com                                                     AS Tipo,")
            lcComandoSeleccionar.AppendLine("       Renglones_Recepciones.Renglon,")
            lcComandoSeleccionar.AppendLine("		Renglones_Recepciones.Cod_Art, ")
            lcComandoSeleccionar.AppendLine("       Articulos.Nom_Art,")
            lcComandoSeleccionar.AppendLine("       Articulos.Generico,")
            lcComandoSeleccionar.AppendLine("		Renglones_Recepciones.Precio1	                                        AS Precio,  ")
            lcComandoSeleccionar.AppendLine("       Renglones_Recepciones.Mon_Bru	                                        AS Monto, ")
            lcComandoSeleccionar.AppendLine("       Renglones_Recepciones.Can_Art1,")
            lcComandoSeleccionar.AppendLine("		Renglones_Recepciones.Cod_Uni,")
            lcComandoSeleccionar.AppendLine("		Renglones_Recepciones.Comentario                                        AS Com_Ren,")
            lcComandoSeleccionar.AppendLine("		Renglones_Recepciones.Notas,")
            lcComandoSeleccionar.AppendLine("		Renglones_Recepciones.Doc_Ori,")
            lcComandoSeleccionar.AppendLine("		Renglones_Recepciones.Cod_Alm,")
            lcComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cod_Lot,'')	                                AS Lote,")
            lcComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cantidad,0)	                                AS Cantidad_Lote,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num, 0)				                                AS Piezas,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Desperdicio.Res_Num, 0)		                                AS Porc_Desperdicio,")
            lcComandoSeleccionar.AppendLine("		COALESCE((Desperdicio.Res_Num*Operaciones_Lotes.Cantidad)/100, ")
            lcComandoSeleccionar.AppendLine("		(Desperdicio.Res_Num*Renglones_Recepciones.Can_Art1)/100, 0)	        AS Cant_Desperdicio,")
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
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Alm_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm  FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Alm_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSuc_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Suc FROM Sucursales  WHERE Cod_Suc = @lcCodSuc_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Suc_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSuc_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Suc  FROM Sucursales  WHERE Cod_Suc = @lcCodSuc_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				AS Suc_Hasta")
            lcComandoSeleccionar.AppendLine("FROM Recepciones ")
            lcComandoSeleccionar.AppendLine("	INNER JOIN Renglones_Recepciones ON Recepciones.Documento = Renglones_Recepciones.Documento ")
            lcComandoSeleccionar.AppendLine("	INNER JOIN Proveedores ON Recepciones.Cod_Pro = Proveedores.Cod_Pro ")
            lcComandoSeleccionar.AppendLine("	INNER JOIN Articulos ON Renglones_Recepciones.Cod_Art	= Articulos.Cod_Art")
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
            lcComandoSeleccionar.AppendLine("WHERE Recepciones.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Renglones_Recepciones.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Proveedores.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Recepciones.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Recepciones.Cod_Suc	BETWEEN	@lcCodSuc_Desde AND @lcCodSuc_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Recepciones.Status IN ( " & lcParametro7Desde & ")")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Recepciones.Doc_Ori LIKE '%'+RTRIM(@lcDocOri)+'%'")
            
            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rNRecepcion_ADM", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rNRecepcion_ADM.ReportSource = loObjetoReporte

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

