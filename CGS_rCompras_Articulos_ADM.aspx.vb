'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rCompras_Articulos_ADM"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rCompras_Articulos_ADM
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
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(8) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(8) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Desde AS VARCHAR(10) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Hasta AS VARCHAR(10) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Desde AS VARCHAR(2) = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Hasta AS VARCHAR(2) = " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Desde AS VARCHAR(2) = " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Hasta AS VARCHAR(2) = " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("       Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("       Compras.Documento, ")
            loComandoSeleccionar.AppendLine("       Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("       Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("       Compras.Factura, ")
            loComandoSeleccionar.AppendLine("       Renglones_Compras.Notas,  ")
            loComandoSeleccionar.AppendLine("       Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine("       Renglones_Compras.Can_Art1, ")
            loComandoSeleccionar.AppendLine("       Renglones_Compras.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("       Renglones_Compras.Precio1, ")
            loComandoSeleccionar.AppendLine("       Renglones_Compras.Mon_Net,")
            loComandoSeleccionar.AppendLine("       Renglones_Compras.Tip_Ori,")
            loComandoSeleccionar.AppendLine("       Renglones_Compras.Doc_Ori,")
            loComandoSeleccionar.AppendLine("       CASE Renglones_Compras.Tip_Ori ")
            loComandoSeleccionar.AppendLine("           WHEN 'recepciones' THEN")
            loComandoSeleccionar.AppendLine("               (SELECT Fec_Ini FROM Recepciones WHERE Documento = Renglones_Compras.Doc_Ori)")
            loComandoSeleccionar.AppendLine("           WHEN 'ordenes_compras' THEN")
            loComandoSeleccionar.AppendLine("               (SELECT Fec_Ini FROM Ordenes_Compras WHERE Documento = Renglones_Compras.Doc_Ori)")
            loComandoSeleccionar.AppendLine("           ELSE '' END                 AS Fec_Ori,")
            loComandoSeleccionar.AppendLine("       Articulos.Generico,")
            loComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
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
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Sec_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSec_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Sec_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodPro_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro FROM Proveedores  WHERE Cod_Pro = @lcCodPro_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Pro_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodPro_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro  FROM Proveedores  WHERE Cod_Pro = @lcCodPro_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Pro_Hasta")
            loComandoSeleccionar.AppendLine("FROM Articulos")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Compras ON  Renglones_Compras.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("   JOIN Compras ON  Compras.Documento = Renglones_Compras.Documento")
            loComandoSeleccionar.AppendLine("   JOIN Proveedores ON Compras.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE Renglones_Compras.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("	AND Compras.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("	AND Compras.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            loComandoSeleccionar.AppendLine("	AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("   AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            If lcParametro5Desde = "GEN" Then
                loComandoSeleccionar.AppendLine("   AND Articulos.Generico = 1")
            End If
            loComandoSeleccionar.AppendLine("ORDER BY Renglones_Compras.Cod_Art")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rCompras_Articulos_ADM", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rCompras_Articulos_ADM.ReportSource = loObjetoReporte

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
' JJD: 09/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' CMS: 22/06/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS: 04/08/09: Secciones.Cod_Dep = Departamentos.Cod_Dep
'-------------------------------------------------------------------------------------------'