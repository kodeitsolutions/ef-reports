'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rVentas_Articulos_MONO"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rVentas_Articulos_MONO
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
           

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(8) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(8) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodCli_Desde AS VARCHAR(10) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodCli_Hasta AS VARCHAR(10) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Desde AS VARCHAR(2) = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Hasta AS VARCHAR(2) = " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Desde AS VARCHAR(2) = " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Hasta AS VARCHAR(2) = " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("       Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("		COALESCE(CAST(Articulos.Contable AS XML).value('(/contable/ficha/cue_con)[2]', 'varchar(13)'),")
            loComandoSeleccionar.AppendLine("				CAST(Articulos.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(13)'),'')		AS CC_Art,")
            loComandoSeleccionar.AppendLine("		COALESCE((SELECT Nom_Cue FROM Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("				WHERE Cod_Cue = COALESCE(CAST(Articulos.Contable AS XML)")
            loComandoSeleccionar.AppendLine("				.value('(/contable/ficha/cue_con)[2]', 'varchar(13)'),'')),")
            loComandoSeleccionar.AppendLine("				(SELECT Nom_Cue FROM Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("				WHERE Cod_Cue = COALESCE(CAST(Articulos.Contable AS XML)")
            loComandoSeleccionar.AppendLine("				.value('(/contable/ficha/cue_con)[1]', 'varchar(13)'),'')),'NO ASIGNADO')						AS CCNom_Art,")
            loComandoSeleccionar.AppendLine("		Renglones_Facturas.Notas,  ")
            loComandoSeleccionar.AppendLine("		Departamentos.Nom_Dep, ")
            loComandoSeleccionar.AppendLine("		Secciones.Nom_Sec,")
            loComandoSeleccionar.AppendLine("		COALESCE(CAST(Secciones.Contable AS XML) .value('(/contable/ficha/cue_con)[2]', 'varchar(13)'),")
            loComandoSeleccionar.AppendLine("				CAST(Secciones.Contable AS XML) .value('(/contable/ficha/cue_con)[1]', 'varchar(13)'),'')		AS CC_Sec,")
            loComandoSeleccionar.AppendLine("		COALESCE((SELECT Nom_Cue FROM Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("				WHERE Cod_Cue = COALESCE(CAST(Secciones.Contable AS XML)")
            loComandoSeleccionar.AppendLine("				.value('(/contable/ficha/cue_con)[1]', 'varchar(13)'),'')),")
            loComandoSeleccionar.AppendLine("				(SELECT Nom_Cue FROM Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("				WHERE Cod_Cue = COALESCE(CAST(Secciones.Contable AS XML)")
            loComandoSeleccionar.AppendLine("				.value('(/contable/ficha/cue_con)[1]', 'varchar(13)'),'')),'NO ASIGNADO')						AS CCNom_Sec,")
            loComandoSeleccionar.AppendLine("		COALESCE(CAST(Departamentos.Contable AS XML) .value('(/contable/ficha/cue_con)[2]', 'varchar(13)'),")
            loComandoSeleccionar.AppendLine("				CAST(Departamentos.Contable AS XML) .value('(/contable/ficha/cue_con)[1]', 'varchar(13)'),'')	AS CC_Dep,")
            loComandoSeleccionar.AppendLine("		COALESCE((SELECT Nom_Cue FROM Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("				WHERE Cod_Cue = COALESCE(CAST(Departamentos.Contable AS XML)")
            loComandoSeleccionar.AppendLine("				.value('(/contable/ficha/cue_con)[1]', 'varchar(13)'),'')),")
            loComandoSeleccionar.AppendLine("				(SELECT Nom_Cue FROM Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("				WHERE Cod_Cue = COALESCE(CAST(Departamentos.Contable AS XML)")
            loComandoSeleccionar.AppendLine("				.value('(/contable/ficha/cue_con)[1]', 'varchar(13)'),'')),'NO ASIGNADO')						AS CCNom_Dep,")
            loComandoSeleccionar.AppendLine("       Facturas.Documento, ")
            loComandoSeleccionar.AppendLine("       Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("       Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine("       Facturas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("       Renglones_Facturas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("       Renglones_Facturas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("       Renglones_Facturas.Precio1, ")
            loComandoSeleccionar.AppendLine("       Renglones_Facturas.Mon_Net,")
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
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodCli_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Cli FROM Clientes  WHERE Cod_Cli = @lcCodCli_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Cli_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodCli_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Cli  FROM Clientes  WHERE Cod_Cli = @lcCodCli_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Cli_Hasta")
            loComandoSeleccionar.AppendLine("FROM Articulos")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Facturas ON  Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("   JOIN Facturas ON  Facturas.Documento = Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("   JOIN Clientes ON Facturas.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("   JOIN Almacenes ON Renglones_Facturas.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("   JOIN Secciones ON Articulos.Cod_Sec = Secciones.Cod_Sec")
            loComandoSeleccionar.AppendLine("   JOIN Departamentos ON Secciones.Cod_Dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine("WHERE Renglones_Facturas.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("	AND Facturas.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("	AND Facturas.Cod_Cli BETWEEN @lcCodCli_Desde AND @lcCodCli_Hasta")
            loComandoSeleccionar.AppendLine("	AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            loComandoSeleccionar.AppendLine("ORDER BY Renglones_Facturas.Cod_Art, " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rVentas_Articulos_MONO", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rVentas_Articulos_MONO.ReportSource = loObjetoReporte

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