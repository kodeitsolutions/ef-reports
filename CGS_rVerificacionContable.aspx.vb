'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rVerificacionContable"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rVerificacionContable
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try
            lcComandoSeleccionar.AppendLine("DECLARE @Fec_Ini AS DATETIME;")
            lcComandoSeleccionar.AppendLine("DECLARE @Fec_Fin AS DATETIME;")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SET @Fec_Ini = " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("SET @Fec_Fin = " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("")
            
            If lcParametro1Desde = "'Compras'" Then
                lcComandoSeleccionar.AppendLine("SELECT 'Facturas de Compras'						AS Tipo,")
                lcComandoSeleccionar.AppendLine("		Compras.Factura								AS Documento,")
                lcComandoSeleccionar.AppendLine("		Renglones_Compras.Cod_Art					AS Cod_Art,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Articulos.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')         AS CC_Art,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Secciones.Contable AS XML) .value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')        AS CC_Sec,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Departamentos.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')	    AS CC_Dep,")
                lcComandoSeleccionar.AppendLine("		Renglones_Compras.Cod_Imp              	    AS Cod_Imp,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Impuestos.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')         AS CC_Imp,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Tipos_Proveedores.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'') AS CC_Pro,")
                lcComandoSeleccionar.AppendLine("		Tipos_Proveedores.Cod_Tip                   AS Tip_Pro,")
                lcComandoSeleccionar.AppendLine("		'N/A'										AS Cod_Con,")
                lcComandoSeleccionar.AppendLine("		'N/A'										AS CC_Concepto")
                lcComandoSeleccionar.AppendLine("FROM Compras")
                lcComandoSeleccionar.AppendLine("	JOIN Renglones_Compras ON Compras.Documento = Renglones_Compras.Documento")
                lcComandoSeleccionar.AppendLine("	JOIN Impuestos ON Impuestos.Cod_Imp = Renglones_Compras.Cod_Imp")
                lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Renglones_Compras.Cod_Art")
                lcComandoSeleccionar.AppendLine("	JOIN Departamentos ON Departamentos.Cod_Dep = Articulos.Cod_Dep")
                lcComandoSeleccionar.AppendLine("	JOIN Secciones ON Secciones.Cod_Sec = Articulos.Cod_Sec")
                lcComandoSeleccionar.AppendLine("		AND Secciones.Cod_Dep = Departamentos.Cod_Dep")
                lcComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Compras.Cod_Pro")
                lcComandoSeleccionar.AppendLine("	JOIN Tipos_Proveedores ON Tipos_Proveedores.Cod_Tip = Proveedores.Cod_Tip")
                lcComandoSeleccionar.AppendLine("WHERE Compras.Fec_Ini BETWEEN @Fec_Ini AND @Fec_Fin")
                lcComandoSeleccionar.AppendLine("")
                lcComandoSeleccionar.AppendLine("UNION ALL")
                lcComandoSeleccionar.AppendLine("")
                lcComandoSeleccionar.AppendLine("SELECT 'Cuentas por Pagar'							AS Tipo,")
                lcComandoSeleccionar.AppendLine("		Cuentas_Pagar.Factura						AS Documento,")
                lcComandoSeleccionar.AppendLine("		'N/A'										AS Cod_Art,")
                lcComandoSeleccionar.AppendLine("		'N/A'										AS CC_Art,")
                lcComandoSeleccionar.AppendLine("		'N/A'										AS CC_Sec,")
                lcComandoSeleccionar.AppendLine("		'N/A'										AS CC_Dep,")
                lcComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Imp                 	    AS Cod_Imp,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Impuestos.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')         AS CC_Imp,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Tipos_Proveedores.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'') AS CC_Pro,")
                lcComandoSeleccionar.AppendLine("		Tipos_Proveedores.Cod_Tip                   AS Tip_Pro,")
                lcComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Con						AS Cod_Con,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Conceptos.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')         AS CC_Concepto")
                lcComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
                lcComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
                lcComandoSeleccionar.AppendLine("	JOIN Tipos_Proveedores ON Tipos_Proveedores.Cod_Tip = Proveedores.Cod_Tip")
                lcComandoSeleccionar.AppendLine("	JOIN Impuestos ON Impuestos.Cod_Imp = Cuentas_Pagar.Cod_Imp")
                lcComandoSeleccionar.AppendLine("	JOIN Conceptos ON Cuentas_Pagar.Cod_Con = Conceptos.Cod_Con ")
                lcComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Fec_Ini BETWEEN @Fec_Ini AND @Fec_Fin")
                lcComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Cod_Tip = 'FACT'")
                lcComandoSeleccionar.AppendLine("ORDER BY Documento")
            End If

            If lcParametro1Desde = "'Ventas'" Then
                lcComandoSeleccionar.AppendLine("SELECT 'Facturas de Ventas'						AS Tipo,")
                lcComandoSeleccionar.AppendLine("		Facturas.Documento							AS Documento,")
                lcComandoSeleccionar.AppendLine("		Renglones_Facturas.Cod_Art					AS Cod_Art,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Articulos.Contable AS XML).value('(/contable/ficha/cue_con)[2]', 'varchar(12)'),'')     AS CC_Art,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Secciones.Contable AS XML).value('(/contable/ficha/cue_con)[2]', 'varchar(12)'),'')     AS CC_Sec,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Departamentos.Contable AS XML).value('(/contable/ficha/cue_con)[2]', 'varchar(12)'),'')	AS CC_Dep,")
                lcComandoSeleccionar.AppendLine("		Renglones_Facturas.Cod_Imp                  AS Cod_Imp,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Impuestos.Contable AS XML).value('(/contable/ficha/cue_con)[2]', 'varchar(12)'),'')     AS CC_Imp,")
                lcComandoSeleccionar.AppendLine("		Tipos_Clientes.Cod_Tip                   	AS Tip_Cli,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Tipos_Clientes.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')AS CC_Cli,")
                lcComandoSeleccionar.AppendLine("		'N/A'										AS Cod_Con,")
                lcComandoSeleccionar.AppendLine("		'N/A'										AS CC_Concepto")
                lcComandoSeleccionar.AppendLine("FROM Facturas")
                lcComandoSeleccionar.AppendLine("	JOIN Renglones_Facturas ON Facturas.Documento = Renglones_Facturas.Documento")
                lcComandoSeleccionar.AppendLine("	JOIN Impuestos ON Impuestos.Cod_Imp = Renglones_Facturas.Cod_Imp")
                lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
                lcComandoSeleccionar.AppendLine("	JOIN Departamentos ON Departamentos.Cod_Dep = Articulos.Cod_Dep")
                lcComandoSeleccionar.AppendLine("	JOIN Secciones ON Secciones.Cod_Sec = Articulos.Cod_Sec")
                lcComandoSeleccionar.AppendLine("		AND Secciones.Cod_Dep = Departamentos.Cod_Dep")
                lcComandoSeleccionar.AppendLine("	JOIN Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli")
                lcComandoSeleccionar.AppendLine("	JOIN Tipos_Clientes ON Tipos_Clientes.Cod_Tip = Clientes.Cod_Tip")
                lcComandoSeleccionar.AppendLine("WHERE Facturas.Fec_Ini BETWEEN @Fec_Ini AND @Fec_Fin")
                lcComandoSeleccionar.AppendLine("")
                lcComandoSeleccionar.AppendLine("UNION ALL")
                lcComandoSeleccionar.AppendLine("")
                lcComandoSeleccionar.AppendLine("SELECT 'Cuentas por Cobrar'						AS Tipo,")
                lcComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Factura						AS Documento,")
                lcComandoSeleccionar.AppendLine("		'N/A'										AS Cod_Art,")
                lcComandoSeleccionar.AppendLine("		'N/A'										AS CC_Art,")
                lcComandoSeleccionar.AppendLine("		'N/A'										AS CC_Sec,")
                lcComandoSeleccionar.AppendLine("		'N/A'										AS CC_Dep,")
                lcComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Imp                   	AS Cod_Imp,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Impuestos.Contable AS XML).value('(/contable/ficha/cue_con)[2]', 'varchar(12)'),'')     AS CC_Imp,")
                lcComandoSeleccionar.AppendLine("		Tipos_Clientes.Cod_Tip                      AS Tip_Cli,")
                lcComandoSeleccionar.AppendLine("		ISNULL(CAST(Tipos_Clientes.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')AS CC_Cli,")
                lcComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Con						AS Cod_Con,")
                lcComandoSeleccionar.AppendLine("		COALESCE(CAST(Conceptos.Contable AS XML).value('(/contable/ficha/cue_con)[2]', 'varchar(12)'),")
                lcComandoSeleccionar.AppendLine("				 CAST(Conceptos.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')   AS CC_Concepto")
                lcComandoSeleccionar.AppendLine("FROM Cuentas_Cobrar")
                lcComandoSeleccionar.AppendLine("	JOIN Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
                lcComandoSeleccionar.AppendLine("	JOIN Tipos_Clientes ON Tipos_Clientes.Cod_Tip = Clientes.Cod_Tip")
                lcComandoSeleccionar.AppendLine("	JOIN Impuestos ON Impuestos.Cod_Imp = Cuentas_Cobrar.Cod_Imp")
                lcComandoSeleccionar.AppendLine("	JOIN Conceptos ON Conceptos.Cod_Con = Cuentas_Cobrar.Cod_Con")
                lcComandoSeleccionar.AppendLine("WHERE Cuentas_Cobrar.Fec_Ini BETWEEN @Fec_Ini AND @Fec_Fin")
                lcComandoSeleccionar.AppendLine("	AND Cuentas_Cobrar.Cod_Tip = 'FACT'")
                lcComandoSeleccionar.AppendLine("ORDER BY Documento")
            End If

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rVerificacionContable", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rVerificacionContable.ReportSource = loObjetoReporte

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