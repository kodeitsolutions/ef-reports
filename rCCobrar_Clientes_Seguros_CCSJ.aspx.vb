'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCCobrar_Clientes_Seguros_CCSJ"
'-------------------------------------------------------------------------------------------'
Partial Class rCCobrar_Clientes_Seguros_CCSJ
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosIniciales(7)
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim loSeleccion As New StringBuilder()

            loSeleccion.AppendLine("CREATE TABLE #tmpFacturas(  Seguro_Codigo   VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loSeleccion.AppendLine("                            Seguro_Nombre   VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loSeleccion.AppendLine("                            Seguro_Rif      VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loSeleccion.AppendLine("                            Numero          INT,")
            loSeleccion.AppendLine("                            Fecha_Factura   DATETIME,")
            loSeleccion.AppendLine("                            Fecha_Recepcion DATETIME,")
            loSeleccion.AppendLine("                            Fecha_Entrega   DATETIME,")
            loSeleccion.AppendLine("                            Titular_Nombre  VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loSeleccion.AppendLine("                            Titular_Cedula  VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loSeleccion.AppendLine("                            Paciente_Nombre VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loSeleccion.AppendLine("                            Paciente_Cedula VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loSeleccion.AppendLine("                            Factura         VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loSeleccion.AppendLine("                            Clave           VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loSeleccion.AppendLine("                            Monto_Neto      DECIMAL(28,10),")
            loSeleccion.AppendLine("                            Monto_Cobertura DECIMAL(28,10) DEFAULT(0),")
            loSeleccion.AppendLine("                            tmp_Cliente     VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loSeleccion.AppendLine("                            tmp_Rif         VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loSeleccion.AppendLine("                            tmp_Nit         VARCHAR(20) COLLATE DATABASE_DEFAULT ")
            loSeleccion.AppendLine("                            );")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("INSERT INTO #tmpFacturas(   Numero, Seguro_Codigo, Seguro_Nombre, Seguro_Rif,")
            loSeleccion.AppendLine("                            Fecha_Factura,  Factura, Monto_Neto, ")
            loSeleccion.AppendLine("                            Paciente_Nombre, Paciente_Cedula,")
            loSeleccion.AppendLine("                            Titular_Nombre, Titular_Cedula,")
            loSeleccion.AppendLine("                            Clave, Fecha_Entrega, Fecha_Recepcion,")
            loSeleccion.AppendLine("                            tmp_Cliente, tmp_Rif, tmp_Nit)")
            loSeleccion.AppendLine("SELECT      ROW_NUMBER() OVER(PARTITION BY facturas.cod_cli ")
            loSeleccion.AppendLine("                ORDER BY facturas.cod_cli ASC, facturas.fec_ini), ")
            loSeleccion.AppendLine("            cuentas_cobrar.cod_cli, ")
            loSeleccion.AppendLine("            clientes.nom_cli, ")
            loSeleccion.AppendLine("            clientes.rif, ")
            loSeleccion.AppendLine("            cuentas_cobrar.fec_ini, ")
            loSeleccion.AppendLine("            cuentas_cobrar.documento, ")
            loSeleccion.AppendLine("            cuentas_cobrar.mon_net,")
            loSeleccion.AppendLine("            Pacientes.nom_cli,")
            loSeleccion.AppendLine("            Pacientes.rif,")
            loSeleccion.AppendLine("            COALESCE(Representantes.nom_cli, ''),")
            loSeleccion.AppendLine("            COALESCE(Representantes.rif, ''),")
            loSeleccion.AppendLine("            COALESCE(Claves_Polizas.val_car, ''),")
            loSeleccion.AppendLine("            Fechas_Entrega.val_fec,")
            loSeleccion.AppendLine("            Fechas_Recepcion.val_fec,")
            loSeleccion.AppendLine("            facturas.cod_cli, ")
            loSeleccion.AppendLine("            facturas.rif, ")
            loSeleccion.AppendLine("            facturas.nit ")
            loSeleccion.AppendLine("FROM        cuentas_cobrar")
            loSeleccion.AppendLine("    JOIN    clientes ON clientes.cod_cli = cuentas_cobrar.cod_cli")
            loSeleccion.AppendLine("    LEFT JOIN facturas ON facturas.documento = cuentas_cobrar.documento")
            loSeleccion.AppendLine("    LEFT JOIN clientes Pacientes ")
            loSeleccion.AppendLine("    	ON	Pacientes.cod_cli = COALESCE(facturas.rif, cuentas_cobrar.rif)")
            loSeleccion.AppendLine("    LEFT JOIN clientes Representantes ")
            loSeleccion.AppendLine("		ON	Representantes.cod_cli = COALESCE(facturas.nit, cuentas_cobrar.nit)")
            loSeleccion.AppendLine("    LEFT JOIN campos_propiedades AS Claves_Polizas")
            loSeleccion.AppendLine("		ON	Claves_Polizas.cod_reg = cuentas_cobrar.documento")
            loSeleccion.AppendLine("		AND Claves_Polizas.origen = 'facturas'")
            loSeleccion.AppendLine("		AND Claves_Polizas.cod_pro = '400-01-002'")
            loSeleccion.AppendLine("    LEFT JOIN campos_propiedades AS Fechas_Entrega")
            loSeleccion.AppendLine("		ON	Fechas_Entrega.cod_reg = cuentas_cobrar.documento")
            loSeleccion.AppendLine("		AND Fechas_Entrega.origen = 'facturas'")
            loSeleccion.AppendLine("		AND Fechas_Entrega.cod_pro = '400-90-001'")
            loSeleccion.AppendLine("		AND Fechas_Entrega.val_fec > CAST('19000101' AS DATETIME)")
            loSeleccion.AppendLine("    LEFT JOIN campos_propiedades AS Fechas_Recepcion")
            loSeleccion.AppendLine("		ON	Fechas_Recepcion.cod_reg = cuentas_cobrar.documento")
            loSeleccion.AppendLine("		AND Fechas_Recepcion.origen = 'facturas'")
            loSeleccion.AppendLine("		AND Fechas_Recepcion.cod_pro = '400-90-002'")
            loSeleccion.AppendLine("		AND Fechas_Recepcion.val_fec > CAST('19000101' AS DATETIME)")
            loSeleccion.AppendLine("WHERE       cuentas_cobrar.cod_tip = 'FACT'")
            loSeleccion.AppendLine("            AND Cuentas_Cobrar.Documento     BETWEEN " & lcParametro0Desde)
            loSeleccion.AppendLine("            AND " & lcParametro0Hasta)
            loSeleccion.AppendLine("            AND Cuentas_Cobrar.Fec_Ini   BETWEEN " & lcParametro1Desde)
            loSeleccion.AppendLine("            AND " & lcParametro1Hasta)
            loSeleccion.AppendLine("            AND Clientes.Cod_Cli         BETWEEN " & lcParametro2Desde)
            loSeleccion.AppendLine("            AND " & lcParametro2Hasta)
            loSeleccion.AppendLine("            AND Clientes.Cod_Tip         BETWEEN " & lcParametro3Desde)
            loSeleccion.AppendLine("            AND " & lcParametro3Hasta)
            loSeleccion.AppendLine("            AND Cuentas_Cobrar.Cod_Ven   BETWEEN " & lcParametro4Desde)
            loSeleccion.AppendLine("            AND " & lcParametro4Hasta)
            loSeleccion.AppendLine("            AND Clientes.Cod_Zon		    BETWEEN " & lcParametro5Desde)
            loSeleccion.AppendLine("            AND " & lcParametro5Hasta)
            loSeleccion.AppendLine("            AND Cuentas_Cobrar.Cod_Mon   BETWEEN " & lcParametro6Desde)
            loSeleccion.AppendLine("            AND " & lcParametro6Hasta)
            If lcParametro7Desde.ToString() = "Si" Then
                loSeleccion.AppendLine("            AND Cuentas_Cobrar.Mon_Sal > 0.01")
            End If
            loSeleccion.AppendLine("            AND Cuentas_Cobrar.Cod_Suc   BETWEEN " & lcParametro8Desde)
            loSeleccion.AppendLine("            AND " & lcParametro8Hasta)
            loSeleccion.AppendLine("            AND Cuentas_Cobrar.Cod_Rev      BETWEEN " & lcParametro9Desde)
            loSeleccion.AppendLine("            AND " & lcParametro9Hasta)
            'loSeleccion.AppendLine("ORDER BY      " & lcOrdenamiento)
            loSeleccion.AppendLine("ORDER BY    cuentas_cobrar.cod_cli, cuentas_cobrar.fec_ini ASC, cuentas_cobrar.documento ASC")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("UPDATE #tmpFacturas")
            loSeleccion.AppendLine("SET Monto_Cobertura = Montos.Cobertura")
            loSeleccion.AppendLine("FROM  (")
            loSeleccion.AppendLine("		SELECT	#tmpFacturas.Factura, ")
            loSeleccion.AppendLine("				#tmpFacturas.Fecha_Factura, ")
            loSeleccion.AppendLine("				campos_propiedades.val_num As Cobertura")
            loSeleccion.AppendLine("		FROM renglones_facturas RF")
            loSeleccion.AppendLine("			JOIN #tmpFacturas ON #tmpFacturas.Factura = RF.documento")
            loSeleccion.AppendLine("			JOIN campos_propiedades ON campos_propiedades.cod_reg = RF.doc_ori")
            loSeleccion.AppendLine("				AND campos_propiedades.origen = 'pedidos'")
            loSeleccion.AppendLine("				AND campos_propiedades.cod_pro = '400-01-003'")
            loSeleccion.AppendLine("		WHERE RF.tip_ori = 'Pedidos'")
            loSeleccion.AppendLine("	) As Montos")
            loSeleccion.AppendLine("WHERE Montos.Factura = #tmpFacturas.Factura")
            loSeleccion.AppendLine("	AND Montos.Fecha_Factura = #tmpFacturas.Fecha_Factura;")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")

            Dim lcRifEmpresa As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcRifEmpresa)
            Dim lcCorreoEmpresa As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCorreo)

            loSeleccion.AppendLine("SELECT  #tmpFacturas.*, ")
            loSeleccion.AppendLine("        (" & lcRifEmpresa & ") AS Rif_Empresa, ")
            loSeleccion.AppendLine("        (" & lcCorreoEmpresa & ") AS Correo_Empresa")
            loSeleccion.AppendLine("FROM    #tmpFacturas;")
            loSeleccion.AppendLine("DROP TABLE #tmpFacturas;")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loSeleccion.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loSeleccion.ToString(), "curReportes")


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCCobrar_Clientes_Seguros_CCSJ", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCCobrar_Clientes_Seguros_CCSJ.ReportSource = loObjetoReporte

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
' Fin del codigo.                                                                           '
'-------------------------------------------------------------------------------------------'
' Bitacora de cambios.                                                                      '
'-------------------------------------------------------------------------------------------'
' RJG: 11/03/14: Programacion inicial.                                                      '
'-------------------------------------------------------------------------------------------'
' RJG: 17/03/14: Se agregó el nombre del titular al RPT (no estaba en el modelo original).  '
'-------------------------------------------------------------------------------------------'
