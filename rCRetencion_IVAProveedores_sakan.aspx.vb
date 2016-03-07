'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCRetencion_IVAProveedores_sakan"
'-------------------------------------------------------------------------------------------'
Partial Class rCRetencion_IVAProveedores_sakan
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcAño As String = CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Year)
            Dim lcMes As String = CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Month)
            Dim lcTelefonos As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcTelefonoEmpresa)

            Dim loConsulta As New StringBuilder()




            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpDocumentos(Tipo_Origen                 VARCHAR(50) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Fecha_Retencion             DATETIME,")
            loConsulta.AppendLine("                            Numero_Comprobante          VARCHAR(50) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Numero_DocumentoRet         VARCHAR(50) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Numero_ControlRet           VARCHAR(50) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Numero_Documento            VARCHAR(50) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Tipo_Documento              VARCHAR(50) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Numero_ControlDoc           VARCHAR(50) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Monto_Documento             DECIMAL(28,10),")
            loConsulta.AppendLine("                            Monto_ExentoDoc             DECIMAL(28,10),")
            loConsulta.AppendLine("                            Monto_BaseDoc               DECIMAL(38,10),")
            loConsulta.AppendLine("                            Porcentaje_ImpuestoDoc      DECIMAL(38,10),")
            loConsulta.AppendLine("                            Monto_ImpuestoDoc           DECIMAL(38,10),")
            loConsulta.AppendLine("                            Numero_Factura              VARCHAR(50) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Base_Retencion              DECIMAL(28,10),")
            loConsulta.AppendLine("                            Porcentaje_Retenido         DECIMAL(28,10),")
            loConsulta.AppendLine("                            Concepto                    VARCHAR(300) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Monto_Retenido              DECIMAL(28,10),")
            loConsulta.AppendLine("                            Cod_Pro                     VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Nom_Pro                     VARCHAR(200) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Rif                         VARCHAR(50) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Nit                         VARCHAR(50) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Direccion                   VARCHAR(200) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Anio                        INT,")
            loConsulta.AppendLine("                            Mes                         INT,")
            loConsulta.AppendLine("                            telefonos_empresa_cliente   VARCHAR(200) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Numero_NDB                  VARCHAR(200) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Numero_NCR                  VARCHAR(200) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Tipo_Transaccion            VARCHAR(200) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Numero_Factura_Afectada     VARCHAR(200) COLLATE DATABASE_DEFAULT);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpDocumentos( Tipo_Origen, Fecha_Retencion, Numero_Comprobante, Numero_DocumentoRet, Numero_ControlRet,")
            loConsulta.AppendLine("                            Numero_Documento, Tipo_Documento, Numero_ControlDoc, Monto_Documento, Monto_ExentoDoc,")
            loConsulta.AppendLine("                            Monto_BaseDoc, Porcentaje_ImpuestoDoc, Monto_ImpuestoDoc, Numero_Factura, Base_Retencion,")
            loConsulta.AppendLine("                            Porcentaje_Retenido, Concepto, Monto_Retenido, Cod_Pro, Nom_Pro, Rif, Nit, Direccion,")
            loConsulta.AppendLine("                            Anio, Mes, telefonos_empresa_cliente)")
            loConsulta.AppendLine("SELECT		Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loConsulta.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loConsulta.AppendLine("				Retenciones_Documentos.Num_Com		AS Numero_Comprobante,			")
            loConsulta.AppendLine("				Cuentas_Pagar.Documento				AS Numero_DocumentoRet,			")
            loConsulta.AppendLine("				Cuentas_Pagar.Control				AS Numero_ControlRet,			")
            loConsulta.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loConsulta.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loConsulta.AppendLine("				Documentos.Control					AS Numero_ControlDoc,")
            loConsulta.AppendLine("				Renglones_Pagos.Mon_Net				AS Monto_Documento,")
            loConsulta.AppendLine("				Documentos.Mon_Exe					AS Monto_ExentoDoc,				")
            loConsulta.AppendLine("				Documentos.Mon_Bas1					AS Monto_BaseDoc,				")
            loConsulta.AppendLine("				Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,		")
            loConsulta.AppendLine("				Documentos.Mon_Imp1					AS Monto_ImpuestoDoc,			")
            loConsulta.AppendLine("				Documentos.Factura		            AS Numero_Factura,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loConsulta.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loConsulta.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loConsulta.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loConsulta.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loConsulta.AppendLine("				Proveedores.Rif						AS Rif,")
            loConsulta.AppendLine("				Proveedores.Nit						AS Nit,")
            loConsulta.AppendLine("				Proveedores.Dir_Fis					AS Direccion,")
            loConsulta.AppendLine("				" & lcAño & "						AS Anio,")
            loConsulta.AppendLine("				" & lcMes & "						AS Mes,")
            loConsulta.AppendLine("				" & lcTelefonos & "					AS telefonos_empresa_cliente")
            loConsulta.AppendLine("FROM			Cuentas_Pagar")
            loConsulta.AppendLine("		JOIN	Pagos")
            loConsulta.AppendLine("			ON	Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Retenciones_Documentos ")
            loConsulta.AppendLine("			ON	Retenciones_Documentos.Documento = Pagos.Documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.Documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.clase	= 'IMPUESTO'")
            loConsulta.AppendLine("		JOIN	Renglones_Pagos ")
            loConsulta.AppendLine("			ON	Renglones_Pagos.Documento = Pagos.Documento")
            loConsulta.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loConsulta.AppendLine("		LEFT JOIN Cuentas_Pagar AS Documentos")
            loConsulta.AppendLine("			ON	Documentos.Documento = Renglones_Pagos.Doc_Ori")
            loConsulta.AppendLine("			AND	Documentos.Cod_Tip = Renglones_Pagos.Cod_Tip")
            loConsulta.AppendLine("		JOIN	Proveedores")
            loConsulta.AppendLine("			ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("		LEFT JOIN Retenciones")
            loConsulta.AppendLine("			ON	Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE		Cuentas_Pagar.Cod_Tip = 'RETIVA'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")

            loConsulta.AppendLine("           AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("         		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("           AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("         		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("           AND Pagos.Cod_Mon BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("         		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("           AND Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("         		AND " & lcParametro3Hasta)

            loConsulta.AppendLine("UNION ALL		")

            loConsulta.AppendLine("SELECT		Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loConsulta.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loConsulta.AppendLine("				Retenciones_Documentos.Num_Com		AS Numero_Comprobante,				")
            loConsulta.AppendLine("				Cuentas_Pagar.Documento				AS Numero_DocumentoRet,				")
            loConsulta.AppendLine("				Cuentas_Pagar.Control				AS Numero_ControlRet,				")
            loConsulta.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loConsulta.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loConsulta.AppendLine("				Documentos.Control					AS Numero_ControlDoc,")
            loConsulta.AppendLine("				Documentos.Mon_Net					AS Monto_Documento,")
            loConsulta.AppendLine("				Documentos.Mon_Exe					AS Monto_ExentoDoc,					")
            loConsulta.AppendLine("				Documentos.Mon_Bas1					AS Monto_BaseDoc,					")
            loConsulta.AppendLine("				Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,			")
            loConsulta.AppendLine("				Documentos.Mon_Imp1					AS Monto_ImpuestoDoc,				")
            loConsulta.AppendLine("				Documentos.Factura		            AS Numero_Factura,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loConsulta.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loConsulta.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loConsulta.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loConsulta.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loConsulta.AppendLine("				Proveedores.Rif						AS Rif,")
            loConsulta.AppendLine("				Proveedores.Nit						AS Nit,")
            loConsulta.AppendLine("				Proveedores.Dir_Fis					AS Direccion,")
            loConsulta.AppendLine("				" & lcAño & "						AS Anio,")
            loConsulta.AppendLine("				" & lcMes & "						AS Mes,")
            loConsulta.AppendLine("				" & lcTelefonos & "					AS telefonos_empresa_cliente")
            loConsulta.AppendLine("FROM			Cuentas_Pagar")
            loConsulta.AppendLine("		JOIN	Cuentas_Pagar AS Documentos ")
            loConsulta.AppendLine("			ON	Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loConsulta.AppendLine("		JOIN	Retenciones_Documentos")
            loConsulta.AppendLine("			ON	Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Proveedores ")
            loConsulta.AppendLine("			ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("		LEFT JOIN	Retenciones")
            loConsulta.AppendLine("			ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'RETIVA'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")

            loConsulta.AppendLine("       	    AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("       	  		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("       	    AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("       	  		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("       	    AND Cuentas_Pagar.Cod_Mon BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("       	  		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("       	    AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("       	  		AND " & lcParametro3Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE #tmpDocumentos ")
            loConsulta.AppendLine("SET Tipo_Transaccion    = ( CASE Tipo_Documento")
            loConsulta.AppendLine("                                WHEN 'FACT' THEN '01-REG'")
            loConsulta.AppendLine("                                WHEN 'N/DB' THEN '02-REG'")
            loConsulta.AppendLine("                                WHEN 'N/CR' THEN '03-REG'")
            loConsulta.AppendLine("                                ELSE '' ")
            loConsulta.AppendLine("                            END ),")
            loConsulta.AppendLine("    Numero_DocumentoRet = (  CASE Tipo_Documento")
            loConsulta.AppendLine("                            WHEN 'FACT' THEN Numero_DocumentoRet")
            loConsulta.AppendLine("                            WHEN 'N/DB' THEN ''")
            loConsulta.AppendLine("                            WHEN 'N/CR' THEN ''")
            loConsulta.AppendLine("                            ELSE Numero_DocumentoRet ")
            loConsulta.AppendLine("                        END ),")
            loConsulta.AppendLine("    Numero_NDB          = ( CASE Tipo_Documento")
            loConsulta.AppendLine("                                WHEN 'FACT' THEN ''")
            loConsulta.AppendLine("                                WHEN 'N/DB' THEN Numero_DocumentoRet")
            loConsulta.AppendLine("                                WHEN 'N/CR' THEN ''")
            loConsulta.AppendLine("                                ELSE ''")
            loConsulta.AppendLine("                            END ),")
            loConsulta.AppendLine("    Numero_NCR          = ( CASE Tipo_Documento")
            loConsulta.AppendLine("                                WHEN 'FACT' THEN ''")
            loConsulta.AppendLine("                                WHEN 'N/DB' THEN ''")
            loConsulta.AppendLine("                                WHEN 'N/CR' THEN Numero_DocumentoRet")
            loConsulta.AppendLine("                                ELSE '' ")
            loConsulta.AppendLine("                            END );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE  #tmpDocumentos ")
            loConsulta.AppendLine("SET     Monto_Documento     = - Monto_Documento,")
            loConsulta.AppendLine("        Monto_ExentoDoc     = - Monto_ExentoDoc,")
            loConsulta.AppendLine("        Monto_BaseDoc       = - Monto_BaseDoc,")
            loConsulta.AppendLine("        Monto_ImpuestoDoc   = - Monto_ImpuestoDoc,")
            loConsulta.AppendLine("        Base_Retencion      = - Base_Retencion,")
            loConsulta.AppendLine("        Monto_Retenido      = - Monto_Retenido ")
            loConsulta.AppendLine("WHERE   Tipo_Documento = 'N/CR';")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE  #tmpDocumentos")
            loConsulta.AppendLine("SET     Numero_Factura_Afectada = Facturas_Origen.Factura")
            loConsulta.AppendLine("FROM    (")
            loConsulta.AppendLine("            SELECT  Devoluciones_Proveedores.Doc_Des1, ")
            loConsulta.AppendLine("                    Devoluciones_Proveedores.Cla_Des1, ")
            loConsulta.AppendLine("                    Renglones_dProveedores.Doc_Ori,")
            loConsulta.AppendLine("                    Renglones_dProveedores.Tip_Ori, ")
            loConsulta.AppendLine("                    (CASE WHEN (Cuentas_Pagar.Factura > '') ")
            loConsulta.AppendLine("                        THEN Cuentas_Pagar.Factura")
            loConsulta.AppendLine("                        ELSE Cuentas_Pagar.Factura")
            loConsulta.AppendLine("                    END) AS Factura")
            loConsulta.AppendLine("            FROM    #tmpDocumentos")
            loConsulta.AppendLine("                JOIN  devoluciones_proveedores")
            loConsulta.AppendLine("                    ON  devoluciones_proveedores.cla_des1 = #tmpDocumentos.Tipo_Documento")
            loConsulta.AppendLine("                    AND #tmpDocumentos.Numero_Documento = devoluciones_proveedores.doc_des1")
            loConsulta.AppendLine("                    AND #tmpDocumentos.Tipo_Documento = 'N/CR'")
            loConsulta.AppendLine("                    AND devoluciones_proveedores.tip_des1 = 'Cuentas_Pagar'")
            loConsulta.AppendLine("                    AND devoluciones_proveedores.Status IN ('Confirmado')")
            loConsulta.AppendLine("                JOIN Renglones_dProveedores")
            loConsulta.AppendLine("                    ON  Renglones_dProveedores.Documento = devoluciones_proveedores.Documento")
            loConsulta.AppendLine("                    AND Renglones_dProveedores.Tip_Ori ='Compras'")
            loConsulta.AppendLine("                JOIN Cuentas_Pagar ")
            loConsulta.AppendLine("                    ON  Cuentas_Pagar.Documento = Renglones_dProveedores.Doc_Ori")
            loConsulta.AppendLine("                    AND Cuentas_Pagar.Cod_Tip = 'FACT'")
            loConsulta.AppendLine("        ) Facturas_Origen")
            loConsulta.AppendLine("WHERE   #tmpDocumentos.Numero_Documento = Facturas_Origen.Doc_Des1")
            loConsulta.AppendLine("    AND #tmpDocumentos.Tipo_Documento = Facturas_Origen.Cla_Des1")
            loConsulta.AppendLine("    AND #tmpDocumentos.Tipo_Documento = 'N/CR'")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   *")
            loConsulta.AppendLine("FROM     #tmpDocumentos")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento & ", Fecha_Retencion ASC")
            loConsulta.AppendLine("")



            Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº 0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCRetencion_IVAProveedores_sakan", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrrCRetencion_IVAProveedores_sakan.ReportSource = loObjetoReporte


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
' RJG: 01/08/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RJG: 20/03/10: Agregado el filtro para que distinga retenciones de IVA de las de ISLR.	'
'-------------------------------------------------------------------------------------------'
' RJG: 01/06/11: Cambiado el número de control para que muestre el del documetno de origen	'
'				 en lugar del de la retención. Correcciones varioas en la interface.		'
'-------------------------------------------------------------------------------------------'
' RJG: 17/03/15: Se modificó el layout según requerimiento del cliente.		                '
'-------------------------------------------------------------------------------------------'
' RJG: 20/03/15: Se agregaron columnas de número de documento factura, N/DB, N/CR,          '
'                comprobante y tipo de operación, según requerimiento.		                '
'-------------------------------------------------------------------------------------------'
