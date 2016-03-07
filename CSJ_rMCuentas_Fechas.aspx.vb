'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CSJ_rMCuentas_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class CSJ_rMCuentas_Fechas

    Inherits vis2formularios.frmReporte

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
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            

            lcParametro4Hasta = IIf(lcParametro4Hasta = "'zzzzzzz'", "'zzzzzzz'", lcParametro4Desde)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT     Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Status, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Nom_Cue ")
            loComandoSeleccionar.AppendLine("INTO       #Tmp_Mov_Pagos ")
            loComandoSeleccionar.AppendLine("FROM       Movimientos_Cuentas, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias ")
            loComandoSeleccionar.AppendLine("WHERE      Movimientos_Cuentas.Cod_Cue     =   Cuentas_Bancarias.Cod_Cue ")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Documento   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Fec_Ini     Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Cue     Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Mon     Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Referencia  Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Status      IN  (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Tipo      IN  (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Tip_Ori = 'Pagos'")
            loComandoSeleccionar.AppendLine("ORDER BY   Movimientos_Cuentas.Cod_Cue,Movimientos_Cuentas.Tipo,DATEPART(YEAR,Movimientos_Cuentas.Fec_Ini), DATEPART(DAYOFYEAR,Movimientos_Cuentas.Fec_Ini) ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT     Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Status, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Nom_Cue ")
            loComandoSeleccionar.AppendLine("INTO       #Tmp_Mov_OPagos ")
            loComandoSeleccionar.AppendLine("FROM       Movimientos_Cuentas, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias ")
            loComandoSeleccionar.AppendLine("WHERE      Movimientos_Cuentas.Cod_Cue     =   Cuentas_Bancarias.Cod_Cue ")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Documento   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Fec_Ini     Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Cue     Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Mon     Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Referencia  Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Status      IN  (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Tipo      IN  (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Tip_Ori = 'Ordenes_Pagos'")
            loComandoSeleccionar.AppendLine("ORDER BY   Movimientos_Cuentas.Cod_Cue,Movimientos_Cuentas.Tipo,DATEPART(YEAR,Movimientos_Cuentas.Fec_Ini), DATEPART(DAYOFYEAR,Movimientos_Cuentas.Fec_Ini) ")


            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT     #Tmp_Mov_Pagos.*, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Cod_Tip AS Tipo_Origen, ")
            loComandoSeleccionar.AppendLine("           Renglones_pagos.Doc_Ori AS Nro_Origen, ")
            loComandoSeleccionar.AppendLine("           Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Pagos.Fec_Ini As Fec_Ori, ")
            loComandoSeleccionar.AppendLine("        CASE    WHEN renglones_pagos.renglon=1 THEN ")
            loComandoSeleccionar.AppendLine(" 		        CASE WHEN renglones_pagos.cod_tip='FACT' ")
            loComandoSeleccionar.AppendLine("                    THEN Left(renglones_pagos.cod_tip,5) + renglones_pagos.factura ")
            loComandoSeleccionar.AppendLine("                    ELSE left(renglones_pagos.cod_tip,5) + renglones_pagos.DOC_ORI END")
            loComandoSeleccionar.AppendLine("        ELSE '' END AS doc1, ")
            loComandoSeleccionar.AppendLine("        CASE    WHEN renglones_pagos.renglon=2 THEN ")
            loComandoSeleccionar.AppendLine(" 		        CASE WHEN renglones_pagos.cod_tip='FACT'")
            loComandoSeleccionar.AppendLine("                    THEN Left(renglones_pagos.cod_tip,5) + renglones_pagos.factura ")
            loComandoSeleccionar.AppendLine("                    ELSE left(renglones_pagos.cod_tip,5) + renglones_pagos.DOC_ORI END")
            loComandoSeleccionar.AppendLine("                ELSE '' END AS doc2, ")
            loComandoSeleccionar.AppendLine("        CASE    WHEN renglones_pagos.renglon=3 THEN ")
            loComandoSeleccionar.AppendLine(" 		        CASE WHEN renglones_pagos.cod_tip='FACT' ")
            loComandoSeleccionar.AppendLine(" 		            THEN Left(renglones_pagos.cod_tip,5) + renglones_pagos.factura ")
            loComandoSeleccionar.AppendLine("                    ELSE left(renglones_pagos.cod_tip,5) + renglones_pagos.DOC_ORI END")
            loComandoSeleccionar.AppendLine("        ELSE '' END AS doc3, ")
            loComandoSeleccionar.AppendLine("        CASE    WHEN renglones_pagos.renglon=4 THEN ")
            loComandoSeleccionar.AppendLine("     		    CASE WHEN renglones_pagos.cod_tip='FACT' ")
            loComandoSeleccionar.AppendLine("                    THEN Left(renglones_pagos.cod_tip,5) + renglones_pagos.factura ")
            loComandoSeleccionar.AppendLine("                    ELSE left(renglones_pagos.cod_tip,5) + renglones_pagos.DOC_ORI END")
            loComandoSeleccionar.AppendLine("        ELSE '' END AS doc4, ")
            loComandoSeleccionar.AppendLine("        CASE    WHEN renglones_pagos.renglon=5 THEN ")
            loComandoSeleccionar.AppendLine(" 		        CASE WHEN renglones_pagos.cod_tip='FACT' ")
            loComandoSeleccionar.AppendLine("                    THEN Left(renglones_pagos.cod_tip,5) + renglones_pagos.factura ")
            loComandoSeleccionar.AppendLine("                    ELSE left(renglones_pagos.cod_tip,5) + renglones_pagos.DOC_ORI END")
            loComandoSeleccionar.AppendLine("        ELSE '' END AS doc5 ")
            loComandoSeleccionar.AppendLine("INTO      #Tmp_Mov ")
            loComandoSeleccionar.AppendLine("FROM      #Tmp_Mov_Pagos ")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Renglones_Pagos ")
            loComandoSeleccionar.AppendLine("       ON  Renglones_Pagos.documento = #Tmp_Mov_Pagos.Doc_Ori ")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Pagos ")
            loComandoSeleccionar.AppendLine("       ON  Pagos.documento = #Tmp_Mov_Pagos.Doc_Ori ")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT     #Tmp_Mov_OPagos.*, ")
            loComandoSeleccionar.AppendLine("           Renglones_OPagos.Cod_Con AS Tipo_Origen, ")
            loComandoSeleccionar.AppendLine("           Renglones_OPagos.Notas AS Nro_Origen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Ini As Fec_Ori, ")
            loComandoSeleccionar.AppendLine("       CASE WHEN renglones_Opagos.renglon=1 ")
            loComandoSeleccionar.AppendLine("     		THEN Left(renglones_Opagos.cod_con,6) + renglones_Opagos.notas ")
            loComandoSeleccionar.AppendLine("            ELSE ''   END AS doc1, ")
            loComandoSeleccionar.AppendLine("       CASE WHEN renglones_Opagos.renglon=2 ")
            loComandoSeleccionar.AppendLine("     		THEN Left(renglones_Opagos.cod_con,6) + renglones_Opagos.notas ")
            loComandoSeleccionar.AppendLine("            ELSE ''   END AS doc2, ")
            loComandoSeleccionar.AppendLine("       CASE WHEN renglones_Opagos.renglon=3 ")
            loComandoSeleccionar.AppendLine(" 		    THEN Left(renglones_Opagos.cod_con,6) + renglones_Opagos.notas ")
            loComandoSeleccionar.AppendLine("            ELSE ''   END AS doc3, ")
            loComandoSeleccionar.AppendLine("       CASE WHEN renglones_Opagos.renglon=4 ")
            loComandoSeleccionar.AppendLine(" 		    THEN Left(renglones_Opagos.cod_con,6) + renglones_Opagos.notas ")
            loComandoSeleccionar.AppendLine("            ELSE ''   END AS doc4, ")
            loComandoSeleccionar.AppendLine("       CASE WHEN renglones_Opagos.renglon=5 ")
            loComandoSeleccionar.AppendLine(" 		    THEN Left(renglones_Opagos.cod_con,6) + renglones_Opagos.notas ")
            loComandoSeleccionar.AppendLine("            ELSE ''   END AS doc5 ")
            loComandoSeleccionar.AppendLine("FROM      #Tmp_Mov_OPagos ")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Renglones_OPagos ")
            loComandoSeleccionar.AppendLine("       ON  Renglones_OPagos.documento = #Tmp_Mov_OPagos.Doc_Ori ")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Ordenes_Pagos ")
            loComandoSeleccionar.AppendLine("       ON  Ordenes_Pagos.documento = #Tmp_Mov_OPagos.Doc_Ori ")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine(" SELECT    #Tmp_Mov.Documento As Documento, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Cod_cue) As Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Fec_Ini) As Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Status)As status, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Referencia) As Referencia,")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Cod_Mon) As Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Mon_Deb) As Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Mon_Hab) As Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Tip_ori) As Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Doc_Ori) As Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Fec_Ori) As Fec_Ori, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Tipo) As Tipo, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Num_cue) As Num_cue, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Nom_Cue) As Nom_cue, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Tipo_Origen) As Tipo_origen, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Nro_origen) As Nro_Origen, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Cod_Pro) As Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Doc1) As Doc1, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Doc2) As Doc2, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Doc3) As Doc3, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Doc4) As Doc4, ")
            loComandoSeleccionar.AppendLine("           MAX(#Tmp_Mov.Doc5) As Doc5, ")
            loComandoSeleccionar.AppendLine("           MAX(Proveedores.Nom_Pro) As Nom_Pro")
            loComandoSeleccionar.AppendLine("INTO       #Tmp_Mov2 ")
            loComandoSeleccionar.AppendLine("FROM       #Tmp_Mov ")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Proveedores ")
            loComandoSeleccionar.AppendLine("       ON  Proveedores.Cod_Pro = #Tmp_Mov.Cod_Pro ")
            loComandoSeleccionar.AppendLine("GROUP BY #Tmp_Mov.Documento ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT    #Tmp_Mov2.* ")
            loComandoSeleccionar.AppendLine("FROM      #Tmp_Mov2 ")
            loComandoSeleccionar.AppendLine("ORDER BY  #Tmp_Mov2.Cod_Cue, #Tmp_Mov2.Tipo ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("DROP TABLE #Tmp_Mov_Pagos ")
            loComandoSeleccionar.AppendLine("DROP TABLE #Tmp_Mov_OPagos ")
            loComandoSeleccionar.AppendLine("DROP TABLE #Tmp_Mov ")
            loComandoSeleccionar.AppendLine("DROP TABLE #Tmp_Mov2 ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en curReportes          '
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CSJ_rMCuentas_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCSJ_rMCuentas_Fechas.ReportSource = loObjetoReporte

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
' JJD: 21/02/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GCR: 26/03/09: Ajustes al diseño
'-------------------------------------------------------------------------------------------'
' YJP: 15/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS:  03/07/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' RJG:  09/09/13: (NOTA: El archivo fue duplicado a partir de rMCuentas_Fechas, pero no     '
'                 quedó auditoria de quien o cuando se hizo. La coloqué manual según fecha  '
'                 de última modificación).                                                  '
'-------------------------------------------------------------------------------------------'
' RJG:  08/07/14: Ajustes en formato del SELECT.                                            '
'-------------------------------------------------------------------------------------------'
