﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCompras_aArticuloProveedorMensual"
'-------------------------------------------------------------------------------------------'
Partial Class rCompras_aArticuloProveedorMensual
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            loComandoSeleccionar.AppendLine("    SELECT    ")
            loComandoSeleccionar.AppendLine("            Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("            Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("            Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine("            CASE  ")
            loComandoSeleccionar.AppendLine("             WHEN DATEPART(MONTH, Compras.Fec_ini) = 1 THEN  Renglones_Compras.Can_Art1")
            loComandoSeleccionar.AppendLine("             else 0  ")
            loComandoSeleccionar.AppendLine("            END as Ene,")
            loComandoSeleccionar.AppendLine("            CASE  ")
            loComandoSeleccionar.AppendLine("             WHEN DATEPART(MONTH, Compras.Fec_ini) = 2 THEN Renglones_Compras.Can_Art1")
            loComandoSeleccionar.AppendLine("             else 0 ")
            loComandoSeleccionar.AppendLine("            END as Feb, ")
            loComandoSeleccionar.AppendLine("               CASE  ")
            loComandoSeleccionar.AppendLine("    	            WHEN DATEPART(MONTH, Compras.Fec_ini) = 3 THEN Renglones_Compras.Can_Art1")
            loComandoSeleccionar.AppendLine("    	            else 0 ")
            loComandoSeleccionar.AppendLine("                END as Mar, ")
            loComandoSeleccionar.AppendLine("                CASE  ")
            loComandoSeleccionar.AppendLine("    	            WHEN DATEPART(MONTH, Compras.Fec_ini) = 4 THEN  Renglones_Compras.Can_Art1")
            loComandoSeleccionar.AppendLine("    	            else 0 ")
            loComandoSeleccionar.AppendLine("                END as Abr, ")
            loComandoSeleccionar.AppendLine("                CASE  ")
            loComandoSeleccionar.AppendLine("    	            WHEN DATEPART(MONTH, Compras.Fec_ini) = 5 THEN  Renglones_Compras.Can_Art1")
            loComandoSeleccionar.AppendLine("    	            else 0 ")
            loComandoSeleccionar.AppendLine("                END as May, ")
            loComandoSeleccionar.AppendLine("                CASE  ")
            loComandoSeleccionar.AppendLine("    	            WHEN DATEPART(MONTH, Compras.Fec_ini) = 6 THEN  Renglones_Compras.Can_Art1")
            loComandoSeleccionar.AppendLine("    	            else 0 ")
            loComandoSeleccionar.AppendLine("                END as Jun, ")
            loComandoSeleccionar.AppendLine("                CASE  ")
            loComandoSeleccionar.AppendLine("    	            WHEN DATEPART(MONTH, Compras.Fec_ini) = 7 THEN  Renglones_Compras.Can_Art1")
            loComandoSeleccionar.AppendLine("    	            else 0 ")
            loComandoSeleccionar.AppendLine("                END as Jul, ")
            loComandoSeleccionar.AppendLine("                CASE  ")
            loComandoSeleccionar.AppendLine("    	            WHEN DATEPART(MONTH, Compras.Fec_ini) = 8 THEN Renglones_Compras.Can_Art1")
            loComandoSeleccionar.AppendLine("    	            else 0 ")
            loComandoSeleccionar.AppendLine("                END as Ago, ")
            loComandoSeleccionar.AppendLine("                CASE  ")
            loComandoSeleccionar.AppendLine("    	            WHEN DATEPART(MONTH, Compras.Fec_ini) = 9 THEN Renglones_Compras.Can_Art1")
            loComandoSeleccionar.AppendLine("    	            else 0 ")
            loComandoSeleccionar.AppendLine("                END as Sep, ")
            loComandoSeleccionar.AppendLine("                CASE  ")
            loComandoSeleccionar.AppendLine("    	            WHEN DATEPART(MONTH, Compras.Fec_ini) = 10 THEN  Renglones_Compras.Can_Art1")
            loComandoSeleccionar.AppendLine("    	            else 0 ")
            loComandoSeleccionar.AppendLine("                END as Oct, ")
            loComandoSeleccionar.AppendLine("                CASE  ")
            loComandoSeleccionar.AppendLine("    	            WHEN DATEPART(MONTH, Compras.Fec_ini) = 11 THEN  Renglones_Compras.Can_Art1")
            loComandoSeleccionar.AppendLine("    	            else 0 ")
            loComandoSeleccionar.AppendLine("                END as Nov, ")
            loComandoSeleccionar.AppendLine("                CASE  ")
            loComandoSeleccionar.AppendLine("    	            WHEN DATEPART(MONTH, Compras.Fec_ini) = 12 THEN  Renglones_Compras.Can_Art1")
            loComandoSeleccionar.AppendLine("    	            else 0 ")
            loComandoSeleccionar.AppendLine("                END as Dic, ")
            loComandoSeleccionar.AppendLine("                Renglones_Compras.Can_Art1 AS Total ")
            loComandoSeleccionar.AppendLine("        INTO	#tmpTemporal ")
            loComandoSeleccionar.AppendLine("        FROM")
            loComandoSeleccionar.AppendLine("                Articulos, ")
            loComandoSeleccionar.AppendLine("                Compras, ")
            loComandoSeleccionar.AppendLine("                Renglones_Compras, ")
            loComandoSeleccionar.AppendLine("        Proveedores")
            loComandoSeleccionar.AppendLine("        WHERE")
            loComandoSeleccionar.AppendLine("        Compras.Documento = Renglones_Compras.Documento")
            loComandoSeleccionar.AppendLine("                And Renglones_Compras.Cod_Art       =   Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("                And Compras.Cod_Pro                 =   Proveedores.Cod_Pro ")


            loComandoSeleccionar.AppendLine("           AND Renglones_Compras.Cod_Art       BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Fec_Ini                 BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Pro					BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Ven                 BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep               BETWEEN" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Sec               BETWEEN" & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar               BETWEEN" & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Renglones_Compras.Cod_Alm       BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Status                  IN (" & lcParametro8Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_rev                 BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            
            
			loComandoSeleccionar.AppendLine("        SELECT  ")
            loComandoSeleccionar.AppendLine("                #tmpTemporal.Nom_Art,")
            loComandoSeleccionar.AppendLine("                #tmpTemporal.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("                #tmpTemporal.Nom_Pro,   ")
            loComandoSeleccionar.AppendLine("                sum(ene) as Ene, ")
            loComandoSeleccionar.AppendLine("                sum(feb) as Feb, ")
            loComandoSeleccionar.AppendLine("                sum(mar) as Mar, ")
            loComandoSeleccionar.AppendLine("                sum(abr) as Abr, ")
            loComandoSeleccionar.AppendLine("                sum(may) as May, ")
            loComandoSeleccionar.AppendLine("                sum(jun) as Jun, ")
            loComandoSeleccionar.AppendLine("                sum(jul) as Jul, ")
            loComandoSeleccionar.AppendLine("                sum(ago) as Ago, ")
            loComandoSeleccionar.AppendLine("                sum(sep) as Sep, ")
            loComandoSeleccionar.AppendLine("                sum(oct) as Oct, ")
            loComandoSeleccionar.AppendLine("                sum(nov) as Nov, ")
            loComandoSeleccionar.AppendLine("                sum(dic) as Dic, ")
            loComandoSeleccionar.AppendLine("                sum(total) as Total")
            loComandoSeleccionar.AppendLine("        FROM #tmpTemporal ")
            loComandoSeleccionar.AppendLine("        Group By   Cod_Pro, Nom_Pro, Nom_Art ")
            loComandoSeleccionar.AppendLine("ORDER BY   Cod_Pro, Nom_Pro," & lcOrdenamiento)



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCompras_aArticuloProveedorMensual", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCompras_aArticuloProveedorMensual.ReportSource = loObjetoReporte


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
' CMS: 13/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' CMS: 04/08/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS: 26/03/10: Se cambio la funcion DATENAME por DATEPART
'-------------------------------------------------------------------------------------------'

