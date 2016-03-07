'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTVentas_Marcas"
'-------------------------------------------------------------------------------------------'
Partial Class rTVentas_Marcas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

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
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim leParametro11Desde As Integer = cusAplicacion.goReportes.paParametrosIniciales(11)
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = cusAplicacion.goReportes.paParametrosIniciales(13)
            Dim lcParametro14Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(14))
            Dim lcParametro14Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(14))
            Dim lcParametro15Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(15))
            Dim lcParametro15Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(15))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim lcSql As String

            Dim loComandoSeleccionar As New StringBuilder()


            If leParametro11Desde > 0 Then
                lcSql = "Select Top " + leParametro11Desde.ToString
            Else
                lcSql = "Select "
            End If

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine("                 Articulos.Cod_Mar                                 AS	Cod_Mar, ")
            loComandoSeleccionar.AppendLine("                 Marcas.Nom_Mar                                    AS	Nom_Mar, ")
            loComandoSeleccionar.AppendLine("                 Renglones_Facturas.Can_Art1                       AS  Can_Art, ")
            loComandoSeleccionar.AppendLine("                 Renglones_Facturas.Mon_Net                        AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("                 Renglones_Facturas.Cos_Ult1                       AS  Mon_Cos ")
            loComandoSeleccionar.AppendLine(" INTO #curTemporal ")
            loComandoSeleccionar.AppendLine(" FROM Facturas ")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Facturas	ON Facturas.Documento		=   Renglones_Facturas.Documento ")
            loComandoSeleccionar.AppendLine("                 			AND Renglones_Facturas.Cod_Art      BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("                 			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("                 			AND Renglones_Facturas.Cod_Alm     BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("                 			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" JOIN Articulos ON Renglones_Facturas.Cod_Art		=   Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Mar     BETWEEN " & lcParametro15Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro15Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Dep      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Sec      BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Cla      BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Tip      BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Pro     BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" JOIN Clientes ON  Facturas.Cod_Cli				=   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" JOIN Marcas ON Articulos.Cod_Mar					=   Marcas.Cod_Mar")
            loComandoSeleccionar.AppendLine(" WHERE		Facturas.Status  IN ('Confirmado','Afectado','Procesado')")
            loComandoSeleccionar.AppendLine("         	AND Facturas.Fec_Ini		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("         	AND Facturas.Cod_Cli      BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("         	AND Facturas.Cod_Ven      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("         	AND Facturas.Cod_Mon      BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro10Hasta)
            
            If lcParametro13Desde = "Igual" Then
                loComandoSeleccionar.AppendLine("                 AND Facturas.Cod_Rev BETWEEN " & lcParametro12Desde)
            Else
                loComandoSeleccionar.AppendLine("                 AND Facturas.Cod_Rev NOT BETWEEN " & lcParametro12Desde)
            End If

            loComandoSeleccionar.AppendLine("    	          AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("                 AND Facturas.Cod_Suc between " & lcParametro14Desde)
            loComandoSeleccionar.AppendLine("    	          AND " & lcParametro14Hasta)
            
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")            
            loComandoSeleccionar.AppendLine("			")
            
            loComandoSeleccionar.AppendLine(" SELECT SUM(Mon_Net) AS  Tot_Net ")
            loComandoSeleccionar.AppendLine(" INTO #curTemporal1 ")
            loComandoSeleccionar.AppendLine(" FROM #curTemporal ")
            
            
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")            
            loComandoSeleccionar.AppendLine("			")
            

            loComandoSeleccionar.AppendLine(" SELECT		  Cod_Mar					AS  Cod_Mar, ")
            loComandoSeleccionar.AppendLine("                 Nom_Mar                   AS  Nom_Mar, ")
            loComandoSeleccionar.AppendLine("                 SUM(Can_Art)              AS  Can_Art, ")
            loComandoSeleccionar.AppendLine("                 SUM(Mon_Net)              AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("                 SUM(Mon_Cos)              AS  Mon_Cos, ")
            loComandoSeleccionar.AppendLine("                 SUM(Mon_Net)-SUM(Mon_Cos)                         AS  Mon_Gan, ")
            loComandoSeleccionar.AppendLine("                 (((SUM(Mon_Net)-SUM(Mon_Cos))/SUM(Mon_Net))*100)  AS  Por_Gan, ")
            loComandoSeleccionar.AppendLine("                 12                        AS  Por_Ven ")
            loComandoSeleccionar.AppendLine(" INTO #curTemporal2 ")
            loComandoSeleccionar.AppendLine(" FROM #curTemporal ")
            loComandoSeleccionar.AppendLine(" GROUP BY Cod_Mar, Nom_Mar ")
            loComandoSeleccionar.AppendLine(" ORDER BY      " & lcOrdenamiento)
            
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")
           

            loComandoSeleccionar.AppendLine(" Select #curTemporal2.*, #curTemporal1.Tot_Net into #curTemporal3 From #curTemporal2, #curTemporal1 ")


			loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("			")            
            loComandoSeleccionar.AppendLine("			")
            
            loComandoSeleccionar.AppendLine(lcSql)
            loComandoSeleccionar.AppendLine(" Cod_Mar               AS  Cod_Mar, ")
            loComandoSeleccionar.AppendLine(" Nom_Mar               AS  Nom_Mar, ")
            loComandoSeleccionar.AppendLine(" Can_Art               AS  Can_Art, ")
            loComandoSeleccionar.AppendLine(" Mon_Net               AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine(" Mon_Cos               AS  Mon_Cos, ")
            loComandoSeleccionar.AppendLine(" Mon_Gan               AS  Mon_Gan, ")
            loComandoSeleccionar.AppendLine(" Por_Gan               AS  Por_Gan, ")
            loComandoSeleccionar.AppendLine(" Tot_Net               AS  Tot_Net, ")
            loComandoSeleccionar.AppendLine(" (Mon_Net/Tot_Net)*100 AS  Por_Ven ")
            loComandoSeleccionar.AppendLine(" FROM #curTemporal3 ")
            
            
            
'me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTVentas_Marcas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrTVentas_Marcas.ReportSource = loObjetoReporte

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
' CMS: 28/08/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 30/06/11: Ajuste del Select, Eliminación del filtro Status
'-------------------------------------------------------------------------------------------'
' MAT: 30/06/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'