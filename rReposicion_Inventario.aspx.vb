'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rReposicion_Inventario"
'-------------------------------------------------------------------------------------------'
Partial Class rReposicion_Inventario
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
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
            Dim lcParametro10Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro11Desde As String = cusAplicacion.goReportes.paParametrosIniciales(11)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Cod_Uni1, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Exi_Act1, ")
            
            if lcParametro11Desde = "Si" then
				loComandoSeleccionar.AppendLine(" 			(Articulos.Exi_Ped1 + Articulos.Exi_Cot1) AS Exi_Ped1, ")
			else
				loComandoSeleccionar.AppendLine(" 			Articulos.Exi_Ped1, ")
			End If
            
            loComandoSeleccionar.AppendLine(" 			(Articulos.Exi_Act1 - Articulos.Exi_Ped1) AS Disponible, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Exi_Por1, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Exi_Max, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Exi_Min, ")
            loComandoSeleccionar.AppendLine(" 			(Articulos.Exi_Max - Articulos.Exi_Act1 - Articulos.Exi_Ped1 - Articulos.Exi_Por1) AS Faltante, ")
            loComandoSeleccionar.AppendLine(" 			ISNULL(MAX(Compras.Fec_Ini),'') AS Fecha_Compras, ")
            loComandoSeleccionar.AppendLine("           CASE  ")
            loComandoSeleccionar.AppendLine("             		WHEN DATEDIFF( d, ISNULL(MAX(Compras.Fec_Ini),'') , Getdate()) >= 365 THEN 365 ")
            loComandoSeleccionar.AppendLine("             		ELSE DATEDIFF( d, ISNULL(MAX(Compras.Fec_Ini),'') , Getdate())  ")
            loComandoSeleccionar.AppendLine("           END AS							DSMC, ")
            loComandoSeleccionar.AppendLine(" 			ISNULL(MAX(Facturas.Fec_Ini),'') AS Fecha_Facturas, ")

            loComandoSeleccionar.AppendLine("           CASE  ")
            loComandoSeleccionar.AppendLine("             		WHEN DATEDIFF( d, ISNULL(MAX(Facturas.Fec_Ini),'') , Getdate()) >= 365 THEN 365 ")
            loComandoSeleccionar.AppendLine("             		ELSE DATEDIFF( d, ISNULL(MAX(Facturas.Fec_Ini),'') , Getdate())  ")
            loComandoSeleccionar.AppendLine("           END AS							DSMV ")
            loComandoSeleccionar.AppendLine(" FROM Articulos ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Renglones_Compras ON Renglones_Compras.Cod_Art = Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Compras ON Renglones_Compras.Documento = Compras.Documento AND Compras.Status IN ('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Renglones_Facturas ON Renglones_Facturas.Cod_Art = Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Facturas ON Renglones_Facturas.Documento = Facturas.Documento AND FActuras.Status IN ('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine("                 Articulos.Cod_Art		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Dep      BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Sec      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Tip     BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Cla      BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Mar      BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Pro     BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Uni1     BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Ubi      BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Cod_Suc     BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("                 AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("                 AND Articulos.Status      In ( " & lcParametro10Desde & " ) ")
            loComandoSeleccionar.AppendLine(" GROUP BY ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Cod_Uni1, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Exi_Act1, ")
            
            if lcParametro11Desde = "Si" then
				loComandoSeleccionar.AppendLine(" 			(Articulos.Exi_Ped1 + Articulos.Exi_Cot1), ")
			else
				loComandoSeleccionar.AppendLine(" 			Articulos.Exi_Ped1, ")
			End If
            
            loComandoSeleccionar.AppendLine(" 			Articulos.Exi_Ped1, ")
            loComandoSeleccionar.AppendLine(" 			(Articulos.Exi_Act1 - Articulos.Exi_Ped1), ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Exi_Por1, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Exi_Max, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Exi_Min, ")
            loComandoSeleccionar.AppendLine(" 			(Articulos.Exi_Max - Articulos.Exi_Act1 - Articulos.Exi_Ped1 - Articulos.Exi_Por1) ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rReposicion_Inventario", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrReposicion_Inventario.ReportSource = loObjetoReporte

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
' CMS: 29/08/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RJG: 08/12/10: Ajustado Estatus de Facturas de Venta: Ahora omite las facturas de venta	'
'				 pendientes e incluye las confirmadas.										'
'-------------------------------------------------------------------------------------------'
