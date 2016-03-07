'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rFacturas_EnLotes"
'-------------------------------------------------------------------------------------------'
Partial Class rFacturas_EnLotes

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
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
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
            Dim lcParametro14Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(14))
            Dim lcParametro14Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(14))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Facturas.Documento, ")
            loComandoSeleccionar.AppendLine("           Facturas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Facturas.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Facturas.Tasa, ")
            loComandoSeleccionar.AppendLine("           Facturas.Status, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Bru		AS Tot_Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Imp1		AS Tot_Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Des1		AS Tot_Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Rec1		AS Tot_Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Net		AS Tot_Mon_Net,")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Facturas.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Facturas.Rif = '') THEN Clientes.Rif ELSE Facturas.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Facturas.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Facturas.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Facturas.Telefonos = '') THEN Clientes.Telefonos ELSE Facturas.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli,")
            'loComandoSeleccionar.AppendLine("           Clientes.Rif,")
            'loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis,")
            'loComandoSeleccionar.AppendLine("           Clientes.Telefonos,")
            loComandoSeleccionar.AppendLine("           Clientes.Nit,")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Comentario, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven,")
            loComandoSeleccionar.AppendLine("           Transportes.Nom_Tra,")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For")
            loComandoSeleccionar.AppendLine("FROM		Facturas ")
            loComandoSeleccionar.AppendLine("   JOIN    Renglones_Facturas	ON Facturas.Documento = Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("   JOIN    Clientes			ON Facturas.Cod_Cli	= Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("   JOIN    Articulos			ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
            loComandoSeleccionar.AppendLine("   JOIN    Vendedores			ON Facturas.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("   JOIN    Formas_Pagos		ON Facturas.Cod_For = Formas_Pagos.Cod_For")
            loComandoSeleccionar.AppendLine("   JOIN    Transportes			ON Facturas.Cod_Tra = Transportes.Cod_Tra")
            loComandoSeleccionar.AppendLine("WHERE		")
            loComandoSeleccionar.AppendLine("           Facturas.Documento					BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Facturas.Cod_Art		BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Fec_Ini				BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Cli				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Ven				BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Dep       		BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Mon				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Tra				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Mar				BETWEEN" & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Status					IN (" & lcParametro10Desde & ")")
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Facturas.Cod_Alm				BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_For				BETWEEN" & lcParametro12Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Rev between " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine("    			AND " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Suc				BETWEEN" & lcParametro14Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro14Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY   Facturas.Documento, Facturas.Fec_Ini, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rFacturas_EnLotes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrFacturas_EnLotes.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG:  16/06/12: Código inicial (usado en taller de reportes).								'
'-------------------------------------------------------------------------------------------'
