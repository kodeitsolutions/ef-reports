'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTVentas_Marcas_Ipos"
'-------------------------------------------------------------------------------------------'
Partial Class rTVentas_Marcas_Ipos
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
			Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" WITH curTemporal AS ( ")

            loComandoSeleccionar.AppendLine("SELECT			Articulos.Cod_Mar					  AS Cod_Mar, ")
            loComandoSeleccionar.AppendLine("               Marcas.Nom_Mar						  AS Nom_Mar, ")
            loComandoSeleccionar.AppendLine("               Renglones_Facturas.Can_Art1           AS Can_Art, ")
            loComandoSeleccionar.AppendLine("               Renglones_Facturas.Mon_Net            AS Mon_Net ")
            loComandoSeleccionar.AppendLine("FROM	Facturas ")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Facturas ON Facturas.Documento	=   Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("JOIN	Articulos ON Renglones_Facturas.Cod_Art		=   Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Art      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("							AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Dep      BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("							AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Sec      BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("							AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Tip      BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("							AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Cla      BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("							AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Mar      BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine("							AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("JOIN   Marcas ON Marcas.Cod_Mar		=   Articulos.Cod_Mar")
            loComandoSeleccionar.AppendLine("JOIN   Clientes ON Facturas.Cod_Cli				=   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE	Facturas.Status		IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine("		AND	Facturas.Fec_Ini		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND	Facturas.Cod_Cli		BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND	Facturas.Cod_Ven		BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND	Facturas.Cod_Rev		BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("   	AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("   	AND	Facturas.Cod_Suc		BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("   	AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND Facturas.Usu_Cre      BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(") ")


            loComandoSeleccionar.AppendLine(" SELECT		  Cod_Mar, ")
            loComandoSeleccionar.AppendLine("                 Nom_Mar, ")
            loComandoSeleccionar.AppendLine("                 SUM(Can_Art) AS  Can_Art, ")
            loComandoSeleccionar.AppendLine("                 SUM(Mon_Net) AS  Mon_Net ")
            loComandoSeleccionar.AppendLine(" FROM curTemporal ")
            loComandoSeleccionar.AppendLine(" GROUP BY Cod_Mar, Nom_Mar ")
            loComandoSeleccionar.AppendLine("ORDER BY  " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

           ' Me.mEscribirConsulta(loComandoSeleccionar.ToString)
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTVentas_Marcas_Ipos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTVentas_Marcas_Ipos.ReportSource = loObjetoReporte

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
' MAT: 20/06/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 21/06/11: Adición del Filtro Marca
'-------------------------------------------------------------------------------------------'

