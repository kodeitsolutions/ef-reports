'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTVentas_Horas_Ipos"
'-------------------------------------------------------------------------------------------'
Partial Class rTVentas_Horas_Ipos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
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

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            
            
            loComandoSeleccionar.AppendLine("SELECT		")
            
            If cusAplicacion.goReportes.paParametrosIniciales(4) = "Si"	 Then
            
				loComandoSeleccionar.AppendLine(" SUBSTRING(CONVERT(nchar(30), Facturas.Fec_Ini,114),1,2)   AS	Hora, ")
			Else
			
				loComandoSeleccionar.AppendLine(" SUBSTRING(CONVERT(nchar(30), Facturas.Fec_Ini,114),1,5)   AS	Hora, ")   
			End If								 
			loComandoSeleccionar.AppendLine("			(CASE WHEN Renglones_Facturas.Renglon = '1' THEN 1 ELSE 0 END)	As  Renglon,")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Can_Art1                       As  Can_Art, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Status									  As  Status, ")
            loComandoSeleccionar.AppendLine(" 			(CASE WHEN Renglones_Facturas.Renglon = '1' THEN Facturas.Mon_Net   ELSE 0 END)  As  Mon_Net ")
            loComandoSeleccionar.AppendLine("INTO #tmpTemporal")
            loComandoSeleccionar.AppendLine("FROM		Facturas ")
            loComandoSeleccionar.AppendLine("JOIN Renglones_Facturas ON  Facturas.Documento		=   Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("JOIN Articulos ON  Renglones_Facturas.Cod_Art		=   Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Art      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("							AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Dep      BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("							AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Sec      BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("							AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Mar      BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("							AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Tip      BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("							AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Cla      BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine("							AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("JOIN Vendedores ON (Facturas.Cod_Ven	=   Vendedores.Cod_Ven) ")
            loComandoSeleccionar.AppendLine("JOIN Clientes ON  (Clientes.Cod_Cli	=   Facturas.Cod_Cli)")
            loComandoSeleccionar.AppendLine("WHERE	Facturas.Status			IN ('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("		AND Facturas.Ipos = '1'")
            loComandoSeleccionar.AppendLine("		AND Facturas.Fec_Ini		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Cli      BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Ven      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Art      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Suc      BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND Facturas.Usu_Cre      BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro6Hasta)				   
			loComandoSeleccionar.AppendLine("GROUP BY Facturas.Fec_Ini,Renglones_Facturas.Renglon,Renglones_Facturas.Can_Art1, Facturas.Mon_Net,Facturas.Status ")
	
            

            loComandoSeleccionar.AppendLine("SELECT		Hora			AS  Hora, ")
            loComandoSeleccionar.AppendLine(" 			SUM(Renglon) 	AS  Documentos, ")
            loComandoSeleccionar.AppendLine(" 			SUM(Can_Art)	AS  Can_Art, ")
            loComandoSeleccionar.AppendLine(" 			SUM(Mon_Net)	AS  Mon_Net ")
            loComandoSeleccionar.AppendLine("FROM		#tmpTemporal ")
			loComandoSeleccionar.AppendLine("GROUP BY	Hora ")
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTVentas_Horas_Ipos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTVentas_Horas_Ipos.ReportSource = loObjetoReporte

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
' MAT: 16/06/11: Codigo Inicial																' 
'-------------------------------------------------------------------------------------------'
' MAT: 21/07/11: Corrección del total de documentos por hora								' 
'-------------------------------------------------------------------------------------------'
