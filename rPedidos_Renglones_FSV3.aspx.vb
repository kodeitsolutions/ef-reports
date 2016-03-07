Imports System.Data
Partial Class rPedidos_Renglones_FSV3
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

		Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
		Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
		Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
		Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
		Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
		Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
		Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosIniciales(5)
		Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
		Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
		
		Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine("SELECT	Pedidos.Documento, ")
			loComandoSeleccionar.AppendLine("		Pedidos.Control, ")
			loComandoSeleccionar.AppendLine("		Pedidos.Fec_Ini, ")
			loComandoSeleccionar.AppendLine("		Pedidos.Fec_Fin, ")
			loComandoSeleccionar.AppendLine("		Pedidos.Cod_Cli, ")
			loComandoSeleccionar.AppendLine("		Pedidos.Cod_Ven, ")
			loComandoSeleccionar.AppendLine("		Renglones_Pedidos.Renglon, ")
			loComandoSeleccionar.AppendLine("		Renglones_Pedidos.Cod_Art, ")
			loComandoSeleccionar.AppendLine("		Renglones_Pedidos.Cod_Alm, ")
			loComandoSeleccionar.AppendLine("		Renglones_Pedidos.Precio1, ")
			loComandoSeleccionar.AppendLine("		Renglones_Pedidos.Can_Pen1  As  Can_Art1, ")
			loComandoSeleccionar.AppendLine("		Renglones_Pedidos.Cod_Uni, ")
			loComandoSeleccionar.AppendLine("		Renglones_Pedidos.Por_Imp1, ")
			loComandoSeleccionar.AppendLine("		Renglones_Pedidos.Por_Des, ")
			loComandoSeleccionar.AppendLine("		Renglones_Pedidos.Mon_Net, ")
			loComandoSeleccionar.AppendLine("		Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("		Clientes.Kilometros ")
			loComandoSeleccionar.AppendLine("FROM	Pedidos ")
			loComandoSeleccionar.AppendLine("JOIN Renglones_Pedidos ON Pedidos.Documento = Renglones_Pedidos.Documento")
			loComandoSeleccionar.AppendLine("							AND	Renglones_Pedidos.Can_Pen1 <> 0 ")
            loComandoSeleccionar.AppendLine("							AND (substring(Renglones_Pedidos.Cod_Art,1,3)='SOP' OR substring(Renglones_Pedidos.Cod_Art,1,3)='PRO' OR substring(Renglones_Pedidos.Cod_Art,1,4)='SER-') ")
			loComandoSeleccionar.AppendLine("JOIN Clientes ON Pedidos.Cod_Cli = Clientes.Cod_Cli")
			loComandoSeleccionar.AppendLine("							AND Clientes.Cod_Cli BETWEEN " & lcParametro2Desde )
			loComandoSeleccionar.AppendLine("							AND " & lcParametro2Hasta )
			loComandoSeleccionar.AppendLine("JOIN Articulos ON Renglones_Pedidos.Cod_Art = Articulos.Cod_Art")
			loComandoSeleccionar.AppendLine("							AND Articulos.Cod_Art BETWEEN " & lcParametro6Desde )
			loComandoSeleccionar.AppendLine("							AND " & lcParametro6Hasta )
			loComandoSeleccionar.AppendLine("WHERE	Pedidos.Status <> 'Anulado'  ")
			loComandoSeleccionar.AppendLine("		AND Pedidos.Documento BETWEEN " & lcParametro0Desde )
			loComandoSeleccionar.AppendLine("		AND " & lcParametro0Hasta )
			loComandoSeleccionar.AppendLine("		AND Pedidos.Fec_Ini BETWEEN " & lcParametro1Desde )
			loComandoSeleccionar.AppendLine("		AND " &  lcParametro1Hasta )
			loComandoSeleccionar.AppendLine("		AND Pedidos.Status IN ( " & lcParametro3Desde & ")" )
			
			If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Pedidos.Cod_Rev between " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Pedidos.Cod_Rev NOT between " & lcParametro4Desde)
            End If

            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY Pedidos.Cod_Cli, Pedidos.Documento")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString , "curReportes")

            '--------------------------------------------------'
			' Carga la imagen del logo de la empresa           '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPedidos_Renglones_FSV3", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPedidos_Renglones_FSV3.ReportSource = loObjetoReporte

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
' JFP: 04/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 21/07/11: Ajuste del Select, Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
' JFP: 15/03/12: Adicion de los servicios
'-------------------------------------------------------------------------------------------'