'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCronograma_Requerimientos"
'-------------------------------------------------------------------------------------------'
Partial Class rCronograma_Requerimientos
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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Pedidos.Documento, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Control, ")
            loComandoSeleccionar.AppendLine("           CASE ")
            loComandoSeleccionar.AppendLine("                       WHEN  DATEPART(weekday, Pedidos.Fec_Ini) = 2 THEN 'Lun' ")
            loComandoSeleccionar.AppendLine("               	    WHEN  DATEPART(weekday, Pedidos.Fec_Ini) = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine("                       WHEN  DATEPART(weekday, Pedidos.Fec_Ini) = 4 THEN 'Mie'")
            loComandoSeleccionar.AppendLine("                       WHEN  DATEPART(weekday, Pedidos.Fec_Ini) = 5 THEN 'Jue'")
            loComandoSeleccionar.AppendLine("                       WHEN  DATEPART(weekday, Pedidos.Fec_Ini) = 6 THEN 'Vie'")
            loComandoSeleccionar.AppendLine("                       WHEN  DATEPART(weekday, Pedidos.Fec_Ini) = 7 THEN 'Sab'")
            loComandoSeleccionar.AppendLine("                       WHEN  DATEPART(weekday, Pedidos.Fec_Ini) = 1 THEN 'Dom'")
            loComandoSeleccionar.AppendLine("           END as Dia,")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_cli, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Mon ")
            loComandoSeleccionar.AppendLine(" FROM      Pedidos, Vendedores, Clientes ")
            loComandoSeleccionar.AppendLine(" WHERE     Pedidos.Cod_Ven = Vendedores.Cod_ven ")
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" 			AND Pedidos.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Pedidos.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Pedidos.Cod_Cli between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Pedidos.Cod_Ven between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Pedidos.Status  IN ( " & lcParametro4Desde & " ) ")
            loComandoSeleccionar.AppendLine(" 			AND Pedidos.Cod_Mon  between" & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY Pedidos.cod_cli ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCronograma_Requerimientos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCronograma_Requerimientos.ReportSource = loObjetoReporte

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
' CMS: 05/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 08/05/09: Ordenamiento, se le agrego la columna dia 
'-------------------------------------------------------------------------------------------'
' CMS: 26/03/10: Se cambio la funcion DATENAME por DATEPART
'-------------------------------------------------------------------------------------------'
' CMS: 30/04/10: Se aplico el metodo carga de imagen
'-------------------------------------------------------------------------------------------'