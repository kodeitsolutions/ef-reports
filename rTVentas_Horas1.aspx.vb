'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTVentas_Horas1"
'-------------------------------------------------------------------------------------------'
Partial Class rTVentas_Horas1
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Dim lcParametro4Desde As String = cusAplicacion.goReportes.paParametrosIniciales(4)
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            
            
            loComandoSeleccionar.AppendLine("WITH curTemporal AS ( ")
            loComandoSeleccionar.AppendLine("SELECT		")
            If lcParametro4Desde = "Si"	 Then
				loComandoSeleccionar.AppendLine(" SUBSTRING(CONVERT(nchar(30), Facturas.Fec_Ini,114),1,2)   AS	Hora, ")
			Else
				loComandoSeleccionar.AppendLine(" SUBSTRING(CONVERT(nchar(30), Facturas.Fec_Ini,114),1,5)   AS	Hora, ")   
			End If
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Can_Art1                       As  Can_Art, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Status									  As  Status, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Mon_Net                        As  Mon_Net ")
            loComandoSeleccionar.AppendLine("FROM		Facturas, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas, ")
            loComandoSeleccionar.AppendLine(" 			Articulos, ")
            loComandoSeleccionar.AppendLine(" 			Vendedores, ")
            loComandoSeleccionar.AppendLine(" 			Clientes ")
            loComandoSeleccionar.AppendLine("WHERE		Facturas.Documento				=   Renglones_Facturas.Documento ")
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Cli				=   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Ven				=   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("		AND Renglones_Facturas.Cod_Art		=   Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("		AND Facturas.Status					IN ('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("		AND Facturas.Fec_Ini		BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia))
            loComandoSeleccionar.AppendLine("			AND " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia))
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Cli      BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1)))
            loComandoSeleccionar.AppendLine("			AND " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1)))
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Ven      BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2)))
            loComandoSeleccionar.AppendLine("			AND " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2)))
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Art      BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3)))
            loComandoSeleccionar.AppendLine("			AND " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3)))
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Suc      BETWEEN " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5)))
            loComandoSeleccionar.AppendLine("			AND " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5)))
             If lcParametro4Desde = "Si"	 Then
				loComandoSeleccionar.AppendLine("GROUP BY SUBSTRING(CONVERT(nchar(30), Facturas.Fec_Ini,114),1,2),Renglones_Facturas.Can_Art1, Renglones_Facturas.Mon_Net,Facturas.Status ")
			Else																								   
				loComandoSeleccionar.AppendLine("GROUP BY SUBSTRING(CONVERT(nchar(30), Facturas.Fec_Ini,114),1,5),Renglones_Facturas.Can_Art1, Renglones_Facturas.Mon_Net,Facturas.Status ")
			End If
            loComandoSeleccionar.AppendLine(" ) ")

            loComandoSeleccionar.AppendLine("SELECT		Hora			AS  Hora, ")
            loComandoSeleccionar.AppendLine(" 			Count(Hora) 	AS  Documentos, ")
            loComandoSeleccionar.AppendLine(" 			SUM(Can_Art)	AS  Can_Art, ")
            loComandoSeleccionar.AppendLine(" 			SUM(Mon_Net)	AS  Mon_Net ")
            loComandoSeleccionar.AppendLine("FROM		curTemporal ")
			loComandoSeleccionar.AppendLine("GROUP BY	Hora ")
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTVentas_Horas1", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTVentas_Horas1.ReportSource = loObjetoReporte

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
' JFP: 07/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  03/07/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' RJG: 09/12/10: Ajustado el estatus de las facturas de venta en el filtro.					' 
'-------------------------------------------------------------------------------------------'
' MAT: 21/02/11: Ajuste del Select															' 
'-------------------------------------------------------------------------------------------'
