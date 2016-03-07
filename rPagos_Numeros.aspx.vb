'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPagos_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rPagos_Numeros

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
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosFinales(9)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.appendline(" SELECT    Pagos.Documento, ")
            loComandoSeleccionar.appendline("           Pagos.Status, ")
            loComandoSeleccionar.appendline("           Pagos.Fec_Ini, ")
            loComandoSeleccionar.appendline("           Pagos.Cod_Pro, ") 
            loComandoSeleccionar.appendline("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.appendline("           Pagos.Mon_Net   AS  Neto, ") 
            loComandoSeleccionar.appendline("           Pagos.Cod_Mon, ")
            loComandoSeleccionar.appendline("           Detalles_Pagos.Renglon, ")
            loComandoSeleccionar.appendline("           Detalles_Pagos.Tip_Ope, ")
            'loComandoSeleccionar.appendline("           Detalles_Pagos.Num_Doc, ")
            loComandoSeleccionar.appendline("           Detalles_Pagos.Cod_Ban, ")
            'loComandoSeleccionar.appendline("           Detalles_Pagos.Cod_Caj, ")
            'loComandoSeleccionar.appendline("           Detalles_Pagos.Cod_Cue, ")
            
            loComandoSeleccionar.AppendLine("           ISNULL((select nom_caj from cajas where Cajas.Cod_Caj = Detalles_Pagos.Cod_Caj),'') AS Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Num_Doc,  ")
            loComandoSeleccionar.AppendLine("           ISNULL((Select Num_Cue From Cuentas_Bancarias Where Cuentas_Bancarias.Cod_Cue = Detalles_Pagos.Cod_Cue),'') As Cod_Cue,  ")
            
            loComandoSeleccionar.appendline("           Detalles_Pagos.Cod_Tar, ")
            loComandoSeleccionar.appendline("           Detalles_Pagos.Mon_Net ")
            loComandoSeleccionar.appendline(" FROM      Pagos, ")
            loComandoSeleccionar.appendline("           Detalles_Pagos, ")
            loComandoSeleccionar.appendline("           Proveedores ")
            loComandoSeleccionar.appendline(" WHERE     Pagos.Documento         =   Detalles_Pagos.Documento ")
            loComandoSeleccionar.appendline("           AND Pagos.Cod_Pro       =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.appendline("           AND Pagos.Documento     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.appendline("           AND " & lcParametro1Hasta )
            loComandoSeleccionar.appendline("           AND Pagos.Fec_Ini       BETWEEN " & lcParametro1Desde )
            loComandoSeleccionar.appendline("           AND " & lcParametro1Hasta )
            loComandoSeleccionar.appendline("           AND Pagos.Cod_Pro       BETWEEN " & lcParametro2Desde )
            loComandoSeleccionar.appendline("           AND " & lcParametro2Hasta )
            loComandoSeleccionar.appendline("           AND Pagos.Cod_Mon       BETWEEN " & lcParametro3Desde )
            loComandoSeleccionar.appendline("           AND " & lcParametro3Hasta )
            loComandoSeleccionar.appendline("           AND Pagos.Status        IN ( " & lcParametro4Desde & " ) ")

            
            loComandoSeleccionar.AppendLine("           And Detalles_Pagos.Cod_Caj    Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Detalles_Pagos.Cod_Cue    Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           And Pagos.Cod_Suc    Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            
            If lcParametro9Desde = "Igual" Then
                loComandoSeleccionar.AppendLine("           And Pagos.Cod_rev   Between " & lcParametro8Desde)
            Else
                loComandoSeleccionar.AppendLine("           And Pagos.Cod_rev  Not Between " & lcParametro8Desde)
            End If

            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)

            'loComandoSeleccionar.appendline(" ORDER BY  Detalles_Pagos.Documento, Detalles_Pagos.Tip_Ope, Detalles_Pagos.Renglon ")
            loComandoSeleccionar.AppendLine("ORDER BY    " & lcOrdenamiento & ", Detalles_Pagos.Tip_Ope, Detalles_Pagos.Renglon ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rPagos_Numeros", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPagos_Numeros.ReportSource =	 loObjetoReporte

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
' JJD: 09/10/08: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' CMS: 22/04/09: Se completa la estandarización del código									'
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"														'
'-------------------------------------------------------------------------------------------'
' RJG:  10/04/12: Se agregó el total de registros.											'
'-------------------------------------------------------------------------------------------'
