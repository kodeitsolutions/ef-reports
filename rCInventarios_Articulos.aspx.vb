'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCInventarios_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rCInventarios_Articulos 
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden



        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))


            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT      Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("             Cortes.Status, ")
            loComandoSeleccionar.AppendLine("             Cortes.Documento, ")
            loComandoSeleccionar.AppendLine("             Cortes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("             Renglones_Cortes.Renglon, ")
            loComandoSeleccionar.AppendLine("             Renglones_Cortes.Cod_Art, ")
            loComandoSeleccionar.AppendLine("             Renglones_Cortes.Can_Teo, ")
            loComandoSeleccionar.AppendLine("             Renglones_Cortes.Can_Rea, ")
            loComandoSeleccionar.AppendLine("             Renglones_Cortes.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("             Renglones_Cortes.Cos_Ult1 As Costo, ")
            loComandoSeleccionar.AppendLine("             (Renglones_Cortes.Cos_Ult1 * Renglones_Cortes.Can_Rea) AS Mon_Aju ")
            loComandoSeleccionar.AppendLine(" FROM        Articulos, ")
            loComandoSeleccionar.AppendLine("             Cortes, ")
            loComandoSeleccionar.AppendLine("             Renglones_Cortes ")
            loComandoSeleccionar.AppendLine(" WHERE       Cortes.Documento				=	Renglones_Cortes.Documento	")
            loComandoSeleccionar.AppendLine("             And Articulos.Cod_Art			=	Renglones_Cortes.Cod_Art ")
            loComandoSeleccionar.AppendLine("             And Renglones_Cortes.Cod_Art	Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("             And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("             And Cortes.Documento			Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("             And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("             And Cortes.Fec_Ini			Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("             And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("             And Cortes.Status				IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("             And Cortes.Cod_Alm			Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("             And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("      	      AND Cortes.Cod_Rev            Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 		      AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("      	      AND Cortes.Cod_Suc            Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 		      AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY    " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString(), "curReportes")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            
            
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
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCInventarios_Articulos", laDatosReporte)
            
			Me.mTraducirReporte(loObjetoReporte)
          
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrCInventarios_Articulos.ReportSource = loObjetoReporte


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
' JJD: 11/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 14/05/09: Filtro “Revisión:”, Estadarización del código
'-------------------------------------------------------------------------------------------'
' AAP: 29/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' AAP: 02/07/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' JJD: 30/03/10: Ajuste a la seleccion de los parametros
'-------------------------------------------------------------------------------------------'
' MAT: 22/02/11: Ajuste del Select y la presentación del reporte
'-------------------------------------------------------------------------------------------'