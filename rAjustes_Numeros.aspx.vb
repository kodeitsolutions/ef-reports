'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAjustes_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rAjustes_Numeros
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("  SELECT		Ajustes.Documento, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes.Comentario, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes.Status, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes.Cod_Mon, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes.Tip_Ori, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes.Doc_Ori, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes.Mon_Net, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes.Can_Art1 ")
            loComandoSeleccionar.AppendLine(" FROM			Ajustes ")
            loComandoSeleccionar.AppendLine(" WHERE			Ajustes.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				And Ajustes.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				And Ajustes.Status IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine("               And Ajustes.Tipo		IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("      			And Ajustes.Cod_Rev between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("      			And Ajustes.Cod_Suc between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
           ' Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAjustes_Numeros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrAjustes_Numeros.ReportSource = loObjetoReporte

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
' JJD: 25/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  16/03/09: Estandarizacion de codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  12/05/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  29/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  27/03/09: Filtro Tipo de Ajuste, se aplico el metodo de validacion de registro cero
'-------------------------------------------------------------------------------------------'