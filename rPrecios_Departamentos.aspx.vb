'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPrecios_Departamentos"
'-------------------------------------------------------------------------------------------'
Partial Class rPrecios_Departamentos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            
            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Web, ")
            loComandoSeleccionar.AppendLine("           Articulos.Exi_Act1, ")
            loComandoSeleccionar.AppendLine("           Articulos.Exi_Ped1, ")
            
             Select Case lcParametro8Desde
                Case "Si"
                    loComandoSeleccionar.AppendLine("	'Si' AS Disponible,")
				Case "No"
                    loComandoSeleccionar.AppendLine("	'No' AS Disponible,")
            End Select
            
            Select Case lcParametro9Desde
                Case "Precio1"
                    loComandoSeleccionar.AppendLine("	ISNULL (Articulos.Precio1,0) AS Precio,")
				Case "Precio2"
                    loComandoSeleccionar.AppendLine("	ISNULL (Articulos.Precio2,0) AS Precio,")
                Case "Precio3"
                    loComandoSeleccionar.AppendLine("	ISNULL (Articulos.Precio3,0) AS Precio,")
                Case "Precio4"
                    loComandoSeleccionar.AppendLine("	ISNULL (Articulos.Precio4,0) AS Precio,")
                Case "Precio5"
                    loComandoSeleccionar.AppendLine("	ISNULL (Articulos.Precio5,0) AS Precio,")      
            End Select
            
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Dep, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Sec, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Cla, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Mar, ")
            loComandoSeleccionar.AppendLine("           Departamentos.Nom_Dep ")
            loComandoSeleccionar.AppendLine(" FROM      Articulos, ")
            loComandoSeleccionar.AppendLine("           Departamentos, ")
            loComandoSeleccionar.AppendLine("           Secciones, ")
            loComandoSeleccionar.AppendLine("           Marcas, ")
            loComandoSeleccionar.AppendLine("           Tipos_Articulos, ")
            loComandoSeleccionar.AppendLine("           Clases_Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Articulos.Cod_Dep           =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec       =   Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine("           And Departamentos.Cod_Dep   =   Secciones.Cod_Dep ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar       =   Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Tip       =   Tipos_Articulos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Cla       =   Clases_Articulos.Cod_Cla ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art       Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Status        IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep       Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec       Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar       Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Tip       Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Cla       Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Ubi    Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPrecios_Departamentos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPrecios_Departamentos.ReportSource = loObjetoReporte


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
        
			  Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")
        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' MAT: 24/01/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 01/03/11: Hipervinculo para la pág. Web de cada artículo
'-------------------------------------------------------------------------------------------'

