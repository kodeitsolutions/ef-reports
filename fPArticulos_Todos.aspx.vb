'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPArticulos_Todos"
'-------------------------------------------------------------------------------------------'
Partial Class fPArticulos_Todos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            'Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            'Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            'Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            'Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            'Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            'Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            'Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            'Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            'Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            'Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            'Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            'Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            'Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            'Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcComandoSeleccionar As New StringBuilder()

            lcComandoSeleccionar.AppendLine(" SELECT    Articulos.Cod_Art, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Precio1, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Precio2, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Precio3, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Precio4, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Precio5, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cod_Dep, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cod_Sec, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cod_Mar, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("			Departamentos.Nom_Dep ")
            lcComandoSeleccionar.AppendLine(" FROM		Articulos, ")
            lcComandoSeleccionar.AppendLine("			Departamentos ")
            lcComandoSeleccionar.AppendLine(" WHERE		Articulos.Cod_Dep           =   Departamentos.Cod_Dep ")
            'lcComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Sec       =   Secciones.Cod_Sec ")
            'lcComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Mar       =   Marcas.Cod_Mar ")
            'lcComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Tip       =   Tipos_Articulos.Cod_Tip ")
            'lcComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Cla       =   Clases_Articulos.Cod_Cla ")
            'lcComandoSeleccionar.AppendLine("          AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            'lcComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Art       Between " & lcParametro0Desde)
            'lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            'lcComandoSeleccionar.AppendLine(" 			AND Articulos.Status        IN  (" & lcParametro1Desde & ")")
            'lcComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Dep       Between " & lcParametro2Desde)
            'lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            'lcComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Sec       Between " & lcParametro3Desde)
            'lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            'lcComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Mar       Between " & lcParametro4Desde)
            'lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            'lcComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Tip       Between " & lcParametro5Desde)
            'lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
            'lcComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Cla       Between " & lcParametro6Desde)
            'lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
            lcComandoSeleccionar.AppendLine(" ORDER BY  Articulos.Cod_Dep, Articulos.Cod_Art")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

			'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPArticulos_Todos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfPArticulos_Todos.ReportSource = loObjetoReporte

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
' JJD: 07/01/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT:  19/04/11 : Ajuste de la vista de diseño.
'-------------------------------------------------------------------------------------------'