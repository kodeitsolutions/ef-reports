'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPArticulos_Existencias"
'-------------------------------------------------------------------------------------------'
Partial Class fPArticulos_Existencias
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            'Dim lcOrdenamiento As String = cusAplicacion.goOpciones.goFormatos.pcOrden

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
            lcComandoSeleccionar.AppendLine(" 			AND Articulos.Exi_Act1      >   0 ")
            lcComandoSeleccionar.AppendLine(" ORDER BY  Articulos.Cod_Dep, Articulos.Cod_Art")

            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(lcComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return

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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPArticulos_Existencias", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfPArticulos_Existencias.ReportSource = loObjetoReporte

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