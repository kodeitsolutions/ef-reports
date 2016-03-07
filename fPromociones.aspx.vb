'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPromociones"
'-------------------------------------------------------------------------------------------'
Partial Class fPromociones

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Promociones.Documento, ")
            loComandoSeleccionar.AppendLine("           Promociones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Promociones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Promociones.Status, ")
            loComandoSeleccionar.AppendLine("           Promociones.Tipo, ")
            loComandoSeleccionar.AppendLine("           Promociones.Clase, ")
			loComandoSeleccionar.AppendLine("           Promociones.Referencia, ")
			loComandoSeleccionar.AppendLine("           Promociones.Cod_Mon, ")
			loComandoSeleccionar.AppendLine("           Monedas.Nom_Mon, ")
			loComandoSeleccionar.AppendLine("           Promociones.Tasa, ")
			loComandoSeleccionar.AppendLine("           Promociones.Comentario, ")
            loComandoSeleccionar.AppendLine("           Renglones_Promociones.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Promociones.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Promociones.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Promociones.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Promociones.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Promociones.Comentario AS Comentario_renglon ")
            loComandoSeleccionar.AppendLine(" FROM  Promociones ")
            loComandoSeleccionar.AppendLine(" JOIN  Renglones_Promociones ON  Promociones.Documento  =   Renglones_Promociones.Documento")
            loComandoSeleccionar.AppendLine(" JOIN  Monedas ON  Monedas.Cod_Mon = Promociones.Cod_Mon")
            loComandoSeleccionar.AppendLine(" JOIN  Articulos ON Articulos.Cod_Art       =   Renglones_Promociones.Cod_Art")
            loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPromociones", laDatosReporte)
           
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfPromociones.ReportSource = loObjetoReporte

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
' MAT: 23/08/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
