'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPropiedades_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class fPropiedades_Articulos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            Dim lcTipo As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo & "Clientes")


            loComandoSeleccionar.AppendLine(" SELECT")
			loComandoSeleccionar.AppendLine(" 			Articulos.Cod_Art,")
			loComandoSeleccionar.AppendLine(" 			Articulos.Nom_Art,")
			loComandoSeleccionar.AppendLine(" 			Propiedades.Nom_Pro,")
			loComandoSeleccionar.AppendLine("			Campos_Propiedades.Cod_Pro,")
			loComandoSeleccionar.AppendLine("			Campos_Propiedades.Tip_Pro,")
			loComandoSeleccionar.AppendLine("			Campos_Propiedades.Val_Log,")
			loComandoSeleccionar.AppendLine("			Campos_Propiedades.Val_Num,")
			loComandoSeleccionar.AppendLine("			Campos_Propiedades.Val_Car,")
			loComandoSeleccionar.AppendLine("			Campos_Propiedades.Val_Fec,")
			loComandoSeleccionar.AppendLine("			Campos_Propiedades.Val_Mem")
			loComandoSeleccionar.AppendLine(" FROM Articulos")
			loComandoSeleccionar.AppendLine(" JOIN Campos_Propiedades ON (Campos_Propiedades.Cod_Reg = Articulos.Cod_Art)")
			loComandoSeleccionar.AppendLine(" JOIN Propiedades ON (Propiedades.Cod_Pro = Campos_Propiedades.Cod_Pro)")
			loComandoSeleccionar.AppendLine(" WHERE")
			loComandoSeleccionar.AppendLine("		Campos_Propiedades.Origen = 'Articulos'")   			
            loComandoSeleccionar.AppendLine("       AND" & cusAplicacion.goFormatos.pcCondicionPrincipal)
	


            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPropiedades_Articulos", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfPropiedades_Articulos.ReportSource = loObjetoReporte

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
' RAC: CODIGO INICIAL
'-------------------------------------------------------------------------------------------'
' RAC: 24/03/2011 Modificacion en la ubicacion de la etiqueta del campo Memo en el archivo rpt.
'-------------------------------------------------------------------------------------------'