'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDatosBasicos_Combos_iPos"
'-------------------------------------------------------------------------------------------'
Partial Class fDatosBasicos_Combos_iPos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine(" SELECT")				
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Cod_Com,")
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Nom_Com,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Combos_iPos.Status = 'A' THEN 'Activo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Combos_iPos.Status = 'I' THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Combos_iPos.Status = 'S' THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine(" 			END AS Status,")
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Grupo,")
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Tipo,")
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Clase,")
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Cod_Art,")
			loComandoSeleccionar.AppendLine(" 			Articulos.Nom_Art,")
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Can_Art1,")
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Precio1,")
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Mos_Pre,")
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Fila,")
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Columna,")
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Texto,")
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Ayuda,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Combos_iPos.Col_Fon = '#8C4AE6' THEN 'Azul Claro'")
			loComandoSeleccionar.AppendLine(" 				WHEN Combos_iPos.Col_Fon = '#003399' THEN 'Azul Oscuro'")
			loComandoSeleccionar.AppendLine(" 				WHEN Combos_iPos.Col_Fon = '#13920D' THEN 'Verde'")
			loComandoSeleccionar.AppendLine(" 				WHEN Combos_iPos.Col_Fon = '#FFCC00' THEN 'Amarillo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Combos_iPos.Col_Fon = '#E35C2F' THEN 'Rojo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Combos_iPos.Col_Fon = '#FF9933' THEN 'Anaranjado'")
			loComandoSeleccionar.AppendLine(" 			END AS Col_Fon,")		
			loComandoSeleccionar.AppendLine(" 			Combos_iPos.Comentario")
			loComandoSeleccionar.AppendLine(" FROM Combos_iPos")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Articulos ON Articulos.Cod_Art = Combos_iPos.Cod_Art")
			loComandoSeleccionar.AppendLine(" WHERE")	   			
            loComandoSeleccionar.AppendLine("           " & cusAplicacion.goFormatos.pcCondicionPrincipal)


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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDatosBasicos_Combos_iPos", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfDatosBasicos_Combos_iPos.ReportSource = loObjetoReporte

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
' MAT: 04/01/11 : Codigo inicial
'-------------------------------------------------------------------------------------------'
