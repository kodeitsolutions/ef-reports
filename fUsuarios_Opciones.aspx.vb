'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fUsuarios_Opciones"
'-------------------------------------------------------------------------------------------'
Partial Class fUsuarios_Opciones

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		U.Cod_usu COLLATE DATABASE_DEFAULT AS Cod_Usu,")
            loComandoSeleccionar.AppendLine("			U.Nom_Usu COLLATE DATABASE_DEFAULT AS Nom_Usu")
            loComandoSeleccionar.AppendLine("INTO		#tmpUsuario")
            loComandoSeleccionar.AppendLine("FROM		Factory_Global.Dbo.Usuarios AS U")
            loComandoSeleccionar.AppendLine("WHERE		" & cusAplicacion.goFormatos.pcCondicionPrincipal.Replace("usuarios.cod_usu", "U.cod_usu"))
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT		O.Cod_Opc		AS Cod_Opc,")
            loComandoSeleccionar.AppendLine("			O.Cod_Usu		AS Cod_Usu,")
            loComandoSeleccionar.AppendLine("			U.Nom_Usu		As Nom_Usu,")
            loComandoSeleccionar.AppendLine("			O.Nom_Opc		AS Nom_Opc,")
            loComandoSeleccionar.AppendLine("			O.Status		AS Estatus_Opcion,")
            loComandoSeleccionar.AppendLine("			O.Tip_Opc		AS Tip_Opc,")
            loComandoSeleccionar.AppendLine("			(CASE	Tip_Opc  ")
            loComandoSeleccionar.AppendLine("				WHEN 'N' THEN 'Número'")
            loComandoSeleccionar.AppendLine("				WHEN 'C' THEN 'Cadena'")
            loComandoSeleccionar.AppendLine("				WHEN 'F' THEN 'Fecha'")
            loComandoSeleccionar.AppendLine("				WHEN 'M' THEN 'Memo'")
            loComandoSeleccionar.AppendLine("				WHEN 'L' THEN 'Lógico'")
            loComandoSeleccionar.AppendLine("				ELSE '[No Válido]'")
            loComandoSeleccionar.AppendLine("			END) AS Tipo_Opccion, ")
            loComandoSeleccionar.AppendLine("			O.Val_Num		AS Val_Num,")
            loComandoSeleccionar.AppendLine("			O.Val_Log		As Val_Log,")
            loComandoSeleccionar.AppendLine("			O.Val_Car		AS Val_Car,")
            loComandoSeleccionar.AppendLine("			O.Val_Mem		AS Val_Mem,")
            loComandoSeleccionar.AppendLine("			O.Val_Fec		AS Val_Fec,")
            loComandoSeleccionar.AppendLine("			O.Comentario	AS Comentario")
            loComandoSeleccionar.AppendLine("FROM		Opciones_Usuarios AS O")
            loComandoSeleccionar.AppendLine("	JOIN	#tmpUsuario AS U")
            loComandoSeleccionar.AppendLine("		ON	O.Cod_Usu = U.Cod_Usu ")
            loComandoSeleccionar.AppendLine("ORDER BY	U.Cod_Usu, O.Cod_Opc ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpUsuario")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
			
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            If (laDatosReporte.Tables(0).Rows.Count =  0) Then
            
				Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                            "No se Encontraron Registros para los Parámetros Especificados. ", _
                                            vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                            "350px", _
                                            "200px")
            
            End If
                
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fUsuarios_Opciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfUsuarios_Opciones.ReportSource = loObjetoReporte

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
' RJG: 03/05/12: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
