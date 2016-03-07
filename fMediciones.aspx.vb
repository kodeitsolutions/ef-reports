'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fMediciones"
'-------------------------------------------------------------------------------------------'
Partial Class fMediciones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()
			
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT  Mediciones.Documento                                                            AS Documento, ")
			loConsulta.AppendLine("        Mediciones.Status                                                               AS Status, ")
			loConsulta.AppendLine("        Mediciones.Posicion                                                             AS Posicion, ")
			loConsulta.AppendLine("        (CASE WHEN Mediciones.Cod_Art = ''")
			loConsulta.AppendLine("            THEN '[SIN DEFINIR]'")
			loConsulta.AppendLine("            ELSE RTRIM(Mediciones.Cod_Art) ")
			loConsulta.AppendLine("                + ' - '")
			loConsulta.AppendLine("                + COALESCE(Articulos.Nom_Art, '[N/A]')")
			loConsulta.AppendLine("        END)                                                                            AS Articulo, ")
			loConsulta.AppendLine("        (CASE WHEN Mediciones.Cod_Alm = ''")
			loConsulta.AppendLine("            THEN '[SIN DEFINIR]'")
			loConsulta.AppendLine("            ELSE RTRIM(Mediciones.Cod_Alm) ")
			loConsulta.AppendLine("                + ' - '")
			loConsulta.AppendLine("                + COALESCE(Almacenes.Nom_Alm, '[N/A]')")
			loConsulta.AppendLine("        END)                                                                            AS Almacen, ")
			loConsulta.AppendLine("        Mediciones.Cod_Reg                                                              AS Documento_Origen,")
			loConsulta.AppendLine("        Mediciones.Origen                                                               AS Tipo_Origen,")
			loConsulta.AppendLine("        Mediciones.Control                                                              AS Control,")
			loConsulta.AppendLine("        Mediciones.Referencia                                                           AS Referencia,")
			loConsulta.AppendLine("        Mediciones.Fec_Ini                                                              AS Fec_Ini,")
			loConsulta.AppendLine("        Mediciones.Comentario                                                           AS Comentario,")
			loConsulta.AppendLine("        (CASE WHEN Mediciones.Cod_Usu_Ela = '' ")
			loConsulta.AppendLine("            THEN ''")
			loConsulta.AppendLine("            ELSE RTRIM(Mediciones.Cod_Usu_Ela)")
			loConsulta.AppendLine("            + ' - ' ")
			loConsulta.AppendLine("            + COALESCE(Elaborado.Nom_Usu COLLATE DATABASE_DEFAULT, '[N/A]') ")
			loConsulta.AppendLine("        END)                                                                            AS Elaborado_Por,")
			loConsulta.AppendLine("        (CASE WHEN Mediciones.Cod_Usu_Rev = '' ")
			loConsulta.AppendLine("            THEN ''")
			loConsulta.AppendLine("            ELSE RTRIM(Mediciones.Cod_Usu_Rev)")
			loConsulta.AppendLine("            + ' - ' ")
			loConsulta.AppendLine("            + COALESCE(Revisado.Nom_Usu COLLATE DATABASE_DEFAULT, '[N/A]') ")
			loConsulta.AppendLine("        END)                                                                            AS Revisado_Por,")
			loConsulta.AppendLine("        (CASE WHEN Mediciones.Cod_Usu_Apr = '' ")
			loConsulta.AppendLine("            THEN ''")
			loConsulta.AppendLine("            ELSE RTRIM(Mediciones.Cod_Usu_Apr)")
			loConsulta.AppendLine("            + ' - ' ")
			loConsulta.AppendLine("            + COALESCE(Aprobado.Nom_Usu COLLATE DATABASE_DEFAULT, '[N/A]') ")
			loConsulta.AppendLine("        END)                                                                            AS Aprovado_Por,")
			loConsulta.AppendLine("        Mediciones.Resumen                                                              AS Resumen,")
			loConsulta.AppendLine("        Mediciones.Presentacion                                                         AS Presentacion,")
			loConsulta.AppendLine("        Mediciones.Con_Man                                                              AS Con_Man,")
			loConsulta.AppendLine("        Mediciones.Doc_Req                                                              AS Doc_Req,")
			loConsulta.AppendLine("        Mediciones.Descripcion                                                          AS Descripcion,")
			loConsulta.AppendLine("        Mediciones.Dec_Cal                                                              AS Dec_Cal,")
			loConsulta.AppendLine("        -- RENGLONES")
			loConsulta.AppendLine("        Renglones_Mediciones.Renglon                                                    AS Renglon, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Cod_Var                                                    AS Cod_Var, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Nom_Var                                                    AS Nom_Var, ")
			loConsulta.AppendLine("        COALESCE(Variables.Tip_Var, 'Numerico')                                         AS Tip_Var, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Cod_Uni                                                    AS Cod_Uni, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Val_Min_Esp                                                AS Val_Min_Esp, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Val_Max_Esp                                                AS Val_Max_Esp, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Res_Num                                                    AS Res_Num, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Res_Car                                                    AS Res_Car, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Res_Log                                                    AS Res_Log, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Res_Fec                                                    AS Res_Fec, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Res_Mem                                                    AS Res_Mem, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Evaluacion                                                 AS Evaluacion, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Clase                                                      AS Clase, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Revisado                                                   AS Revisado, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Auditado                                                   AS Auditado, ")
			loConsulta.AppendLine("        Renglones_Mediciones.Comentario                                                 AS Comentario_Renglon")
			loConsulta.AppendLine("FROM    Mediciones")
			loConsulta.AppendLine("    LEFT JOIN Articulos ON Articulos.Cod_Art = Mediciones.Cod_Art")
			loConsulta.AppendLine("    LEFT JOIN Almacenes ON Almacenes.Cod_Alm = Mediciones.Cod_Alm")
			loConsulta.AppendLine("    LEFT JOIN Factory_Global.dbo.Usuarios Elaborado ")
			loConsulta.AppendLine("        ON Elaborado.Cod_Usu COLLATE DATABASE_DEFAULT = Mediciones.Cod_Usu_Ela COLLATE DATABASE_DEFAULT")
			loConsulta.AppendLine("    LEFT JOIN Factory_Global.dbo.Usuarios Revisado ")
			loConsulta.AppendLine("        ON Revisado.Cod_Usu COLLATE DATABASE_DEFAULT = Mediciones.Cod_Usu_Rev COLLATE DATABASE_DEFAULT")
			loConsulta.AppendLine("    LEFT JOIN Factory_Global.dbo.Usuarios Aprobado ")
			loConsulta.AppendLine("        ON Aprobado.Cod_Usu COLLATE DATABASE_DEFAULT = Mediciones.Cod_Usu_Apr COLLATE DATABASE_DEFAULT")
			loConsulta.AppendLine("    JOIN Renglones_Mediciones ")
			loConsulta.AppendLine("        ON Renglones_Mediciones.Documento = Mediciones.Documento")
			loConsulta.AppendLine("        AND Renglones_Mediciones.Adicional = Mediciones.Adicional")
			loConsulta.AppendLine("    LEFT JOIN Variables ")
			loConsulta.AppendLine("        ON Variables.Cod_Var = Renglones_Mediciones.Cod_Var")
            loConsulta.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loConsulta.AppendLine("ORDER BY Renglon ASC")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
		    
            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fMediciones", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
            Me.mFormatearCamposReporte(loObjetoReporte)
            
            Me.crvfMediciones.ReportSource = loObjetoReporte

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
' Fin del codigo.                                                                           '
'-------------------------------------------------------------------------------------------'
' RJG: 15/01/15: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
