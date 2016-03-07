'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fListas_Desplegables"
'-------------------------------------------------------------------------------------------'
Partial Class fListas_Desplegables

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT Factory_Global.dbo.Combos.cod_com,")
            loConsulta.AppendLine("		(CASE Factory_Global.dbo.Combos.status")
            loConsulta.AppendLine("			WHEN 'A' THEN 'ACTIVO'")
            loConsulta.AppendLine("			WHEN 'I' THEN 'INACTIVO'")
            loConsulta.AppendLine("			WHEN 'S' THEN 'SUSPENDIDO'")
            loConsulta.AppendLine("		END) AS status,")
            loConsulta.AppendLine("		Factory_Global.dbo.Combos.nom_com,")
            loConsulta.AppendLine("		Factory_Global.dbo.combos.clave,")
            loConsulta.AppendLine("		Factory_Global.dbo.Combos.comentario,")
            loConsulta.AppendLine("		Renglones.Cla,")
            loConsulta.AppendLine("		Renglones.Valor,")
            loConsulta.AppendLine("		Renglones.Campo")
            loConsulta.AppendLine("FROM Factory_Global.dbo.combos")
            loConsulta.AppendLine("	LEFT JOIN (	    SELECT ")
            loConsulta.AppendLine("						    Lista.Cod_Com,")
            loConsulta.AppendLine("						    T.C.value('(@clave)[1]', 'varchar(max)')            AS Cla,")
            loConsulta.AppendLine("						    T.C.value('(@valor)[1]', 'varchar(max)')            AS valor,")
            loConsulta.AppendLine("					    	T.C.value('(text())[1]', 'varchar(max)')            AS Campo")
            loConsulta.AppendLine("				    FROM(SELECT Cod_Com,Cast(Contenido AS XML) AS Contenido FROM Factory_Global.dbo.combos)")
            loConsulta.AppendLine("				    	AS Lista")
            loConsulta.AppendLine("					    OUTER APPLY Contenido.nodes('//elementos/elemento') AS T(C)")
            loConsulta.AppendLine("				) AS Renglones ON  Renglones.Cod_Com = Factory_Global.dbo.Combos.Cod_Com")
            loConsulta.AppendLine("WHERE        " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fListas_Desplegables", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfListas_Desplegables.ReportSource = loObjetoReporte

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
' EAG: 25/08/15: Programacion inicial
'-------------------------------------------------------------------------------------------'
