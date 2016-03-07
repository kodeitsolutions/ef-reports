'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fSolicitudes_Requerimientos"
'-------------------------------------------------------------------------------------------'
Partial Class fSolicitudes_Requerimientos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            
            loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" SELECT CAST(Solicitudes.Seguimientos as XML) AS seguimiento INTO #xmlData from Solicitudes")
            loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" ")	
            
			loComandoSeleccionar.AppendLine(" SELECT    '1' AS Tabla, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Documento, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Status, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Asunto, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Requerimiento, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Tip_Sol, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Etapa, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Fec_Ini, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Fec_Fin, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Fec_Rec, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Cod_Usu, ")
			'loComandoSeleccionar.AppendLine("           Factory_Global.dbo.Usuarios.Nom_Usu    AS  Nom_Usu, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Tipo, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Clase, ")
			loComandoSeleccionar.AppendLine("		    Solicitudes.Cod_Reg, ")
			loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Cod_Con, ")
			loComandoSeleccionar.AppendLine("           Contactos.Nom_Con, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Telefonos, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Correo, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Cod_Ven, ")
			loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
			loComandoSeleccionar.AppendLine("           Solicitudes.Comentario AS Comentarios, ")
			loComandoSeleccionar.AppendLine(" 			Solicitudes.Facturar,")
			loComandoSeleccionar.AppendLine(" 			Solicitudes.Cod_Mon,")
			loComandoSeleccionar.AppendLine(" 			Monedas.Nom_Mon,")
			loComandoSeleccionar.AppendLine(" 			Solicitudes.Tasa,")
			loComandoSeleccionar.AppendLine(" 			Solicitudes.Precio,")
			loComandoSeleccionar.AppendLine(" 			Solicitudes.Costo,")	
			loComandoSeleccionar.AppendLine(" 			Solicitudes.Prioridad,")
			loComandoSeleccionar.AppendLine(" 			Solicitudes.Nivel,")	
			loComandoSeleccionar.AppendLine(" 			NULL as se_renglon,")	
			loComandoSeleccionar.AppendLine(" 			NULL as se_fecha,")	
			loComandoSeleccionar.AppendLine(" 			NULL as se_contacto,")	
			loComandoSeleccionar.AppendLine(" 			NULL as se_accion,")
			loComandoSeleccionar.AppendLine(" 			NULL as se_medio,")
			loComandoSeleccionar.AppendLine(" 			NULL as se_comentario,")
			loComandoSeleccionar.AppendLine(" 			NULL as se_prioridad,")
			loComandoSeleccionar.AppendLine(" 			NULL as se_etapa,")
			loComandoSeleccionar.AppendLine(" 			NULL as se_usuario")
			loComandoSeleccionar.AppendLine(" FROM      Solicitudes")
            loComandoSeleccionar.AppendLine(" JOIN Clientes ON (Solicitudes.Cod_Reg  =   Clientes.Cod_Cli) ")
            loComandoSeleccionar.AppendLine(" JOIN Contactos ON Contactos.Cod_Con = Solicitudes.Cod_Con AND Contactos.Tipo = '" & goEmpresa.pcCodigo & "Clientes'")
            loComandoSeleccionar.AppendLine(" JOIN Monedas ON (Monedas.Cod_Mon = Solicitudes.Cod_Mon)")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores ON (Solicitudes.Cod_Ven   =   Vendedores.Cod_Ven) ")
           'loComandoSeleccionar.AppendLine(" WHERE	(Factory_Global.dbo.Usuarios.Cod_Usu COLLATE Modern_Spanish_CI_AS = Solicitudes.Cod_Usu COLLATE  Modern_Spanish_CI_AS)")
            loComandoSeleccionar.AppendLine("WHERE		 " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" UNION ALL ")	
			loComandoSeleccionar.AppendLine(" ")	
			
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT    '2' AS Tabla, ")
			loComandoSeleccionar.AppendLine("           NULL AS Documento, ")
			loComandoSeleccionar.AppendLine("           NULL AS Status, ")
			loComandoSeleccionar.AppendLine("           NULL AS Asunto, ")
			loComandoSeleccionar.AppendLine("           NULL AS Requerimiento, ")
			loComandoSeleccionar.AppendLine("           NULL AS Tip_Sol, ")
			loComandoSeleccionar.AppendLine("           NULL AS Etapa, ")
			loComandoSeleccionar.AppendLine("           NULL AS Fec_Ini, ")
			loComandoSeleccionar.AppendLine("           NULL AS Fec_Fin, ")
			loComandoSeleccionar.AppendLine("           NULL AS Fec_Rec, ")
			loComandoSeleccionar.AppendLine("           NULL AS Cod_Usu, ")
			'loComandoSeleccionar.AppendLine("           NULL AS Nom_Usu, ")
			loComandoSeleccionar.AppendLine("           NULL AS Tipo, ")
			loComandoSeleccionar.AppendLine("           NULL AS Clase, ")
			loComandoSeleccionar.AppendLine("		    NULL AS Cod_Cod_Reg, ")
			loComandoSeleccionar.AppendLine("           NULL AS Nom_Cli, ")
			loComandoSeleccionar.AppendLine("           NULL AS Cod_Con, ")
			loComandoSeleccionar.AppendLine("           NULL AS Nom_Con, ")
			loComandoSeleccionar.AppendLine("           NULL AS Telefonos, ")
			loComandoSeleccionar.AppendLine("           NULL AS Correo, ")
			loComandoSeleccionar.AppendLine("           NULL AS Cod_Ven, ")
			loComandoSeleccionar.AppendLine("           NULL AS Nom_Ven, ")
			loComandoSeleccionar.AppendLine("           NULL AS Comentario, ")
			loComandoSeleccionar.AppendLine(" 			NULL AS Facturar,")
			loComandoSeleccionar.AppendLine(" 			NULL AS Cod_Mon,")
			loComandoSeleccionar.AppendLine(" 			NULL AS Nom_Mon,")
			loComandoSeleccionar.AppendLine(" 			NULL AS Tasa,")
			loComandoSeleccionar.AppendLine(" 			NULL AS Precio,")
			loComandoSeleccionar.AppendLine(" 			NULL AS Costo,")	
			loComandoSeleccionar.AppendLine(" 			NULL AS Prioridad,")
			loComandoSeleccionar.AppendLine(" 			NULL AS Nivel,")
			loComandoSeleccionar.AppendLine(" 			ROW_NUMBER() OVER (ORDER BY D.C.value('@status', 'Varchar(15)') DESC) as se_renglon,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@fecha', 'datetime') as se_fecha,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@contacto', 'Varchar(300)') as se_contacto,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@accion', 'Varchar(300)') as se_accion,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@medio', 'Varchar(300)') as se_medio,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@comentario', 'Varchar(5000)') as se_comentario,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@prioridad', 'Varchar(300)') as se_prioridad,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@etapa', 'Varchar(300)') as se_etapa,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@usuario', 'Varchar(100)') as se_usuario")	
			loComandoSeleccionar.AppendLine(" FROM #xmlData")	
			loComandoSeleccionar.AppendLine(" CROSS APPLY seguimiento.nodes('elementos/elemento') D(c) -- recuerda que seguimiento es el nombre del campo donde esta el XML")	
			loComandoSeleccionar.AppendLine(" ORDER BY tabla, Solicitudes.documento DESC ,se_renglon")
			
			
			loComandoSeleccionar.AppendLine(" DROP TABLE #xmlData")
								  
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
					 
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fSolicitudes_Requerimientos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfSolicitudes_Requerimientos.ReportSource = loObjetoReporte

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
' MAT: 11/08/11: Programacion inicial
'-------------------------------------------------------------------------------------------'