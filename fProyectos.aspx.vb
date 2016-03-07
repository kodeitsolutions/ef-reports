'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fProyectos"
'-------------------------------------------------------------------------------------------'
Partial Class fProyectos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            
            loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine("SELECT CAST(Proyectos.Seguimientos as XML) AS seguimiento,Proyectos.Cod_Pro")
			loComandoSeleccionar.AppendLine("INTO #xmlDataSeguimiento") 
			loComandoSeleccionar.AppendLine("FROM Proyectos")
			loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" ")
			
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine("SELECT CAST(Proyectos.Fases as XML) AS fase,Proyectos.Cod_Pro")
			loComandoSeleccionar.AppendLine("INTO #xmlDataFases") 
			loComandoSeleccionar.AppendLine("FROM Proyectos")
			loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" ")
			
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine("SELECT	Cod_Pro,")
			loComandoSeleccionar.AppendLine("		ROW_NUMBER() OVER (ORDER BY D.C.value('@status', 'Varchar(15)') DESC) as se_renglon,")
			loComandoSeleccionar.AppendLine("D.C.value('@fecha', 'datetime')			as se_fecha, ")
			loComandoSeleccionar.AppendLine("D.C.value('@contacto', 'Varchar(300)')		as se_contacto,	")
			loComandoSeleccionar.AppendLine("D.C.value('@accion', 'Varchar(300)')		as se_accion,")
			loComandoSeleccionar.AppendLine("D.C.value('@medio', 'Varchar(300)')		as se_medio,")
			loComandoSeleccionar.AppendLine("D.C.value('@comentario', 'Varchar(5000)')	as se_comentario,")
			loComandoSeleccionar.AppendLine("D.C.value('@prioridad', 'Varchar(300)')	as se_prioridad,")
			loComandoSeleccionar.AppendLine("D.C.value('@etapa', 'Varchar(300)')		as se_etapa, ")
			loComandoSeleccionar.AppendLine("D.C.value('@usuario', 'Varchar(100)')		as se_usuario  ")
			loComandoSeleccionar.AppendLine("INTO #tmpSeguimiento")
			loComandoSeleccionar.AppendLine("FROM #xmlDataSeguimiento")
			loComandoSeleccionar.AppendLine("CROSS APPLY seguimiento.nodes('elementos/elemento') D(c) ")
			loComandoSeleccionar.AppendLine("ORDER BY Cod_Pro DESC ,se_renglon   ")
			loComandoSeleccionar.AppendLine(" ")
			
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine("SELECT	Cod_Pro,")
			loComandoSeleccionar.AppendLine("		ROW_NUMBER() OVER (ORDER BY D.C.value('@fase', 'Varchar(50)') DESC) as fa_renglon,")
			loComandoSeleccionar.AppendLine("		D.C.value('@fase', 'Varchar(300)')			as fa_fase,")
			loComandoSeleccionar.AppendLine("		D.C.value('@responsable', 'Varchar(300)')	as fa_responsable,")
			loComandoSeleccionar.AppendLine("		D.C.value('@actividad', 'Varchar(300)')		as fa_actividad,")
			loComandoSeleccionar.AppendLine("		D.C.value('@etapa', 'Varchar(300)')			as fa_etapa,")
			loComandoSeleccionar.AppendLine("		D.C.value('@por_eje', 'Decimal (28,10)')	as fa_porcentaje,")
			loComandoSeleccionar.AppendLine("		D.C.value('@fec_ini', 'datetime')			as fa_inicio,")
			loComandoSeleccionar.AppendLine("		D.C.value('@fec_fin', 'datetime')			as fa_fin,")
			loComandoSeleccionar.AppendLine("		D.C.value('@costos', 'Decimal (28,10)')		as fa_costos,")
			loComandoSeleccionar.AppendLine("		D.C.value('@comentario', 'Varchar(5000)')	as fa_comentario,")
			loComandoSeleccionar.AppendLine("		D.C.value('@usuario', 'Varchar(100)')		as fa_usuario")
			loComandoSeleccionar.AppendLine("INTO #tmpFases")
			loComandoSeleccionar.AppendLine("FROM #xmlDataFases")
			loComandoSeleccionar.AppendLine("CROSS APPLY fase.nodes('elementos/elemento') D(c) -- recuerda que seguimiento es el nombre del campo donde esta el XML")
			loComandoSeleccionar.AppendLine("ORDER BY Cod_Pro DESC ,fa_renglon")
			loComandoSeleccionar.AppendLine(" ")
			
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine("DROP TABLE #xmlDataSeguimiento")
			loComandoSeleccionar.AppendLine("DROP TABLE #xmlDataFases")
			loComandoSeleccionar.AppendLine(" ")
			
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine("SELECT		Proyectos.Cod_Pro,") 
			loComandoSeleccionar.AppendLine("			Proyectos.Nom_Pro,")
			loComandoSeleccionar.AppendLine("			Proyectos.Status, ")
			loComandoSeleccionar.AppendLine("			Proyectos.Responsable,") 
            loComandoSeleccionar.AppendLine("			Proyectos.Objetivo,")
            loComandoSeleccionar.AppendLine("			Proyectos.Resumen,")
			loComandoSeleccionar.AppendLine("			Proyectos.Comentario,") 
			loComandoSeleccionar.AppendLine("			Proyectos.Por_Eje,") 
			loComandoSeleccionar.AppendLine("			Proyectos.Fec_Ini,") 
			loComandoSeleccionar.AppendLine("			Proyectos.Fec_Fin, ")
			loComandoSeleccionar.AppendLine("			Proyectos.Etapa, ")
			loComandoSeleccionar.AppendLine("			Proyectos.Duracion,") 
			loComandoSeleccionar.AppendLine("			Proyectos.Tipo, ")
			loComandoSeleccionar.AppendLine("			Proyectos.Clase,") 
			loComandoSeleccionar.AppendLine("			Proyectos.Cod_Mon,")
			loComandoSeleccionar.AppendLine("			Monedas.Nom_Mon,")
			loComandoSeleccionar.AppendLine("			Proyectos.Tasa,")
			loComandoSeleccionar.AppendLine("			Proyectos.Prioridad,")
			loComandoSeleccionar.AppendLine("			Proyectos.Nivel, ")
			loComandoSeleccionar.AppendLine("			Proyectos.Importancia,   ")
			loComandoSeleccionar.AppendLine("			Proyectos.Logico1,   ")
			loComandoSeleccionar.AppendLine("			Proyectos.Logico2,   ")
			loComandoSeleccionar.AppendLine("			Proyectos.Logico3,   ")
			loComandoSeleccionar.AppendLine("			Proyectos.Logico4,")
			loComandoSeleccionar.AppendLine("			Proyectos.Logico5,")
			loComandoSeleccionar.AppendLine("			Proyectos.Fecha1,")
			loComandoSeleccionar.AppendLine("			Proyectos.Fecha2,")
			loComandoSeleccionar.AppendLine("			Proyectos.Fecha3,")
			loComandoSeleccionar.AppendLine("			Proyectos.Fecha4,")
			loComandoSeleccionar.AppendLine("			Proyectos.Fecha5,")
			loComandoSeleccionar.AppendLine("			Proyectos.Numerico1,")
			loComandoSeleccionar.AppendLine("			Proyectos.Numerico2,")
			loComandoSeleccionar.AppendLine("			Proyectos.Numerico3,")
			loComandoSeleccionar.AppendLine("			Proyectos.Numerico4,")
			loComandoSeleccionar.AppendLine("			Proyectos.Numerico5,")
			loComandoSeleccionar.AppendLine("			Proyectos.Caracter1,")
			loComandoSeleccionar.AppendLine("			Proyectos.Caracter2,")
			loComandoSeleccionar.AppendLine("			Proyectos.Caracter3,")
			loComandoSeleccionar.AppendLine("			Proyectos.Caracter4,")
			loComandoSeleccionar.AppendLine("			Proyectos.Caracter5,")
			loComandoSeleccionar.AppendLine("			Proyectos.Memo1, ")
			loComandoSeleccionar.AppendLine("			Proyectos.Memo2,")
			loComandoSeleccionar.AppendLine("			Proyectos.Memo3,")
			loComandoSeleccionar.AppendLine("			Proyectos.Memo4,")
			loComandoSeleccionar.AppendLine("			Proyectos.Memo5	")
			loComandoSeleccionar.AppendLine("INTO #tmpOriginal")
			loComandoSeleccionar.AppendLine("FROM	Proyectos  ")
			loComandoSeleccionar.AppendLine("JOIN Monedas ON (Monedas.Cod_Mon = Proyectos.Cod_Mon)")
			loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

			loComandoSeleccionar.AppendLine("DECLARE @VariableTabla TABLE (")
			loComandoSeleccionar.AppendLine("Tabla				CHAR(30),")
			loComandoSeleccionar.AppendLine("Cod_Pro			CHAR(10),")
			loComandoSeleccionar.AppendLine("Nom_Pro			CHAR(250),")
			loComandoSeleccionar.AppendLine("Status				CHAR(50),")
            loComandoSeleccionar.AppendLine("Responsable 		CHAR(250),")
            loComandoSeleccionar.AppendLine("Resumen			TEXT,  ")
			loComandoSeleccionar.AppendLine("Objetivo			TEXT,  ")
			loComandoSeleccionar.AppendLine("Comentario			TEXT,  ")
			loComandoSeleccionar.AppendLine("Por_Eje			DECIMAL(28,10),")
			loComandoSeleccionar.AppendLine("Fec_Ini			DATETIME,")
			loComandoSeleccionar.AppendLine("Fec_Fin			DATETIME,")
			loComandoSeleccionar.AppendLine("Etapa				CHAR(50),")
			loComandoSeleccionar.AppendLine("Duracion			DECIMAL(28,10),")
			loComandoSeleccionar.AppendLine("Tipo				CHAR(50),")
			loComandoSeleccionar.AppendLine("Clase				CHAR(50),  ")
			loComandoSeleccionar.AppendLine("Cod_Mon			CHAR(10),  ")
			loComandoSeleccionar.AppendLine("Nom_Mon			CHAR(50),")
			loComandoSeleccionar.AppendLine("Tasa				DECIMAL(28,10),	")
			loComandoSeleccionar.AppendLine("Prioridad			CHAR(50),")
			loComandoSeleccionar.AppendLine("Nivel				DECIMAL(28,10),")
			loComandoSeleccionar.AppendLine("Importancia 		CHAR(50),")
			loComandoSeleccionar.AppendLine("Logico1 			BIT, ")
			loComandoSeleccionar.AppendLine("Logico2 			BIT,")
			loComandoSeleccionar.AppendLine("Logico3 			BIT,")
			loComandoSeleccionar.AppendLine("Logico4 			BIT,")
			loComandoSeleccionar.AppendLine("Logico5 			BIT,")
			loComandoSeleccionar.AppendLine("Fecha1 			DATETIME,")
			loComandoSeleccionar.AppendLine("Fecha2 			DATETIME,")
			loComandoSeleccionar.AppendLine("Fecha3 			DATETIME,")
			loComandoSeleccionar.AppendLine("Fecha4 			DATETIME,")
			loComandoSeleccionar.AppendLine("Fecha5 			DATETIME,")
			loComandoSeleccionar.AppendLine("Numerico1 			DECIMAL(28,10),")
			loComandoSeleccionar.AppendLine("Numerico2 			DECIMAL(28,10),")
			loComandoSeleccionar.AppendLine("Numerico3 			DECIMAL(28,10),")
			loComandoSeleccionar.AppendLine("Numerico4 			DECIMAL(28,10),")
			loComandoSeleccionar.AppendLine("Numerico5 			DECIMAL(28,10),")
			loComandoSeleccionar.AppendLine("Caracter1 			CHAR(250),")
			loComandoSeleccionar.AppendLine("Caracter2 			CHAR(250),")
			loComandoSeleccionar.AppendLine("Caracter3 			CHAR(250),")
			loComandoSeleccionar.AppendLine("Caracter4 			CHAR(250),")
			loComandoSeleccionar.AppendLine("Caracter5 			CHAR(250),")
			loComandoSeleccionar.AppendLine("Memo1 				TEXT,")
			loComandoSeleccionar.AppendLine("Memo2 				TEXT,")
			loComandoSeleccionar.AppendLine("Memo3 				TEXT,")
			loComandoSeleccionar.AppendLine("Memo4 				TEXT,")
			loComandoSeleccionar.AppendLine("Memo5 				TEXT,")
			loComandoSeleccionar.AppendLine("se_renglon			DECIMAL(28,10),")
			loComandoSeleccionar.AppendLine("se_fecha			DATETIME,")
			loComandoSeleccionar.AppendLine("se_contacto			CHAR(50),")
			loComandoSeleccionar.AppendLine("se_accion			CHAR(50),")
			loComandoSeleccionar.AppendLine("se_medio			CHAR(50),")
			loComandoSeleccionar.AppendLine("se_comentario		CHAR(250),")
			loComandoSeleccionar.AppendLine("se_prioridad		CHAR(50),")
			loComandoSeleccionar.AppendLine("se_etapa			CHAR(50),")
			loComandoSeleccionar.AppendLine("se_usuario			CHAR(50),")
			loComandoSeleccionar.AppendLine("fa_fase				CHAR(50),")
			loComandoSeleccionar.AppendLine("fa_responsable		CHAR(50),")
			loComandoSeleccionar.AppendLine("fa_actividad		CHAR(50),")
			loComandoSeleccionar.AppendLine("fa_etapa			CHAR(50),")
			loComandoSeleccionar.AppendLine("fa_porcentaje		DECIMAL(28,10),")
			loComandoSeleccionar.AppendLine("fa_inicio			DATETIME,")
			loComandoSeleccionar.AppendLine("fa_fin				DATETIME,  ")
			loComandoSeleccionar.AppendLine("fa_costos			DECIMAL(28,10),")
			loComandoSeleccionar.AppendLine("fa_comentario		TEXT,  ")
			loComandoSeleccionar.AppendLine("fa_usuario			CHAR(50)")
			loComandoSeleccionar.AppendLine(")")
			loComandoSeleccionar.AppendLine("")

			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("INSERT INTO @VariableTabla") 
            loComandoSeleccionar.AppendLine("(		Tabla, Cod_Pro,Nom_Pro,Status,Responsable,Resumen,Objetivo,Comentario,Por_Eje,Fec_Ini,Fec_Fin,")
			loComandoSeleccionar.AppendLine("		Etapa,Duracion,Tipo,Clase,Cod_Mon,Nom_Mon,Tasa,Prioridad,Nivel,Importancia,")
			loComandoSeleccionar.AppendLine("		Logico1,Logico2,Logico3,Logico4,Logico5,Fecha1,Fecha2,Fecha3,Fecha4,Fecha5,")
			loComandoSeleccionar.AppendLine("		Numerico1,Numerico2,Numerico3,Numerico4,Numerico5,")
			loComandoSeleccionar.AppendLine("		Caracter1,Caracter2,Caracter3,Caracter4,Caracter5,Memo1,Memo2,Memo3,Memo4,Memo5")
			loComandoSeleccionar.AppendLine(")")
			loComandoSeleccionar.AppendLine("")

			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT	'Base',						#tmpOriginal.Cod_Pro,	#tmpOriginal.Nom_Pro,		#tmpOriginal.Status,")	
            loComandoSeleccionar.AppendLine("		#tmpOriginal.Responsable,	#tmpOriginal.Resumen,  #tmpOriginal.Objetivo,	    #tmpOriginal.Comentario,	#tmpOriginal.Por_Eje,")
			loComandoSeleccionar.AppendLine("		#tmpOriginal.Fec_Ini,		#tmpOriginal.Fec_Fin,	#tmpOriginal.Etapa,			#tmpOriginal.Duracion,")		
			loComandoSeleccionar.AppendLine("		#tmpOriginal.Tipo,			#tmpOriginal.Clase,		#tmpOriginal.Cod_Mon,		#tmpOriginal.Nom_Mon,")		
			loComandoSeleccionar.AppendLine("		#tmpOriginal.Tasa,			#tmpOriginal.Prioridad,	#tmpOriginal.Nivel,			#tmpOriginal.Importancia,")	
			loComandoSeleccionar.AppendLine("		#tmpOriginal.Logico1,		#tmpOriginal.Logico2,	#tmpOriginal.Logico3,		#tmpOriginal.Logico4,")		
			loComandoSeleccionar.AppendLine("		#tmpOriginal.Logico5,		#tmpOriginal.Fecha1,	#tmpOriginal.Fecha2,		#tmpOriginal.Fecha3,")		
			loComandoSeleccionar.AppendLine("		#tmpOriginal.Fecha4,		#tmpOriginal.Fecha5,	#tmpOriginal.Numerico1,		#tmpOriginal.Numerico2,")		
			loComandoSeleccionar.AppendLine("		#tmpOriginal.Numerico3,		#tmpOriginal.Numerico4,	#tmpOriginal.Numerico5,		#tmpOriginal.Caracter1,	#tmpOriginal.Caracter2,")	
			loComandoSeleccionar.AppendLine("		#tmpOriginal.Caracter3,		#tmpOriginal.Caracter4,	#tmpOriginal.Caracter5,		#tmpOriginal.Memo1,")			
			loComandoSeleccionar.AppendLine("		#tmpOriginal.Memo2,			#tmpOriginal.Memo3,		#tmpOriginal.Memo4,			#tmpOriginal.Memo5   ")
			loComandoSeleccionar.AppendLine("FROM #tmpOriginal")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("INSERT INTO @VariableTabla ")
			loComandoSeleccionar.AppendLine("(Tabla,se_renglon,se_fecha,se_contacto,se_accion,se_medio,se_comentario,se_prioridad,se_etapa,se_usuario)")
			loComandoSeleccionar.AppendLine("SELECT	'Seguimiento',		#tmpSeguimiento.se_renglon,		#tmpSeguimiento.se_fecha,	#tmpSeguimiento.se_contacto,")		
			loComandoSeleccionar.AppendLine("		#tmpSeguimiento.se_accion,		#tmpSeguimiento.se_medio,	#tmpSeguimiento.se_comentario,")	
			loComandoSeleccionar.AppendLine("		#tmpSeguimiento.se_prioridad,	#tmpSeguimiento.se_etapa,	#tmpSeguimiento.se_usuario")
			loComandoSeleccionar.AppendLine("FROM #tmpSeguimiento")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("INSERT INTO @VariableTabla") 
			loComandoSeleccionar.AppendLine("(Tabla,fa_fase,fa_responsable,fa_actividad,fa_etapa,fa_porcentaje,fa_inicio,fa_fin,fa_costos,fa_comentario,fa_usuario)")
			loComandoSeleccionar.AppendLine("SELECT	'Fase',	#tmpFases.fa_fase,			#tmpFases.fa_responsable,	#tmpFases.fa_actividad,	#tmpFases.fa_etapa,")
			loComandoSeleccionar.AppendLine("		#tmpFases.fa_porcentaje,	#tmpFases.fa_inicio,		#tmpFases.fa_fin,		#tmpFases.fa_costos,")
			loComandoSeleccionar.AppendLine("		#tmpFases.fa_comentario,	#tmpFases.fa_usuario")
			loComandoSeleccionar.AppendLine("FROM #tmpFases	")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")

			loComandoSeleccionar.AppendLine("SELECT #tmpFinal.* FROM @VariableTabla AS #tmpFinal")
			loComandoSeleccionar.AppendLine("ORDER BY #tmpFinal.Tabla ASC")

			loComandoSeleccionar.AppendLine("DROP TABLE #tmpSeguimiento")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpFases")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpOriginal")
            
								  
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
					 
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fProyectos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfProyectos.ReportSource = loObjetoReporte

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