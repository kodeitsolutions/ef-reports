'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rGrupos_Accesos"
'-------------------------------------------------------------------------------------------'
Partial Class rGrupos_Accesos
    Inherits vis2formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


		Try	
		
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			Dim lcSistema As String = goServicios.mObtenerCampoFormatoSQL("Factory_" & goAplicacion.pcNombre.Trim())
            Dim loComandoSeleccionar As New StringBuilder()
			
    		
            loComandoSeleccionar.AppendLine("DECLARE @lcCliente AS CHAR(10) ")
            loComandoSeleccionar.AppendLine("SET @lcCliente = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @llFalso AS BIT")
            loComandoSeleccionar.AppendLine("SET @llFalso = 0 ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Grupos.Cod_Gru					AS Cod_Gru,	")
            loComandoSeleccionar.AppendLine("			Grupos.Nom_Gru					AS Nom_Gru,	")
            loComandoSeleccionar.AppendLine("			(CASE	Grupos.Tipo					")
            loComandoSeleccionar.AppendLine("				WHEN 'A' THEN 'Administrador'	")
			loComandoSeleccionar.AppendLine("				WHEN 'I' THEN 'Invitado'		")
			loComandoSeleccionar.AppendLine("				WHEN 'U' Then 'Usuario'			")
			loComandoSeleccionar.AppendLine("				WHEN 'S' Then 'Supervisor'		")
			loComandoSeleccionar.AppendLine("				ELSE NULL						")
			loComandoSeleccionar.AppendLine("			END)							AS Tipo,	")
            loComandoSeleccionar.AppendLine("			Grupos.Nivel					AS Nivel,	")
            loComandoSeleccionar.AppendLine("			(CASE	Grupos.Status				")
            loComandoSeleccionar.AppendLine("				WHEN 'A' THEN 'Activo'			")
			loComandoSeleccionar.AppendLine("				WHEN 'I' THEN 'Inactivo'		")
			loComandoSeleccionar.AppendLine("				WHEN 'U' Then 'Suspendido'		")
			loComandoSeleccionar.AppendLine("				ELSE NULL						")
			loComandoSeleccionar.AppendLine("			END)							AS Status, ")
			loComandoSeleccionar.AppendLine("			CAST(								")
			loComandoSeleccionar.AppendLine("				REPLACE(CAST(Accesos AS NVARCHAR(MAX)), ")
			loComandoSeleccionar.AppendLine("				'<?xml version=""1.0"" encoding=""utf-8""?>', '')")
			loComandoSeleccionar.AppendLine("			AS XML)							AS Accesos")
            loComandoSeleccionar.AppendLine("INTO		#tmpGrupos							")
            loComandoSeleccionar.AppendLine("FROM		Grupos								")
            loComandoSeleccionar.AppendLine("WHERE		Grupos.Cod_Gru BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND Grupos.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("		AND Grupos.Cod_Cli  = @lcCliente		")
            loComandoSeleccionar.AppendLine("		ANd Grupos.Sistema = "  & lcSistema)
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Cod_Gru, Nom_Gru, Status, Tipo, Nivel,")
            loComandoSeleccionar.AppendLine("		T.C.value('(../../../@nombre)[1]', 'varchar(MAX)')  AS Sistema,")
            loComandoSeleccionar.AppendLine("		T.C.value('(../../@nombre)[1]', 'varchar(MAX)')		AS Modulo,")
            loComandoSeleccionar.AppendLine("		T.C.value('(../@nombre)[1]', 'varchar(MAX)')		AS Seccion,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@nombre)[1]', 'varchar(MAX)')			AS Formulario, ")
            loComandoSeleccionar.AppendLine("		T.C.value('(@acciones)[1]', 'varchar(MAX)')			AS Acciones,")
            loComandoSeleccionar.AppendLine("		@llFalso											AS Agregar,")
            loComandoSeleccionar.AppendLine("		@llFalso											AS Editar,")
            loComandoSeleccionar.AppendLine("		@llFalso											AS Buscar,")
            loComandoSeleccionar.AppendLine("		@llFalso											AS Eliminar,")
            loComandoSeleccionar.AppendLine("		@llFalso											AS Imprimir,")
            loComandoSeleccionar.AppendLine("		@llFalso											AS Resumen,")
            loComandoSeleccionar.AppendLine("		@llFalso											AS Avanzado")
            loComandoSeleccionar.AppendLine("INTO	#tmpAccesos")
            loComandoSeleccionar.AppendLine("FROM	#tmpGrupos")
            loComandoSeleccionar.AppendLine("	CROSS APPLY Accesos.nodes('//sistema/modulo/seccion/formulario') AS T(C)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpGrupos")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Marca las acciones disponibles")
            loComandoSeleccionar.AppendLine("UPDATE	#tmpAccesos")
            loComandoSeleccionar.AppendLine("SET	Agregar		= (CASE WHEN CHARINDEX('agregar', Acciones) 	<= 0 THEN 0 ELSE 1 END),")
            loComandoSeleccionar.AppendLine("		Editar		= (CASE WHEN CHARINDEX('editar', Acciones)		<= 0 THEN 0 ELSE 1 END),")
            loComandoSeleccionar.AppendLine("		Buscar		= (CASE WHEN CHARINDEX('buscar', Acciones)		<= 0 THEN 0 ELSE 1 END),")
            loComandoSeleccionar.AppendLine("		Eliminar	= (CASE WHEN CHARINDEX('eliminar', Acciones)	<= 0 THEN 0 ELSE 1 END),")
            loComandoSeleccionar.AppendLine("		Imprimir	= (CASE WHEN CHARINDEX('imprimir', Acciones)	<= 0 THEN 0 ELSE 1 END),")
            loComandoSeleccionar.AppendLine("		Resumen		= (CASE WHEN CHARINDEX('resumen', Acciones) 	<= 0 THEN 0 ELSE 1 END),")
            loComandoSeleccionar.AppendLine("		Avanzado	= (CASE WHEN CHARINDEX('avanzado', Acciones)	<= 0 THEN 0 ELSE 1 END)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- Selecciona solo las columnas necesarias")
            loComandoSeleccionar.AppendLine("SELECT	Cod_Gru, Nom_Gru, Status, Sistema, Modulo, Seccion, Formulario, ")
            loComandoSeleccionar.AppendLine("		Agregar, Editar, Buscar, Eliminar, Imprimir, Resumen, Avanzado")
            loComandoSeleccionar.AppendLine("FROM #tmpAccesos")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpAccesos")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
		
			Dim loServicios As New cusDatos.goDatos
	      
			goDatos.pcNombreAplicativoExterno = "Framework"

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
			
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rGrupos_Accesos", laDatosReporte)
		   
			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)
			
			Me.crvrGrupos_Accesos.ReportSource = loObjetoReporte
				
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
' RJG: 13/01/12: Codigo inicial
'-------------------------------------------------------------------------------------------'
