'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fhCambios_cNomina"
'-------------------------------------------------------------------------------------------'
Partial Class fhCambios_cNomina
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("DECLARE @lcCodigoTrabajador CHAR(10);")
            loConsulta.AppendLine("DECLARE @lcNombreTrabajador CHAR(100);")
            loConsulta.AppendLine("DECLARE @lcEstatusTrabajador CHAR(10);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  @lcCodigoTrabajador = Trabajadores.cod_tra,")
            loConsulta.AppendLine("        @lcNombreTrabajador = Trabajadores.nom_tra,")
            loConsulta.AppendLine("    	   @lcEstatusTrabajador = (CASE Trabajadores.Status ")
            loConsulta.AppendLine("                                  WHEN 'A' THEN 'Activo' ")
            loConsulta.AppendLine("                                  WHEN 'I' THEN 'Inactivo' ")
            loConsulta.AppendLine("                                  WHEN 'S' THEN 'Suspendido' ")
            loConsulta.AppendLine("    	                          END) ")
            loConsulta.AppendLine("FROM    Trabajadores")
            loConsulta.AppendLine("WHERE   " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- **************************************************************************** ")
            loConsulta.AppendLine("-- Busca las auditorias de cambios en Campos de Nómina (Renglones_Campos_Nomina)")
            loConsulta.AppendLine("-- **************************************************************************** ")
            loConsulta.AppendLine("CREATE TABLE #tmpAuditorias(    Cod_Usu CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Registro DATETIME,")
            loConsulta.AppendLine("                                Accion CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Codigo_Campo CHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Notas VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Detalle XML,")
            loConsulta.AppendLine("                                Campo VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Antes VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Despues VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Tipo_Auditoria VARCHAR(20) COLLATE DATABASE_DEFAULT                 ")
            loConsulta.AppendLine("                            );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpAuditorias(Cod_Usu, Registro, Accion, Codigo_Campo, Notas, Detalle, Tipo_Auditoria)")
            loConsulta.AppendLine("SELECT  cod_usu, registro, accion, codigo, notas, detalle, 'Campos_Nominas'")
            loConsulta.AppendLine("FROM    auditorias ")
            loConsulta.AppendLine("WHERE   tabla = 'Renglones_Campos_Nomina'")
            loConsulta.AppendLine("    AND clave2 = @lcCodigoTrabajador")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- En las auditorias de Campos de Nómina el XML de detalle solo almacena un campo a la vez, ")
            loConsulta.AppendLine("-- por lo que se usa la función XML.value()")
            loConsulta.AppendLine("UPDATE #tmpAuditorias")
            loConsulta.AppendLine("SET     Campo = Detalle.value('(detalle/campos/campo/@nombre)[1]', 'varchar(MAX)'),")
            loConsulta.AppendLine("        Antes = Detalle.value('(detalle/campos/campo/antes)[1]', 'varchar(MAX)'),")
            loConsulta.AppendLine("        Despues = Detalle.value('(detalle/campos/campo/despues)[1]', 'varchar(MAX)');")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- **************************************************************************** ")
            loConsulta.AppendLine("-- Busca las auditorias de cambios en la tabla Trabajadores (Trabajadores)")
            loConsulta.AppendLine("-- **************************************************************************** ")
            loConsulta.AppendLine("CREATE TABLE #tmpAuditoriasTrabajadores(   ")
            loConsulta.AppendLine("                                Cod_Usu CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Registro DATETIME,")
            loConsulta.AppendLine("                                Accion CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Notas VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Detalle XML")
            loConsulta.AppendLine("                            );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpAuditoriasTrabajadores(Cod_Usu, Registro, Accion, Notas, Detalle)")
            loConsulta.AppendLine("SELECT  cod_usu, registro, accion, notas, detalle")
			loConsulta.AppendLine("FROM    auditorias ")
            loConsulta.AppendLine("WHERE   tabla = 'Trabajadores'")
            loConsulta.AppendLine("    AND codigo = @lcCodigoTrabajador")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- En las auditorias de Trabajadores el XML de detalle puede almacenar más de un campo a la vez, ")
            loConsulta.AppendLine("-- por lo que se usa CROSS APPLY junto a la función XML.Nodes() antes de usar XML.value()")
            loConsulta.AppendLine("INSERT INTO #tmpAuditorias(Cod_Usu, Registro, Accion, Codigo_Campo, Notas, Detalle, Tipo_Auditoria, Campo, Antes, Despues)")
            loConsulta.AppendLine("SELECT  #tmpAuditoriasTrabajadores.Cod_Usu, ")
            loConsulta.AppendLine("        #tmpAuditoriasTrabajadores.Registro, ")
            loConsulta.AppendLine("        #tmpAuditoriasTrabajadores.Accion, ")
            loConsulta.AppendLine("        T.C.value('(@nombre)[1]', 'VARCHAR(MAX)') AS Codigo_Campo,")
            loConsulta.AppendLine("        #tmpAuditoriasTrabajadores.Notas, ")
            loConsulta.AppendLine("        #tmpAuditoriasTrabajadores.Detalle, ")
            loConsulta.AppendLine("        'Trabajadores',")
            loConsulta.AppendLine("        T.C.value('(@nombre)[1]', 'VARCHAR(MAX)') AS Campo,")
            loConsulta.AppendLine("        T.C.value('(antes)[1]', 'VARCHAR(MAX)') AS Antes,")
            loConsulta.AppendLine("        T.C.value('(despues)[1]', 'VARCHAR(MAX)') AS Despues")
			loConsulta.AppendLine("FROM    #tmpAuditoriasTrabajadores")
            loConsulta.AppendLine("	OUTER APPLY Detalle.nodes('//detalle/campos/campo') AS T(C)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- SELECT Final con el resultado")
            loConsulta.AppendLine("SELECT  @lcCodigoTrabajador                 AS Cod_Tra, ")
            loConsulta.AppendLine("        @lcNombreTrabajador                 AS Nom_Tra, ")
            loConsulta.AppendLine("        @lcEstatusTrabajador                AS Estatus, ")
            loConsulta.AppendLine("        #tmpAuditorias.Cod_Usu              AS Cod_Usu, ")
            loConsulta.AppendLine("        #tmpAuditorias.Registro             AS Registro, ")
            loConsulta.AppendLine("        #tmpAuditorias.Accion               AS Accion, ")
            loConsulta.AppendLine("        #tmpAuditorias.Codigo_Campo         AS Codigo_Campo, ")
            loConsulta.AppendLine("        COALESCE(Campos_Nomina.nom_cam, '') AS Nombre_Campo,")
            loConsulta.AppendLine("        #tmpAuditorias.Notas                AS Notas, ")
            loConsulta.AppendLine("        #tmpAuditorias.Campo                AS Campo, ")
            loConsulta.AppendLine("        #tmpAuditorias.Antes                AS Antes, ")
            loConsulta.AppendLine("        #tmpAuditorias.Despues              AS Despues,")
            loConsulta.AppendLine("        #tmpAuditorias.Tipo_Auditoria       AS Tipo_Auditoria,")
            loConsulta.AppendLine("        #tmpAuditorias.detalle")
            loConsulta.AppendLine("FROM    #tmpAuditorias")
            loConsulta.AppendLine("LEFT JOIN Campos_Nomina ")
            loConsulta.AppendLine("    ON  Campos_Nomina.cod_cam = #tmpAuditorias.Codigo_Campo")
            loConsulta.AppendLine("    AND #tmpAuditorias.Tipo_Auditoria = 'Campos_Nominas'")
            loConsulta.AppendLine("ORDER BY Registro DESC;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")																	 

            Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")
            
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fhCambios_cNomina", laDatosReporte)

			
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfhCambios_cNomina.ReportSource = loObjetoReporte

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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 28/06/13: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
