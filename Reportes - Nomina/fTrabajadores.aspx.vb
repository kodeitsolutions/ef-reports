'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fTrabajadores"
'-------------------------------------------------------------------------------------------'
Partial Class fTrabajadores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT	Trabajadores.Cod_Tra			AS Cod_Tra,")
            loConsulta.AppendLine("			Trabajadores.Nom_Tra			AS Nom_Tra, ")
            loConsulta.AppendLine("			Trabajadores.Cedula				AS Cedula, ")
            loConsulta.AppendLine("			CASE Trabajadores.Sexo")
            loConsulta.AppendLine("				WHEN 'F' THEN 'Femenino' ")
            loConsulta.AppendLine("				WHEN 'M' THEN 'Masculino' ")
            loConsulta.AppendLine("				ELSE 'No Aplica' ")
            loConsulta.AppendLine("			END								AS Sexo, ")
            loConsulta.AppendLine("		    Trabajadores.Fec_Nac			AS Fec_Nac, ")
            loConsulta.AppendLine("		    Trabajadores.Telefonos			AS Telefonos, ")
            loConsulta.AppendLine("		    Trabajadores.Movil				AS Movil, ")
            loConsulta.AppendLine("		    Trabajadores.Correo				AS Correo, ")
            loConsulta.AppendLine("		    Trabajadores.Cod_Nac			AS Cod_Nac, ")
            loConsulta.AppendLine("		    Nacionalidades.Nom_Nac			AS Nom_Nac, ")
            loConsulta.AppendLine("		    Trabajadores.Cod_Est			AS Cod_Est, ")
            loConsulta.AppendLine("		    Estados_Civiles.Nom_Est			AS Nom_Est, ")
            loConsulta.AppendLine("		    Trabajadores.Cod_Gra			AS Cod_Gra, ")
            loConsulta.AppendLine("		    Grados_Instruccion.Nom_Gra		AS Nom_Gra, ")
            loConsulta.AppendLine("		    Trabajadores.Direccion			AS Direccion, ")
            loConsulta.AppendLine("    		CASE Trabajadores.Status")
            loConsulta.AppendLine("    			WHEN 'A' THEN 'Activo' ")
            loConsulta.AppendLine("    			WHEN 'I' THEN 'Inactivo' ")
            loConsulta.AppendLine("    			WHEN 'S' THEN 'Suspendido' ")
            loConsulta.AppendLine("    		END								AS Status, ")
            loConsulta.AppendLine("    		Trabajadores.Cod_Dep			AS Cod_Dep, ")
            loConsulta.AppendLine("    		Departamentos_Nomina.Nom_Dep	AS Nom_Dep, ")
            loConsulta.AppendLine("    		Trabajadores.Cod_Car			AS Cod_Car, ")
            loConsulta.AppendLine("    		Cargos.Nom_Car					AS Nom_Car, ")
            loConsulta.AppendLine("    		Trabajadores.Cod_Tur			AS Cod_Tur, ")
            loConsulta.AppendLine("    		Turnos_Nomina.Nom_Tur			AS Nom_Tur, ")
            loConsulta.AppendLine("    		Trabajadores.Cod_Con			AS Cod_Con, ")
            loConsulta.AppendLine("    		Contratos.Nom_Con				AS Nom_Con, ")
            loConsulta.AppendLine("    		Trabajadores.Fec_Ini			AS Fec_Ini, ")
            loConsulta.AppendLine("    		Trabajadores.Fec_Fin			AS Fec_Fin, ")
            loConsulta.AppendLine("    		Trabajadores.Lug_Nac			AS Lug_Nac, ")
            loConsulta.AppendLine("    		Trabajadores.Fax				AS Fax, ")
            loConsulta.AppendLine("    		Trabajadores.Web				AS Web, ")
            loConsulta.AppendLine("    		Trabajadores.Cod_Ban			AS Cod_Ban, ")
            loConsulta.AppendLine("    		Bancos.Nom_Ban					AS Nom_Ban, ")
            loConsulta.AppendLine("    		Trabajadores.Num_Cue			AS Num_Cue, ")
            loConsulta.AppendLine("    		Trabajadores.Comentario			AS Comentario")
            loConsulta.AppendLine("FROM		Trabajadores")
            loConsulta.AppendLine("	JOIN	Estados_Civiles ")
            loConsulta.AppendLine("		ON Estados_Civiles.Cod_Est = Trabajadores.cod_est")
            loConsulta.AppendLine("	JOIN	Nacionalidades")
            loConsulta.AppendLine("		ON Nacionalidades.cod_nac = Trabajadores.cod_nac")
            loConsulta.AppendLine("	JOIN	Departamentos_Nomina")
            loConsulta.AppendLine("		ON Departamentos_Nomina.cod_dep = Trabajadores.cod_dep")
            loConsulta.AppendLine("	JOIN	Cargos")
            loConsulta.AppendLine("		ON Cargos.cod_car = Trabajadores.cod_car")
            loConsulta.AppendLine("	JOIN	Turnos_Nomina")
            loConsulta.AppendLine("		ON Turnos_Nomina.cod_tur = Trabajadores.cod_tur")
            loConsulta.AppendLine("	JOIN	Contratos")
            loConsulta.AppendLine("		ON Contratos.cod_con = Trabajadores.cod_con")
            loConsulta.AppendLine("	JOIN	Grados_Instruccion")
            loConsulta.AppendLine("		ON Grados_Instruccion.Cod_Gra = Trabajadores.Cod_Gra")
            loConsulta.AppendLine("	LEFT JOIN Bancos")
            loConsulta.AppendLine("		ON Bancos.cod_ban = Trabajadores.cod_ban")
            loConsulta.AppendLine("WHERE    " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
			loConsulta.AppendLine("")																	 

            Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")
            
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fTrabajadores", laDatosReporte)

			
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfTrabajadores.ReportSource = loObjetoReporte

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
' CMS: 21/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 06/02/10: Ajustes a los datos.
'-------------------------------------------------------------------------------------------'
' RJG: 27/02/13: Se actualizó el SELECT y los JOINS de a cuerdo a los campos actuales de las'
'				 tablas.																	'
'-------------------------------------------------------------------------------------------'
' RJG: 28/06/13: Se ajustó la unión con la tabla bancos: se cambió por LEFT JOIN porque el  '
'                código de banco es opcional para el trabajador.					        '
'-------------------------------------------------------------------------------------------'
