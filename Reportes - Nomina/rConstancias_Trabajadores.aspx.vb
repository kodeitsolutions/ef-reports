'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rConstancias_Trabajadores"
'-------------------------------------------------------------------------------------------'
Partial Class rConstancias_Trabajadores
    Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument
	
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))'Trabajador
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))'Estatus
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))'Contrato
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3)) 'Departamento
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4)) 'Cargo
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5)) 'Sucursal
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6)) 'Dirigido A
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaSinHoras) 'Emisión
        Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8)) 'Usuario
        
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  Trabajadores.Cod_Tra                            AS Cod_Tra,")
            loConsulta.AppendLine("        Trabajadores.Nom_Tra                            AS Nom_Tra,")
            loConsulta.AppendLine("        Trabajadores.Cedula                             AS Cedula , ")
            loConsulta.AppendLine("        Departamentos_Nomina.Nom_Dep                    AS Nom_Dep,")
            loConsulta.AppendLine("        Cargos.Nom_Car                                  AS Nom_Car,")
            loConsulta.AppendLine("        Trabajadores.Fec_Ini                            AS Fec_Ini,")

            loConsulta.AppendLine("        COALESCE(Renglones_Campos_Nomina.Val_Num, 0)    AS Sueldo_Mensual,")
            loConsulta.AppendLine("        COALESCE(Usuarios.Nom_Usu, '')                  AS Usuario_Firma, ")
            loConsulta.AppendLine("        COALESCE(Usuarios.Cargo, '')                    AS Cargo_Firma,")
            loConsulta.AppendLine("        COALESCE(Usuarios.Correo, '')                   AS Correo_Firma,")
            loConsulta.AppendLine("        COALESCE((CASE WHEN Usuarios.Telefonos <> '' ")
            loConsulta.AppendLine("                     THEN Usuarios.Telefonos ")
            loConsulta.AppendLine("                     ELSE Usuarios.Movil END), '')      AS Telefonos_Firma,")
            loConsulta.AppendLine("        CAST(" & lcParametro7Desde & " AS DATE)         AS Fecha_Emision,")
            loConsulta.AppendLine("        " & lcParametro6Desde & "                       AS Dirigido_A")

            loConsulta.AppendLine("FROM    Trabajadores")
            loConsulta.AppendLine("    JOIN Departamentos_Nomina")
            loConsulta.AppendLine("        ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep ")
            loConsulta.AppendLine("    JOIN Cargos")
            loConsulta.AppendLine("        ON Cargos.Cod_Car = Trabajadores.Cod_Car")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina")
            loConsulta.AppendLine("        ON Renglones_Campos_Nomina.Cod_TRa = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Renglones_Campos_Nomina.Cod_Cam = 'A001'")
            loConsulta.AppendLine("    LEFT JOIN Factory_Global.dbo.Usuarios Usuarios ")
            loConsulta.AppendLine("        ON Usuarios.Cod_Usu = " & lcParametro8Desde)
            loConsulta.AppendLine("WHERE   Trabajadores.Cod_Tra BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("         AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Status IN ( " & lcParametro1Desde & " )")
            loConsulta.AppendLine("        AND Trabajadores.Cod_Con BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("         AND " & lcParametro2Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Dep BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("         AND " & lcParametro3Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Car BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("         AND " & lcParametro4Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Suc BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("         AND " & lcParametro5Hasta)
            loConsulta.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rConstancias_Trabajadores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrConstancias_Trabajadores.ReportSource = loObjetoReporte

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
' RJG: 06/06/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
