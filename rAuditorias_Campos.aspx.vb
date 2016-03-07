'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAuditorias_Campos"
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

Partial Class rAuditorias_Campos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))

            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpAuditorias(Cod_Usu VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Tipo_Documento VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Documento VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Codigo VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Registro DATETIME,")
            loConsulta.AppendLine("                            Accion VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Notas VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Detalle XML)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpAuditorias(Cod_Usu, Tipo_Documento, Documento, Codigo,")
            loConsulta.AppendLine("        Registro, Accion, Notas, Detalle)")
            loConsulta.AppendLine("SELECT  cod_usu, tabla, documento, codigo,")
            loConsulta.AppendLine("        registro, accion, notas, CAST(detalle AS XML)")
            loConsulta.AppendLine("FROM    auditorias")
            loConsulta.AppendLine(" WHERE  Registro BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("            AND " & lcParametro0Hasta)
            loConsulta.AppendLine("    AND Cod_Usu       BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("            AND " & lcParametro1Hasta)
            loConsulta.AppendLine("    AND Tabla         BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("            AND " & lcParametro2Hasta)
            loConsulta.AppendLine("    AND Opcion        BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("            AND " & lcParametro3Hasta)
            loConsulta.AppendLine("    AND Documento     BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("            AND " & lcParametro4Hasta)
            loConsulta.AppendLine("    AND Codigo        BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("            AND " & lcParametro5Hasta)
            loConsulta.AppendLine("    AND Cod_Emp       BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("            AND " & lcParametro6Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpDetalles(  Cod_Usu VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Tipo_Documento VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Documento VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Codigo VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Registro DATETIME,")
            loConsulta.AppendLine("                            Accion VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            E_Campo VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            E_Antes VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            E_Despues VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            R_Tabla VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            R_Accion VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            R_Campo VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            R_Antes VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            R_Despues VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Notas VARCHAR(MAX) COLLATE DATABASE_DEFAULT);")
            loConsulta.AppendLine("                            ")
            loConsulta.AppendLine("CREATE CLUSTERED INDEX PK_Detalles ON #tmpDetalles(Cod_Usu, Tipo_Documento, Documento, Registro);")
            loConsulta.AppendLine("                            ")
            loConsulta.AppendLine("-- Auditorias de Encabezados:")
            loConsulta.AppendLine("INSERT INTO #tmpDetalles(Cod_Usu, Tipo_Documento, Documento, ")
            loConsulta.AppendLine("                        Codigo, Registro, Accion, ")
            loConsulta.AppendLine("                        E_Campo, E_Antes, E_Despues, Notas) ")
            loConsulta.AppendLine("SELECT  #tmpAuditorias.Cod_Usu, ")
            loConsulta.AppendLine("        #tmpAuditorias.Tipo_Documento, ")
            loConsulta.AppendLine("        #tmpAuditorias.Documento, ")
            loConsulta.AppendLine("        #tmpAuditorias.Codigo,")
            loConsulta.AppendLine("        #tmpAuditorias.Registro, ")
            loConsulta.AppendLine("        #tmpAuditorias.Accion,")
            loConsulta.AppendLine("		E.C.value('(@nombre)[1]', 'varchar(MAX)')	AS E_Campo, ")
            loConsulta.AppendLine("		(CASE WHEN SUBSTRING(CAST(E.C.query('(antes)[1]/*') AS VARCHAR(MAX)), 1, 1) = '<' ")
            loConsulta.AppendLine("		    THEN CAST(E.C.query('(antes)[1]/*') AS VARCHAR(MAX))")
            loConsulta.AppendLine("		    ELSE E.C.value('(antes)[1]', 'varchar(MAX)')")
            loConsulta.AppendLine("		END)                                            AS E_Antes,")
            loConsulta.AppendLine("		(CASE WHEN SUBSTRING(CAST(E.C.query('(despues)[1]/*') AS VARCHAR(MAX)), 1, 1) = '<'")
            loConsulta.AppendLine("		    THEN CAST(E.C.query('(despues)[1]/*') AS VARCHAR(MAX))")
            loConsulta.AppendLine("		    ELSE E.C.value('(despues)[1]', 'varchar(MAX)')")
            loConsulta.AppendLine("		END)                                            AS E_Despues,")
            loConsulta.AppendLine("        #tmpAuditorias.Notas")
            loConsulta.AppendLine("FROM    #tmpAuditorias")
            loConsulta.AppendLine("	CROSS APPLY Detalle.nodes('//detalle/campos/campo') AS E(C)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Auditorias de Renglones:")
            loConsulta.AppendLine("INSERT INTO #tmpDetalles(Cod_Usu, Tipo_Documento, Documento, ")
            loConsulta.AppendLine("                        Codigo, Registro, Accion, R_Tabla, R_Accion,")
            loConsulta.AppendLine("                        R_Campo, R_Antes, R_Despues, Notas) ")
            loConsulta.AppendLine("SELECT  #tmpAuditorias.Cod_Usu, ")
            loConsulta.AppendLine("        #tmpAuditorias.Tipo_Documento, ")
            loConsulta.AppendLine("        #tmpAuditorias.Documento, ")
            loConsulta.AppendLine("        #tmpAuditorias.Codigo,")
            loConsulta.AppendLine("        #tmpAuditorias.Registro, ")
            loConsulta.AppendLine("        #tmpAuditorias.Accion,")
            loConsulta.AppendLine("        R.C.value('(../@tabla)[1]', 'varchar(MAX)')	    AS R_Tabla, ")
            loConsulta.AppendLine("        R.C.value('(../@accion)[1]', 'varchar(MAX)')	AS R_Accion, ")
            loConsulta.AppendLine("        R.C.value('(@nombre)[1]', 'varchar(MAX)')	    AS R_Campo, ")
            loConsulta.AppendLine("		(CASE WHEN SUBSTRING(CAST(R.C.query('(antes)[1]/*') AS VARCHAR(MAX)), 1, 1) = '<' ")
            loConsulta.AppendLine("		    THEN CAST(R.C.query('(antes)[1]/*') AS VARCHAR(MAX))")
            loConsulta.AppendLine("		    ELSE R.C.value('(antes)[1]', 'varchar(MAX)')")
            loConsulta.AppendLine("		END)                                            AS R_Antes,")
            loConsulta.AppendLine("		(CASE WHEN SUBSTRING(CAST(R.C.query('(despues)[1]/*') AS VARCHAR(MAX)), 1, 1) = '<'")
            loConsulta.AppendLine("		    THEN CAST(R.C.query('(despues)[1]/*') AS VARCHAR(MAX))")
            loConsulta.AppendLine("		    ELSE R.C.value('(despues)[1]', 'varchar(MAX)')")
            loConsulta.AppendLine("		END)                                            AS R_Despues,")
            loConsulta.AppendLine("        #tmpAuditorias.Notas")
            loConsulta.AppendLine("FROM    #tmpAuditorias")
            loConsulta.AppendLine("	CROSS APPLY Detalle.nodes('//detalle/renglon/campo') AS R(C)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT Cod_Usu, Tipo_Documento, Documento, Codigo,")
            loConsulta.AppendLine("        Registro, Accion, E_Campo, E_Antes, E_Despues,")
            loConsulta.AppendLine("        R_Tabla, R_Accion, R_Campo, R_Antes, R_Despues, ")
            loConsulta.AppendLine("        Notas")
            loConsulta.AppendLine("FROM #tmpDetalles")
            loConsulta.AppendLine("ORDER BY  " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAuditorias_Campos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrAuditorias_Campos.ReportSource = loObjetoReporte

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
' RJG: 18/06/13: Codigo inicial
'-------------------------------------------------------------------------------------------'
