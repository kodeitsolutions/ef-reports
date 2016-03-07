'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCampanasConParticipantes"
'-------------------------------------------------------------------------------------------'
Partial Class rCampanasConParticipantes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))

        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try


            loConsulta.AppendLine(" SELECT  ")
            loConsulta.AppendLine("     Campanas.Documento AS Documento,")
            loConsulta.AppendLine("     Campanas.Nombre AS Nombre,")
            loConsulta.AppendLine("     Campanas.Responsable AS Responsable,")
            loConsulta.AppendLine("     Campanas.Fec_Ini,")
            loConsulta.AppendLine("     Campanas.Fec_Fin,")
            loConsulta.AppendLine("     Campanas.Por_Eje,")
            loConsulta.AppendLine("     Campanas.Status,")
            loConsulta.AppendLine("     Tipo_Renglon,")
            loConsulta.AppendLine("     COALESCE(Renglones.Renglon,0) AS Renglon,")
            loConsulta.AppendLine("     Renglones.Nombre AS Nombre_Reng,")
            loConsulta.AppendLine("     Renglones.Fec_Ini AS Fec_Ini_Reng,")
            loConsulta.AppendLine("     Renglones.Fec_Fin AS Fec_Fin_Reng,")
            loConsulta.AppendLine("     Renglones.Telefonos,")
            loConsulta.AppendLine("     Renglones.Correo,")
            loConsulta.AppendLine("     Campanas.Etapa,")
            loConsulta.AppendLine("     Campanas.Prioridad")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("   FROM Campanas")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("   LEFT JOIN (")
            loConsulta.AppendLine("     SELECT ")
            loConsulta.AppendLine("       ClientesCamp.Documento,")
            loConsulta.AppendLine("        'Clientes'  AS Tipo_Renglon,")
            loConsulta.AppendLine("        COALESCE(T.C.value('(@renglon)[1]', 'Integer'), 0)  AS Renglon,")
            loConsulta.AppendLine("            COALESCE(T.C.value('(@cliente)[1]', 'varchar(MAX)'), '')  AS Nombre,")
            loConsulta.AppendLine("            COALESCE(T.C.value('(@fec_ini)[1]', 'dateTime'), '')  AS Fec_Ini,")
            loConsulta.AppendLine("            COALESCE(T.C.value('(@fec_fin)[1]', 'dateTime'), '')  AS Fec_Fin,")
            loConsulta.AppendLine("        COALESCE(T.C.value('(@telefonos)[1]', 'varchar(MAX)'), '')  AS Telefonos,")
            loConsulta.AppendLine("            COALESCE(T.C.value('(@correo)[1]', 'varchar(MAX)'), '')  AS Correo")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("     FROM (")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("       SELECT Documento, CAST(Clientes AS XML) as Clientes FROM Campanas ")
            loConsulta.AppendLine("     )")
            loConsulta.AppendLine("     AS ClientesCamp")
            loConsulta.AppendLine("     OUTER APPLY Clientes.nodes('//elementos/elemento') AS T(C)")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("     UNION ALL ")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("     SELECT ")
            loConsulta.AppendLine("       SociosCamp.Documento,")
            loConsulta.AppendLine("        'Socios'  AS Tipo_Renglon,")
            loConsulta.AppendLine("        COALESCE(W.C.value('(@renglon)[1]', 'Integer'), 0)  AS Renglon,")
            loConsulta.AppendLine("            COALESCE(W.C.value('(@socio)[1]', 'varchar(MAX)'), '')    AS Nombre,")
            loConsulta.AppendLine("            COALESCE(W.C.value('(@fec_ini)[1]', 'dateTime'), '')  AS Fec_Ini,")
            loConsulta.AppendLine("            COALESCE(W.C.value('(@fec_fin)[1]', 'dateTime'), '')  AS Fec_Fin,")
            loConsulta.AppendLine("        COALESCE(W.C.value('(@telefonos)[1]', 'varchar(MAX)'), '')  AS Telefonos,")
            loConsulta.AppendLine("            COALESCE(W.C.value('(@correo)[1]', 'varchar(MAX)'), '')  AS Correo")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("     FROM (")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("       SELECT Documento, CAST(Socios AS XML) as Socios FROM Campanas ")
            loConsulta.AppendLine("     )")
            loConsulta.AppendLine("     AS SociosCamp")
            loConsulta.AppendLine("     OUTER APPLY Socios.nodes('//elementos/elemento') AS W(C)")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("     UNION ALL")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("     SELECT ")
            loConsulta.AppendLine("        ContactosCamp.Documento,")
            loConsulta.AppendLine("        'Contactos'   AS Tipo_Renglon,")
            loConsulta.AppendLine("        COALESCE(K.C.value('(@renglon)[1]', 'Integer'), 0)  AS Renglon,")
            loConsulta.AppendLine("            COALESCE(K.C.value('(@contacto)[1]', 'varchar(MAX)'), '') AS Nombre,")
            loConsulta.AppendLine("        COALESCE(K.C.value('(@fec_ini)[1]', 'varchar(MAX)'), '')  AS Fec_Ini,")
            loConsulta.AppendLine("            COALESCE(K.C.value('(@fec_fin)[1]', 'varchar(MAX)'), '')  AS Fec_Fin,")
            loConsulta.AppendLine("        COALESCE(K.C.value('(@telefonos)[1]', 'varchar(MAX)'), '')  AS Telefonos,")
            loConsulta.AppendLine("            COALESCE(K.C.value('(@correo)[1]', 'varchar(MAX)'), '')  AS Correo")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("     FROM (")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("       SELECT Documento, CAST(Contactos AS XML) as Contactos FROM Campanas ")
            loConsulta.AppendLine("     )")
            loConsulta.AppendLine("     AS ContactosCamp")
            loConsulta.AppendLine("     OUTER APPLY Contactos.nodes('//elementos/elemento') AS K(C)")
            loConsulta.AppendLine("   )")
            loConsulta.AppendLine("   AS Renglones")
            loConsulta.AppendLine("   ON Campanas.Documento = Renglones.Documento")

            loConsulta.AppendLine("WHERE         Campanas.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("     AND      Campanas.Etapa IN (" & lcParametro1Desde & " )")
            loConsulta.AppendLine("     AND      Campanas.Prioridad IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine("     AND      Campanas.Fec_Ini BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            'loConsulta.AppendLine(" Campanas.Adicional = 'Marketing'")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCampanasConParticipantes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCampanasConParticipantes.ReportSource = loObjetoReporte

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
' JAC : 28/07/15 : Codigo inicial
'-------------------------------------------------------------------------------------------'

