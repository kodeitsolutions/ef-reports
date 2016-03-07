'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCobros_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rCobros_Renglones

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cobros.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Cobros.Documento, ")
            loComandoSeleccionar.AppendLine("           Cobros.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cobros.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cobros.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cobros.Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           Cobros.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Cobros.Comentario, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cobros.Renglon    AS  Ren_Doc, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cobros.Tip_Doc    AS  Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cobros.Cod_Tip    AS  Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cobros.Doc_Ori    AS  Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' THEN Renglones_Cobros.Mon_Net ELSE (Renglones_Cobros.Mon_Net * -1) END)  AS  Mon_NetD, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' THEN Renglones_Cobros.Mon_Abo ELSE (Renglones_Cobros.Mon_Abo * -1) END)  AS  Mon_Abo, ")
            loComandoSeleccionar.AppendLine("           0.00                        AS  Ren_Tip, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Tip_Ope, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Doc_Des, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Num_Doc, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Cod_Tar, ")
            loComandoSeleccionar.AppendLine("           0.00                        AS  Mon_NetTP, ")
            loComandoSeleccionar.AppendLine("           'Documentos'                AS  Tipo, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Nom_Caj, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Nom_Ban, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Nom_Cue, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Nom_Tar ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpDocumentos ")
            loComandoSeleccionar.AppendLine(" FROM      Cobros, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cobros, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar, ")
            loComandoSeleccionar.AppendLine("           Clientes ")
            loComandoSeleccionar.AppendLine(" WHERE     Cobros.Documento                =   Renglones_Cobros.Documento ")
            loComandoSeleccionar.AppendLine("           And Cobros.Cod_Cli              =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tip      =   Renglones_Cobros.Cod_Tip ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    =   Renglones_Cobros.Doc_Ori ")
            loComandoSeleccionar.AppendLine("           And Cobros.Cod_Cli              =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           And Cobros.Documento            Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Cobros.Fec_Ini              Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Cobros.Cod_Cli              Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Cobros.Status               IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Rev between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Suc between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro5Hasta)

            loComandoSeleccionar.AppendLine(" SELECT	Cobros.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Cobros.Documento, ")
            loComandoSeleccionar.AppendLine("           Cobros.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cobros.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cobros.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cobros.Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           Cobros.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Cobros.Comentario, ")
            loComandoSeleccionar.AppendLine("           0                          AS  Ren_Doc, ")
            loComandoSeleccionar.AppendLine("           ''                         AS  Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           ''                         AS  Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           ''                         AS  Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           0.00                       AS  Mon_NetD, ")
            loComandoSeleccionar.AppendLine("           0.00                       AS  Mon_Abo, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Renglon    AS  Ren_Tip, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Tip_Ope    AS  Tip_Ope, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Doc_Des    AS  Doc_Des, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Num_Doc    AS  Num_Doc, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Cod_Caj    AS  Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Cod_Ban    AS  Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Cod_Cue    AS  Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Cod_Tar    AS  Cod_Tar, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Mon_Net    AS  Mon_NetTP, ")
            loComandoSeleccionar.AppendLine("           'TiposPagos'               AS  Tipo ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos1 ")
            loComandoSeleccionar.AppendLine(" FROM      Cobros, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros, ")
            loComandoSeleccionar.AppendLine("           Clientes ")
            loComandoSeleccionar.AppendLine(" WHERE     Cobros.Documento        =   Detalles_Cobros.Documento ")
            loComandoSeleccionar.AppendLine("           And Cobros.Cod_Cli      =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           And Cobros.Documento    Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Cobros.Fec_Ini      Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Cobros.Cod_Cli      Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Cobros.Status       IN (" & lcParametro3Desde &")")
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Suc between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro5Hasta)

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos1.*, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Cajas.Nom_Caj,1,25)   AS  Nom_Caj ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos2 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos1 LEFT JOIN Cajas ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos1.Cod_Caj =   Cajas.Cod_Caj ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos2.*, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Bancos.Nom_Ban,1,25)   AS  Nom_Ban ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos3 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos2 LEFT JOIN Bancos ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos2.Cod_Ban =   Bancos.Cod_Ban ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos3.*, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Cuentas_Bancarias.Nom_Cue,1,25)   AS  Nom_Cue ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos4 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos3 LEFT JOIN Cuentas_Bancarias ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos3.Cod_Cue =   Cuentas_Bancarias.Cod_Cue ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos4.*, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Tarjetas.Nom_Tar,1,25)   AS  Nom_Tar ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos5 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos4 LEFT JOIN Tarjetas ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos4.Cod_Tar =   Tarjetas.Cod_Tar ")

            loComandoSeleccionar.Append(" SELECT        * ")
            loComandoSeleccionar.Append(" FROM          #tmpDocumentos ")
            loComandoSeleccionar.Append(" UNION ALL ")
            loComandoSeleccionar.Append(" SELECT        * ")
            loComandoSeleccionar.Append(" FROM          #tmpTiposPagos5 ")
            'loComandoSeleccionar.Append(" ORDER BY      8, 30, 21")
            loComandoSeleccionar.AppendLine("ORDER BY    8, " & lcOrdenamiento)


            
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCobros_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCobros_Renglones.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' JJD: 07/03/09: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' JJD: 14/03/09: Ajustes a los calculos de la los renglones de documentos cobrados			'
'-------------------------------------------------------------------------------------------'
' GCR: 30/03/09: Ajustes al diseño															'
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”														'
'-------------------------------------------------------------------------------------------'
' AAP:  29/06/09: Filtro “Sucursal:”														'
'-------------------------------------------------------------------------------------------'
' CMS:  17/07/09: Metodo de Ordenamiento, Verificacion de registro							'
'-------------------------------------------------------------------------------------------'
' RJG:  10/04/12: Se agregó el total de documentos.											'
'-------------------------------------------------------------------------------------------'
