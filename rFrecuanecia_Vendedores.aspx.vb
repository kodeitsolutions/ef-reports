'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rFrecuanecia_Vendedores"
'-------------------------------------------------------------------------------------------'
Partial Class rFrecuanecia_Vendedores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

			
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 			Frecuencias.Origen,		")
            loComandoSeleccionar.AppendLine(" 			Frecuencias.Tip_fre,	")
            loComandoSeleccionar.AppendLine(" 			Frecuencias.Cod_reg,	")
            loComandoSeleccionar.AppendLine(" 			Frecuencias.Adicional,	")
            loComandoSeleccionar.AppendLine(" 			Frecuencias.Comentario,	")
            loComandoSeleccionar.AppendLine(" 			Frecuencias.Notas,		")
            loComandoSeleccionar.AppendLine(" 			Frecuencias.Foto,		")
            loComandoSeleccionar.AppendLine(" 			Frecuencias.Hora,		")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_1 = 1 THEN 'x' ELSE '' END AS   Hor_1 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_2 = 1 THEN 'x' ELSE '' END AS   Hor_2 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_3 = 1 THEN 'x' ELSE '' END AS   Hor_3   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_4 = 1 THEN 'x' ELSE '' END AS   Hor_4   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_5 = 1 THEN 'x' ELSE '' END AS   Hor_5   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_6 = 1 THEN 'x' ELSE '' END AS   Hor_6   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_7 = 1 THEN 'x' ELSE '' END AS   Hor_7   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_8 = 1 THEN 'x' ELSE '' END AS   Hor_8   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_9 = 1 THEN 'x' ELSE '' END AS   Hor_9   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_10 = 1 THEN 'x' ELSE '' END AS  Hor_10 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_11 = 1 THEN 'x' ELSE '' END AS  Hor_11 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_12 = 1 THEN 'x' ELSE '' END AS  Hor_12 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_13 = 1 THEN 'x' ELSE '' END AS  Hor_13 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_14 = 1 THEN 'x' ELSE '' END AS  Hor_14 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_15 = 1 THEN 'x' ELSE '' END AS  Hor_15 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_16 = 1 THEN 'x' ELSE '' END AS  Hor_16 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_17 = 1 THEN 'x' ELSE '' END AS  Hor_17 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_18 = 1 THEN 'x' ELSE '' END AS  Hor_18 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_19 = 1 THEN 'x' ELSE '' END AS  Hor_19 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_20 = 1 THEN 'x' ELSE '' END AS  Hor_20 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_21 = 1 THEN 'x' ELSE '' END AS  Hor_21 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_22 = 1 THEN 'x' ELSE '' END AS  Hor_22 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_23 = 1 THEN 'x' ELSE '' END AS  Hor_23 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Hor_24 = 1 THEN 'x' ELSE '' END AS  Hor_24 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Men_Ene = 1 THEN 'x' ELSE '' END AS Men_Ene 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Men_Feb = 1 THEN 'x' ELSE '' END AS Men_Feb 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Men_Mar = 1 THEN 'x' ELSE '' END AS Men_Mar 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Men_Abr = 1 THEN 'x' ELSE '' END AS Men_Abr 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Men_May = 1 THEN 'x' ELSE '' END AS Men_May 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Men_Jun = 1 THEN 'x' ELSE '' END AS Men_Jun 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Men_Jul = 1 THEN 'x' ELSE '' END AS Men_Jul 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Men_Ago = 1 THEN 'x' ELSE '' END AS Men_Ago 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Men_Sep = 1 THEN 'x' ELSE '' END AS Men_Sep 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Men_Oct = 1 THEN 'x' ELSE '' END AS Men_Oct 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Men_Nov = 1 THEN 'x' ELSE '' END AS Men_Nov 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Men_Dic = 1 THEN 'x' ELSE '' END AS Men_Dic 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Sem_Lun = 1 THEN 'x' ELSE '' END AS Sem_Lun 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Sem_Mar = 1 THEN 'x' ELSE '' END AS Sem_Mar 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Sem_Mie = 1 THEN 'x' ELSE '' END AS Sem_Mie 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Sem_Jue = 1 THEN 'x' ELSE '' END AS Sem_Jue 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Sem_Vie = 1 THEN 'x' ELSE '' END AS Sem_Vie 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Sem_Sab = 1 THEN 'x' ELSE '' END AS Sem_Sab 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Sem_Dom = 1 THEN 'x' ELSE '' END AS Sem_Dom 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia1 = 1 THEN 'x' ELSE '' END AS    Dia1   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia2 = 1 THEN 'x' ELSE '' END AS    Dia2   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia3 = 1 THEN 'x' ELSE '' END AS    Dia3   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia4 = 1 THEN 'x' ELSE '' END AS    Dia4   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia5 = 1 THEN 'x' ELSE '' END AS    Dia5   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia6 = 1 THEN 'x' ELSE '' END AS    Dia6   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia7 = 1 THEN 'x' ELSE '' END AS    Dia7   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia8 = 1 THEN 'x' ELSE '' END AS    Dia8   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia9 = 1 THEN 'x' ELSE '' END AS    Dia9   		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia10 = 1 THEN 'x' ELSE '' END AS   Dia10  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia11 = 1 THEN 'x' ELSE '' END AS   Dia11  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia12 = 1 THEN 'x' ELSE '' END AS   Dia12  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia13 = 1 THEN 'x' ELSE '' END AS   Dia13  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia14 = 1 THEN 'x' ELSE '' END AS   Dia14  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia15 = 1 THEN 'x' ELSE '' END AS   Dia15  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia16 = 1 THEN 'x' ELSE '' END AS   Dia16  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia17 = 1 THEN 'x' ELSE '' END AS   Dia17  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia18 = 1 THEN 'x' ELSE '' END AS   Dia18  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia19 = 1 THEN 'x' ELSE '' END AS   Dia19  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia20 = 1 THEN 'x' ELSE '' END AS   Dia20  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia21 = 1 THEN 'x' ELSE '' END AS   Dia21  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia22 = 1 THEN 'x' ELSE '' END AS   Dia22  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia23 = 1 THEN 'x' ELSE '' END AS   Dia23  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia24 = 1 THEN 'x' ELSE '' END AS   Dia24  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia25 = 1 THEN 'x' ELSE '' END AS   Dia25  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia26 = 1 THEN 'x' ELSE '' END AS   Dia26  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia27 = 1 THEN 'x' ELSE '' END AS   Dia27  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia28 = 1 THEN 'x' ELSE '' END AS   Dia28  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia29 = 1 THEN 'x' ELSE '' END AS   Dia29  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia30 = 1 THEN 'x' ELSE '' END AS   Dia30  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Dia31 = 1 THEN 'x' ELSE '' END AS   Dia31  		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Fec_ini = 1 THEN 'x' ELSE '' END AS Fec_ini 		, ")
            loComandoSeleccionar.AppendLine(" 			CASE WHEN Frecuencias.Fec_Fin = 1 THEN 'x' ELSE '' END AS Fec_Fin 		, ")
            loComandoSeleccionar.AppendLine(" 			Vendedores.Nom_Ven,		")
            loComandoSeleccionar.AppendLine(" 			Vendedores.Status			")
            loComandoSeleccionar.AppendLine(" FROM	Frecuencias ")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores ON (Vendedores.Cod_Ven = Frecuencias.Cod_Reg) AND ('Vendedores' = Frecuencias.Origen)")
            loComandoSeleccionar.AppendLine(" WHERE     Vendedores.Cod_Ven	Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Vendedores.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rFrecuanecia_Vendedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrFrecuanecia_Vendedores.ReportSource = loObjetoReporte

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
' CMS: 15/03/10: Codigo inicial.
'-------------------------------------------------------------------------------------------'