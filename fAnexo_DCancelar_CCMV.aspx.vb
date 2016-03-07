'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fAnexo_DCancelar_CCMV"
'-------------------------------------------------------------------------------------------'
Partial Class fAnexo_DCancelar_CCMV
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Cobros.Documento, ")
            loComandoSeleccionar.AppendLine(" 			Cobros.Status, ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Cuentas_Cobrar.Mon_Bru    * (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN -1 ELSE 1 END)))    AS  Mon_Bru, ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Cuentas_Cobrar.Mon_Rec    * (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN -1 ELSE 1 END)))    AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Cuentas_Cobrar.Mon_Des    * (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN -1 ELSE 1 END)))    AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Cuentas_Cobrar.Mon_Imp1   * (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN -1 ELSE 1 END)))    AS  Mon_Imp, ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Cuentas_Cobrar.Mon_Net    * (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN -1 ELSE 1 END)))    AS  Mon_Net ")
            loComandoSeleccionar.AppendLine(" INTO		#tmpTemporalCobros ")
            loComandoSeleccionar.AppendLine(" FROM		Cobros, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Cobros, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine(" WHERE		Cobros.Documento                =   Renglones_Cobros.Documento ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Tip      =   Renglones_Cobros.Cod_Tip ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Documento    =   Renglones_Cobros.Doc_Ori ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Cli      =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" GROUP BY  Cobros.Documento, Cobros.Status ")


            loComandoSeleccionar.AppendLine(" SELECT    #tmpTemporalCobros.Status, ")
            loComandoSeleccionar.AppendLine(" 			#tmpTemporalCobros.Documento, ")
            loComandoSeleccionar.AppendLine(" 			#tmpTemporalCobros.Mon_Bru As Mon_Bru02, ")
            loComandoSeleccionar.AppendLine(" 			#tmpTemporalCobros.Mon_Rec, ")
            loComandoSeleccionar.AppendLine(" 			#tmpTemporalCobros.Mon_Des, ")
            loComandoSeleccionar.AppendLine(" 			#tmpTemporalCobros.Mon_Imp, ")
            loComandoSeleccionar.AppendLine(" 			#tmpTemporalCobros.Mon_Net, ")
            loComandoSeleccionar.AppendLine(" 			Cobros.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 			Cobros.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 			Cobros.Cod_Mon, ")
            loComandoSeleccionar.AppendLine(" 			Cobros.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 			Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Tip_Ope     AS  Tip_Ope, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Mon_Net     AS  Cob_Ope, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Doc_Des     AS  Doc_Des, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Num_Doc     AS  Num_Doc, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Cod_Ban, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Cod_Caj, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Cod_Cue ")
            loComandoSeleccionar.AppendLine(" FROM		#tmpTemporalCobros, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine(" 			Cobros, ")
            loComandoSeleccionar.AppendLine(" 			Vendedores, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros, ")
            loComandoSeleccionar.AppendLine(" 			Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE		Cobros.Documento                =   #tmpTemporalCobros.Documento ")
            loComandoSeleccionar.AppendLine("           AND Cobros.Documento            =   Detalles_Cobros.Documento ")
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Cli              =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Ven              =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Mon              =   Monedas.Cod_Mon ")

            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
			
		'-------------------------------------------------------------------------------------------------------
        ' Verifica si el select (tabla nº0) trae registros
        '-------------------------------------------------------------------------------------------------------
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
						vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", "200px")
            
            Else

				
			'---------------------------------------------------------------'	 
			' Si el cobro no está confirmado no permite que se imprima		'
			'---------------------------------------------------------------'
				Dim lcEstatus	As String = Strings.Trim(laDatosReporte.Tables(0).Rows(0).Item("Status"))
				Dim lcDocumento	As String = Strings.Trim(laDatosReporte.Tables(0).Rows(0).Item("Documento"))
				
				If (lcEstatus.ToLower() <> "confirmado") Then
					
					Dim lcSalida = Me.Request.QueryString("salida")						
					If (lcSalida = "html") Then
					
						Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Operación no Permitida", _
									  "El Cobro '" & lcDocumento & "' no puede imprimirse porque no está Confirmado.", _
									   vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia)
						
						Return 
						
					Else
						
						Me.Response.Redirect(Me.Request.Url.ToString().Replace("salida=" & lcSalida, "salida=html"), True)
						
					End If
					
				End If
				
            End If


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAnexo_DCancelar_CCMV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfAnexo_DCancelar_CCMV.ReportSource = loObjetoReporte

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
' JJD: 06/05/10: Programacion inicial
'-------------------------------------------------------------------------------------------'
