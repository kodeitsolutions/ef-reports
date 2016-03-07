'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fSeriales_TAlmacenes"
'-------------------------------------------------------------------------------------------'
Partial Class fSeriales_TAlmacenes
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSelect As New StringBuilder()

            loComandoSelect.AppendLine(" SELECT Traslados.Documento, ")
            loComandoSelect.AppendLine("        Traslados.Status, ")
            loComandoSelect.AppendLine("        Traslados.Fec_Ini, ")
            loComandoSelect.AppendLine("        Traslados.Fec_Fin, ")
            loComandoSelect.AppendLine("        Traslados.Alm_Ori, ")
            loComandoSelect.AppendLine("        Traslados.Alm_Des, ")
            loComandoSelect.AppendLine("        Traslados.Cod_Tra, ")
            loComandoSelect.AppendLine("        Traslados.Comentario, ")
            loComandoSelect.AppendLine("        Renglones_Traslados.Renglon, ")
            loComandoSelect.AppendLine("        Renglones_Traslados.Cod_Art, ")
            loComandoSelect.AppendLine("        Renglones_Traslados.Cod_Uni, ")
            loComandoSelect.AppendLine("        Renglones_Traslados.Can_Art1, ")
            loComandoSelect.AppendLine("        Renglones_Traslados.Notas, ")
            loComandoSelect.AppendLine("        Transportes.Nom_Tra, ")
            loComandoSelect.AppendLine("        Almacenes.Nom_Alm As Nom_Ori, ")
            loComandoSelect.AppendLine("  		Seriales.Origen,")
            loComandoSelect.AppendLine("  		Seriales.Renglon AS Renglon_Serial,")
            loComandoSelect.AppendLine("  		Seriales.Cod_Art AS Cod_Art_Serial,")
            loComandoSelect.AppendLine("  		Seriales.Nom_Art AS Nom_Art_Serial,")
            loComandoSelect.AppendLine("  		Seriales.Serial,")
            loComandoSelect.AppendLine("  		Seriales.Alm_Ent,")
            loComandoSelect.AppendLine("  		Seriales.Tip_Ent,")
            loComandoSelect.AppendLine("  		Seriales.Doc_Ent,")
            loComandoSelect.AppendLine("  		Seriales.Ren_Ent,")
            loComandoSelect.AppendLine("  		Seriales.Alm_Sal,")
            loComandoSelect.AppendLine("  		Seriales.Tip_Sal,")
            loComandoSelect.AppendLine("  		Seriales.Doc_Sal,")
            loComandoSelect.AppendLine("  		Seriales.Ult_Tip,")
            loComandoSelect.AppendLine("  		Seriales.Ult_Doc,")
            loComandoSelect.AppendLine("  		Seriales.Ren_Sal,")
            loComandoSelect.AppendLine("  		Seriales.Garantia,")
            loComandoSelect.AppendLine("  		Seriales.Fec_Ini AS Fec_Fin_Serial,")
            loComandoSelect.AppendLine("  		Seriales.Fec_Fin AS Fec_Ini_Serial,")
            loComandoSelect.AppendLine("  		Seriales.Disponible,")
            loComandoSelect.AppendLine("  		Seriales.Comentario AS Comentario_Serial ")
            loComandoSelect.AppendLine(" INTO   #tblTemporal ")
            loComandoSelect.AppendLine(" FROM 	Traslados, ")
            loComandoSelect.AppendLine("        Renglones_Traslados, ")
            loComandoSelect.AppendLine("        Transportes, ")
            loComandoSelect.AppendLine("        Almacenes, ")
            loComandoSelect.AppendLine("        Seriales ")
            loComandoSelect.AppendLine(" WHERE	Traslados.Documento     =   Renglones_Traslados.Documento ")
            loComandoSelect.AppendLine("        AND Traslados.Cod_Tra   =   Transportes.Cod_Tra ")
            loComandoSelect.AppendLine("        AND Traslados.Alm_Ori   =   Almacenes.Cod_Alm ")
            loComandoSelect.AppendLine("        AND Seriales.Ult_Tip    =   'Traslados' ")
            loComandoSelect.AppendLine("        AND Seriales.Ult_Doc    =   Traslados.Documento ")
            loComandoSelect.AppendLine("        AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            loComandoSelect.AppendLine(" SELECT #tblTemporal.Documento, ")
            loComandoSelect.AppendLine("        #tblTemporal.Status, ")
            loComandoSelect.AppendLine("        #tblTemporal.Fec_Ini, ")
            loComandoSelect.AppendLine("        #tblTemporal.Fec_Fin, ")
            loComandoSelect.AppendLine("        #tblTemporal.Alm_Ori, ")
            loComandoSelect.AppendLine("        #tblTemporal.Alm_Des, ")
            loComandoSelect.AppendLine("        #tblTemporal.Cod_Tra, ")
            loComandoSelect.AppendLine("        #tblTemporal.Comentario, ")
            loComandoSelect.AppendLine("        #tblTemporal.Renglon, ")
            loComandoSelect.AppendLine("        #tblTemporal.Cod_Art, ")
            loComandoSelect.AppendLine("        #tblTemporal.Cod_Uni, ")
            loComandoSelect.AppendLine("        #tblTemporal.Can_Art1, ")
            loComandoSelect.AppendLine("        #tblTemporal.Notas, ")
            loComandoSelect.AppendLine("        #tblTemporal.Nom_Tra, ")
            loComandoSelect.AppendLine("        #tblTemporal.Nom_Ori, ")
            loComandoSelect.AppendLine("        Almacenes.Nom_Alm As Nom_Des, ")
            loComandoSelect.AppendLine("  		#tblTemporal.Origen,")
            loComandoSelect.AppendLine("  		#tblTemporal.Renglon_Serial,")
            loComandoSelect.AppendLine("  		#tblTemporal.Cod_Art_Serial,")
            loComandoSelect.AppendLine("  		#tblTemporal.Nom_Art_Serial,")
            loComandoSelect.AppendLine("  		#tblTemporal.Serial,")
            loComandoSelect.AppendLine("  		#tblTemporal.Alm_Ent,")
            loComandoSelect.AppendLine("  		#tblTemporal.Tip_Ent,")
            loComandoSelect.AppendLine("  		#tblTemporal.Doc_Ent,")
            loComandoSelect.AppendLine("  		#tblTemporal.Ren_Ent,")
            loComandoSelect.AppendLine("  		#tblTemporal.Alm_Sal,")
            loComandoSelect.AppendLine("  		#tblTemporal.Tip_Sal,")
            loComandoSelect.AppendLine("  		#tblTemporal.Doc_Sal,")
            loComandoSelect.AppendLine("  		#tblTemporal.Ult_Tip,")
            loComandoSelect.AppendLine("  		#tblTemporal.Ult_Doc,")
            loComandoSelect.AppendLine("  		#tblTemporal.Ren_Sal,")
            loComandoSelect.AppendLine("  		#tblTemporal.Garantia,")
            loComandoSelect.AppendLine("  		#tblTemporal.Fec_Fin_Serial,")
            loComandoSelect.AppendLine("  		#tblTemporal.Fec_Ini_Serial,")
            loComandoSelect.AppendLine("  		#tblTemporal.Disponible,")
            loComandoSelect.AppendLine("  		#tblTemporal.Comentario_Serial ")
            loComandoSelect.AppendLine(" FROM 	#tblTemporal, ")
            loComandoSelect.AppendLine("        Almacenes ")
            loComandoSelect.AppendLine(" WHERE	#tblTemporal.Alm_Des   =   Almacenes.Cod_Alm")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSelect.ToString, "curReportes")

   			'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fSeriales_TAlmacenes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfSeriales_TAlmacenes.ReportSource = loObjetoReporte

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
' CMS: 10/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 30/03/10: Ajustes al codigo. Reemplazo de Objeto 'WITH()'
'-------------------------------------------------------------------------------------------'