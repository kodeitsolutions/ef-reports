'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "MCL_fOrdenes_Cheque01"
'-------------------------------------------------------------------------------------------'
Partial Class MCL_fOrdenes_Cheque01
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Ordenes_Pagos.Cod_Pro                                               AS  Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro                                                 AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Documento                                             AS  Documento, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Ini                                               AS  Fec_ini, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Bru                                               AS  Mon_Bru_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Imp                                               AS  Mon_Imp1_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Net                                               AS  Mon_Net_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Ret                                               AS  Mon_Ret_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Motivo                                                AS  Motivo, ")
            loComandoSeleccionar.AppendLine("           Detalles_Opagos.Renglon			                                    AS  Renglon_Cheque, ")
            loComandoSeleccionar.AppendLine("           Detalles_Opagos.Num_Doc			                                    AS  Referencia, ")
            loComandoSeleccionar.AppendLine("           Detalles_Opagos.Fec_ini			                                    AS  Fec_Ini_Cheque, ")
            loComandoSeleccionar.AppendLine("           Detalles_Opagos.Mon_Net			                                    AS  Mon_Net_Cheque, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Nom_Cue		                                    AS  Nom_Cue_Cheque, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Num_Cue		                                    AS  Num_Cue_Cheque, ")
            loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban					                                    AS  Nom_Ban_Che, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Cod_Con                                            AS  Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con + SUBSTRING(Renglones_oPagos.Comentario,1,250)    AS  Nom_Con, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Renglon                                            AS  Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Deb                                            AS  Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Hab                                            AS  Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN Renglones_oPagos.Mon_Deb = 0 ")
            loComandoSeleccionar.AppendLine("            	 THEN Renglones_oPagos.Mon_Net * -1 ")
            loComandoSeleccionar.AppendLine("            	 ELSE Renglones_oPagos.Mon_Net ")
            loComandoSeleccionar.AppendLine("           END                                                                 AS  Mon_Net_Ren, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN Renglones_oPagos.Mon_Deb = 0  ")
            loComandoSeleccionar.AppendLine("            	 THEN Renglones_oPagos.Mon_Imp1 * -1 ")
            loComandoSeleccionar.AppendLine("            	 ELSE Renglones_oPagos.Mon_Imp1 ")
            loComandoSeleccionar.AppendLine("           END                                                                 AS  Mon_Imp_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Cod_Imp                                            AS  Cod_Imp_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Imp1                                           AS  Mon_Imp_Ren, ")
            loComandoSeleccionar.AppendLine("           CAST('' as char(400))                                               AS  Mon_Let, ")
            loComandoSeleccionar.AppendLine("           Campos_Propiedades.Val_Car                                          AS  Nom_Cob ")
            loComandoSeleccionar.AppendLine("FROM Ordenes_Pagos ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_oPagos ON Ordenes_Pagos.Documento = Renglones_oPagos.Documento ")
            loComandoSeleccionar.AppendLine("   JOIN Proveedores ON Ordenes_Pagos.Cod_Pro = Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("   JOIN Conceptos ON Conceptos.Cod_Con = Renglones_oPagos.Cod_Con ")
            loComandoSeleccionar.AppendLine("   JOIN Detalles_Opagos ON (Ordenes_Pagos.Documento = Detalles_Opagos.Documento ")
            loComandoSeleccionar.AppendLine("       AND Detalles_Opagos.Tip_Ope =   'Cheque' ")
            loComandoSeleccionar.AppendLine("       AND Detalles_Opagos.Renglon =   '1')")
            loComandoSeleccionar.AppendLine("   JOIN Cuentas_Bancarias ON (Cuentas_Bancarias.Cod_Cue = Detalles_Opagos.Cod_Cue) ")
            loComandoSeleccionar.AppendLine("   JOIN Bancos ON (Bancos.Cod_Ban = Cuentas_Bancarias.Cod_Ban) ")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Campos_Propiedades ON Ordenes_Pagos.Documento = Campos_Propiedades.Cod_Reg ")
            loComandoSeleccionar.AppendLine("       AND Campos_Propiedades.Cod_Pro = 'NOMCOBCHE1' ")
            loComandoSeleccionar.AppendLine("       AND Campos_Propiedades.Origen = 'Ordenes_Pagos' ")
            loComandoSeleccionar.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)





            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net_Enc"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

            Next loFilas


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("MCL_fOrdenes_Cheque01", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvMCL_fOrdenes_Cheque01.ReportSource = loObjetoReporte

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
' JFP: 28/11/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 03/05/10: Se Ajusto para tomar el debe sea (+) y el haber (-)
'-------------------------------------------------------------------------------------------'
' CMS: 30/06/10: Se Ajusto para tomar el Proveedor generico
'-------------------------------------------------------------------------------------------'
' JJD: 09/03/15: Se ajusto el formato para CGS
'-------------------------------------------------------------------------------------------'
' RJG: 09/04/15: Continuacion de los ajustes para el cliente CEGASA: número de cheque, banco'
'                y cuenta. Ajuste de posicion de etiquetas.                                 '
'-------------------------------------------------------------------------------------------'
