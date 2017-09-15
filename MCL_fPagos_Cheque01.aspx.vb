'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "MCL_fPagos_Cheque01"
'-------------------------------------------------------------------------------------------'
Partial Class MCL_fPagos_Cheque01
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT Pagos.Cod_Pro               AS  Cod_Pro,")
            loComandoSeleccionar.AppendLine("       Proveedores.Nom_Pro         AS  Nom_Pro,")
            loComandoSeleccionar.AppendLine("       Pagos.Documento             AS  Documento,")
            loComandoSeleccionar.AppendLine("       Pagos.Fec_Ini               AS  Fec_Ini,")
            loComandoSeleccionar.AppendLine("       Pagos.Mon_Bru			    AS  Mon_Bru_Enc,")
            loComandoSeleccionar.AppendLine("       (Pagos.Mon_Des * -1)        AS  Mon_Des,")
            loComandoSeleccionar.AppendLine("       Pagos.Mon_Net			    AS  Mon_Net_Enc,")
            loComandoSeleccionar.AppendLine("       (Pagos.Mon_Ret * -1)	    AS  Mon_Ret_Enc,")
            loComandoSeleccionar.AppendLine("       Pagos.Comentario		    AS  Comentario,")
            loComandoSeleccionar.AppendLine("       Detalles_Pagos.Num_Doc		AS  Referencia, ")
            loComandoSeleccionar.AppendLine("       Cuentas_Bancarias.Num_Cue	AS  Num_Cue_Cheque, ")
            loComandoSeleccionar.AppendLine("       Bancos.Nom_Ban				AS  Nom_Ban_Che, ")
            loComandoSeleccionar.AppendLine("       Renglones_Pagos.Cod_Tip     AS  Cod_Tip,")
            loComandoSeleccionar.AppendLine("       Renglones_Pagos.Doc_Ori     AS  Doc_Ori,")
            loComandoSeleccionar.AppendLine("       Renglones_Pagos.Renglon     AS  Renglon,")
            loComandoSeleccionar.AppendLine("       Renglones_Pagos.Mon_Bru     AS  Mon_Bru,")
            loComandoSeleccionar.AppendLine("       Renglones_Pagos.Mon_Imp     AS  Mon_Imp,")
            loComandoSeleccionar.AppendLine("       Renglones_Pagos.Mon_Abo     AS  Mon_Abo,")
            loComandoSeleccionar.AppendLine("       Renglones_Pagos.Mon_Net     AS  Mon_Net_Ren,")
            loComandoSeleccionar.AppendLine("       CAST('' as char(400))       AS  Mon_Let, ")
            loComandoSeleccionar.AppendLine("       Campos_Propiedades.Val_Car  AS  Nom_Cob ")
            loComandoSeleccionar.AppendLine("FROM Pagos ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Pagos ON Pagos.Documento = Renglones_Pagos.Documento")
            loComandoSeleccionar.AppendLine("   JOIN Proveedores ON Pagos.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("   JOIN Detalles_Pagos ON (Pagos.Documento = Detalles_Pagos.Documento ")
            loComandoSeleccionar.AppendLine("       AND Detalles_Pagos.Tip_Ope = 'Cheque' ")
            loComandoSeleccionar.AppendLine("       AND Detalles_Pagos.Renglon = '1')")
            loComandoSeleccionar.AppendLine("   JOIN Cuentas_Bancarias ON (Cuentas_Bancarias.Cod_Cue = Detalles_Pagos.Cod_Cue) ")
            loComandoSeleccionar.AppendLine("   JOIN Bancos ON (Bancos.Cod_ban = Cuentas_Bancarias.Cod_Ban) ")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Campos_Propiedades ON Pagos.Documento = Campos_Propiedades.Cod_Reg ")
            loComandoSeleccionar.AppendLine("       AND Campos_Propiedades.Cod_Pro  = 'NOMCOBCHE2' ")
            loComandoSeleccionar.AppendLine("       AND Campos_Propiedades.Origen = 'Pagos' ")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net_Enc"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

            Next loFilas


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("MCL_fPagos_Cheque01", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvMCL_fPagos_Cheque01.ReportSource = loObjetoReporte

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
' JJD: 13/05/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 11/03/15: Ajuste para el cliente MERCALUM
'-------------------------------------------------------------------------------------------'
' RJG: 09/04/15: Continuacion de los ajustes para el cliente CEGASA: número de cheque, banco'
'                y cuenta. Ajuste de posicion de etiquetas. Número de factura de proveedor. '
'-------------------------------------------------------------------------------------------'
