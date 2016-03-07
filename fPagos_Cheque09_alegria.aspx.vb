'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPagos_Cheque09_alegria"
'-------------------------------------------------------------------------------------------'
Partial Class fPagos_Cheque09_alegria
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	    Pagos.Cod_Pro,")
            'loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine("           (CASE WHEN RTRIM(Pagos.Proyecto)>''")
            loComandoSeleccionar.AppendLine("               THEN Pagos.Proyecto")
            loComandoSeleccionar.AppendLine("               ELSE Proveedores.Nom_Pro")
            loComandoSeleccionar.AppendLine("           END)                        AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif,")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit,")
            loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis,")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos,")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax,")
            loComandoSeleccionar.AppendLine("           Pagos.Documento,")
            loComandoSeleccionar.AppendLine("           Pagos.Fec_Ini,")
            loComandoSeleccionar.AppendLine("           Pagos.Fec_Fin,")

            loComandoSeleccionar.AppendLine("           detalles_pagos.Renglon		AS  Renglon_Cheque, ")
            loComandoSeleccionar.AppendLine("           detalles_pagos.Num_Doc		AS  Referencia, ")
            loComandoSeleccionar.AppendLine("           detalles_pagos.Fec_ini		AS  Fec_Ini_Cheque, ")
            loComandoSeleccionar.AppendLine("           detalles_pagos.Mon_Net		AS  Mon_Net_Cheque, ")
            loComandoSeleccionar.AppendLine("           detalles_pagos.Cod_Cue		AS  Cod_Cue_Che, ")
            loComandoSeleccionar.AppendLine("           cuentas_bancarias.Nom_Cue	AS  Nom_Cue_Cheque, ")
            loComandoSeleccionar.AppendLine("           bancos.Nom_Ban				AS  Nom_Ban_Che, ")


            loComandoSeleccionar.AppendLine("           Pagos.Mon_Bru			    AS  Mon_Bru_Enc,")
            loComandoSeleccionar.AppendLine("           (Pagos.Mon_Des * -1)        AS  Mon_Des,")
            loComandoSeleccionar.AppendLine("           Pagos.Mon_Net			    AS  Mon_Net_Enc,")
            loComandoSeleccionar.AppendLine("           (Pagos.Mon_Ret * -1)	    AS  Mon_Ret_Enc,")
            loComandoSeleccionar.AppendLine("           Pagos.Comentario		    AS  Comentario,")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Cod_Tip     AS  Cod_Tip,")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Doc_Ori     AS  Doc_Ori,")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Renglon     AS  Renglon,")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Bru     AS  Mon_Bru,")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Imp     AS  Mon_Imp,")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Abo     AS  Mon_Abo,")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Net     AS  Mon_Net_Ren,")
            loComandoSeleccionar.AppendLine("           CAST('' AS CHAR(400))       AS  Mon_Let ")

            loComandoSeleccionar.AppendLine("FROM      Pagos")
            loComandoSeleccionar.AppendLine("    JOIN Renglones_Pagos ON (Pagos.Documento =   Renglones_Pagos.Documento)")
            loComandoSeleccionar.AppendLine("    JOIN detalles_pagos ON (Pagos.Documento =   detalles_pagos.Documento AND detalles_pagos.tip_ope = 'Cheque' AND detalles_pagos.Renglon = '1')")
            loComandoSeleccionar.AppendLine("    JOIN Proveedores ON (Pagos.Cod_Pro   =   Proveedores.Cod_Pro) ")
            loComandoSeleccionar.AppendLine("    JOIN cuentas_bancarias ON (cuentas_bancarias.Cod_Cue   =   detalles_pagos.Cod_Cue) ")
            loComandoSeleccionar.AppendLine("    JOIN bancos ON (bancos.Cod_ban   =   cuentas_bancarias.Cod_Ban) ")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net_Enc"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

            Next loFilas


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPagos_Cheque09_alegria", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPagos_Cheque09_alegria.ReportSource = loObjetoReporte

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
' GCR: 31/03/09: Codigo inicial.
'-------------------------------------------------------------------------------------------'
' JJD: 06/02/10: Asignacion del signo negativo de la Retencion.
'-------------------------------------------------------------------------------------------'
' JJD: 30/03/10: Asignacion del signo negativo del Descuento.
'-------------------------------------------------------------------------------------------'
' RJG: 16/07/13: Se modificó para tomar el nombre del proveedor del campo Pagos.Proyecto.
'-------------------------------------------------------------------------------------------'
