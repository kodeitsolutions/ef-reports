Imports System.Data
Partial Class fOrdenes_Pagos_FSV1
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Ordenes_Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Nom_Pro        As  Nombre_Generico, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Rif            As  Rif_Genenerico, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Nit            As  Nit_Generico, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Dir_Fis        As  Dir_Fis_Generico, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Telefonos      As  Telefonos_Generico, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Bru       As  Mon_Bru_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Imp1      As  Mon_Imp1_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Net       As  Mon_Net_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Ret       As  Mon_Ret_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Motivo        As  Motivo, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con + Substring(Renglones_oPagos.Comentario,1,250)    As  Nom_Con, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Deb    As  Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Hab    As  Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Net    As  Mon_Net_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Por_Imp1   As  Por_Imp_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Cod_Imp    As  Cod_Imp_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Comentario As  Comentario_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Imp1   As  Mon_Imp_Ren ")
            loComandoSeleccionar.AppendLine(" FROM      Ordenes_Pagos, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Conceptos ")
            loComandoSeleccionar.AppendLine(" WHERE     Ordenes_Pagos.Documento =   Renglones_oPagos.Documento AND ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Cod_Pro   =   Proveedores.Cod_Pro AND ")
            loComandoSeleccionar.AppendLine("           Conceptos.Cod_Con       =   Renglones_oPagos.Cod_Con AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")
            
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOrdenes_Pagos_FSV1", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfOrdenes_Pagos_FSV1.ReportSource = loObjetoReporte

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
' MAT: 15/03/11: Ajuste del Select
'-------------------------------------------------------------------------------------------'
' MAT: 18/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
