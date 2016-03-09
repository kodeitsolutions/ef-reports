'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rMCajas_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rMCajas_Fechas

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cajas.Documento, ")
            loComandoSeleccionar.AppendLine("		   Movimientos_Cajas.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("		   Cajas.mon_max, ")
            loComandoSeleccionar.AppendLine("		   Movimientos_Cajas.Comentario, ")
            loComandoSeleccionar.AppendLine("		   Movimientos_Cajas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("		   Movimientos_Cajas.Status, ")
            loComandoSeleccionar.AppendLine("		   Movimientos_Cajas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("		   Movimientos_Cajas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("		   SUBSTRING(Movimientos_Cajas.Tip_Ori,1,20) AS Tip_Ori, ")
            loComandoSeleccionar.AppendLine("		   Movimientos_Cajas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("		   Conceptos.Nom_Con, ")
            loComandoSeleccionar.AppendLine("		   Cajas.Nom_Caj   AS  Nombre ")
            loComandoSeleccionar.AppendLine("FROM Movimientos_Cajas ")
            loComandoSeleccionar.AppendLine(" JOIN Cajas ON  Movimientos_Cajas.Cod_Caj		=   Cajas.Cod_Caj ")
            loComandoSeleccionar.AppendLine(" JOIN Conceptos ON Movimientos_Cajas.Cod_Con   =   Conceptos.Cod_Con")
            loComandoSeleccionar.AppendLine(" WHERE     Movimientos_Cajas.Tipo          IN  ('Efectivo','Ticket') AND Movimientos_Cajas.tip_ori IN ('Pagos') ")
            loComandoSeleccionar.AppendLine("		   And Movimientos_Cajas.Cod_Caj       = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("		   And Movimientos_Cajas.Fec_Ini       = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("		   And Movimientos_Cajas.Status        = " & lcParametro2Desde)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            'Me.mEscribirConsulta(loCOmandoSeleccionar.ToString())

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")


            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
      vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_rMCajas_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rMCajas_Fechas.ReportSource = loObjetoReporte

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
' JJD: 24/01/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' GCR: 19/03/09: Adicion de informacion de renglones y modificaciones al diseño
'-------------------------------------------------------------------------------------------'
' CMS: 03/05/10: Se Ajusto para tomar el nombre del proveedor generico y que el debe sea (+)
'					y el haber (-)
'-------------------------------------------------------------------------------------------'
' MAT: 15/03/11: Ajuste del Select
'-------------------------------------------------------------------------------------------'
' MAT: 30/04/11: Corrección para que muestre el Nombre dle Beneficiario
'-------------------------------------------------------------------------------------------'
' JJD: 10/03/14: Ajustes a los campos del formato. Se incluye el comentario del renglon
'-------------------------------------------------------------------------------------------'