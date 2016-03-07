'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rVendedores_VNumericos"
'-------------------------------------------------------------------------------------------'
Partial Class rVendedores_VNumericos
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			Dim loComandoSeleccionar As New StringBuilder()

	Try	

			loComandoSeleccionar.AppendLine("SELECT	cod_Ven, " )
			loComandoSeleccionar.AppendLine("nom_Ven, " )
			loComandoSeleccionar.AppendLine("por_ven, " )
			loComandoSeleccionar.AppendLine("por_cob, " )
			loComandoSeleccionar.AppendLine("por_des, " )
			loComandoSeleccionar.AppendLine("por_imp, " )
			loComandoSeleccionar.AppendLine("por_rec, " )
			loComandoSeleccionar.AppendLine("por_ret, " )
			loComandoSeleccionar.AppendLine("por_lic, " )
			loComandoSeleccionar.AppendLine("por_men, " )
			loComandoSeleccionar.AppendLine("por_mas, " )
			loComandoSeleccionar.AppendLine("por_otr1, " )
			loComandoSeleccionar.AppendLine("por_otr2, " )
			loComandoSeleccionar.AppendLine("por_otr3, " )
			loComandoSeleccionar.AppendLine("por_otr4, " )
			loComandoSeleccionar.AppendLine("por_otr5, " )
			loComandoSeleccionar.AppendLine("mon_otr1, " )
			loComandoSeleccionar.AppendLine("mon_otr2, " )
			loComandoSeleccionar.AppendLine("mon_otr3, " )
			loComandoSeleccionar.AppendLine("mon_otr4, " )
			loComandoSeleccionar.AppendLine("mon_otr5, " )
			loComandoSeleccionar.AppendLine("Status, " )
			loComandoSeleccionar.AppendLine("(Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Vend " )
			loComandoSeleccionar.AppendLine("FROM	 vendedores " )
			loComandoSeleccionar.AppendLine("WHERE cod_ven between " & lcParametro0Desde )
			loComandoSeleccionar.AppendLine(" And " & lcParametro0Hasta )
			loComandoSeleccionar.AppendLine(" And vendedores.Status IN (" & lcParametro1Desde & ")" )
			'loComandoSeleccionar.AppendLine(" ORDER BY Cod_Ven, Nom_Ven " )
			loComandoSeleccionar.AppendLine("ORDER BY  " & lcOrdenamiento)
	
			Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")
 
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rVendedores_VNumericos", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrVendedores_VNumericos.ReportSource = loObjetoReporte
			
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
' YJP:  29/04/09: Codigo inicial
'-------------------------------------------------------------------------------------------'

