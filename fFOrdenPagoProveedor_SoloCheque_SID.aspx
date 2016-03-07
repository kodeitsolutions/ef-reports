<%@ Page Language="VB" AutoEventWireup="false" CodeFile="fFOrdenPagoProveedor_SoloCheque_SID.aspx.vb" Inherits="fFOrdenPagoProveedor_SoloCheque_SID" %>
<!DOCTYPE html>
<%@ Register Assembly="vis3Controles" Namespace="vis3Controles" TagPrefix="vis3Controles" %>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>Formato de Cheque Banesco (SID - Solo Cheque)</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        
		<vis3Controles:wbcAdministradorMensajeModal ID="wbcAdministradorMensajeModal" runat="server" />
		<vis3Controles:wbcAdministradorVentanaModal ID="wbcAdministradorVentanaModal" runat="server" />
		<vis3Controles:pnlVentanaModal ID="PnlVentanaModalOperacion" runat="server" pcEstiloBotonCerrar="BotonCerrarVentanaModal" pcEstiloFondo="FondoVentanaModal" pcEstiloMarco="MarcoVentanaModal" pcTextoBotonCerrar="Cerrar" plMostrarBotonCerrar="false" poAlto="520px" poAncho="550px" Style="left: 386px; top: 28px" />
		<vis3Controles:pnlMensajeModal ID="PnlMensajeModalOperacion" runat="server" pcEstiloContenido="ContenidoMensajeModal" pcEstiloFondo="FondoVentanaModal" pcEstiloTitulo="TituloMensajeModal" pcEstiloVentana="MarcoMensajeModal" poAlto="400px" poAncho="750px" poArriba="20%" poIzquierda="30%" Style="left: 386px; top: 28px" />

    </div>
    </form>

</body>
</html>
