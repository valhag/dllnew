using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Configuration;
using System.IO;
//using BarradeProgreso;
using System.Data.SqlClient ;
using Interfaces ;
using System.Collections;
using System.Data;
using System.Globalization;
using System.Net;
using System.Linq;
using System.Xml;
using System.Data.Odbc;

namespace LibreriaDoctos
{
    public class ClassBD: ISujeto
    {
        

        public List<string> lvar = new List<string>();
        protected decimal lsubtotal;
        protected decimal limpuestos;
        protected string aRutaExe;
        public string productos;
        public string almacenes;
        public bool sdkcomercial = false; 
        public RegDocto primerdocto = new RegDocto();
        List<IObservador> lista = new List<IObservador>();



        [DllImport("MGW_SDK.DLL")]        static extern int fInsertaCteProv();
        [DllImport("MGW_SDK.DLL")]        static extern int fEditaCteProv();
        [DllImport("MGW_SDK.DLL")]        static extern int fGuardaCteProv();
        [DllImport("MGW_SDK.DLL")]        static extern int fSetDatoCteProv(string aCampo, string aValor);

        [DllImport("MGW_SDK.DLL")]        static extern int fInsertaProducto();
        [DllImport("MGW_SDK.DLL")]        static extern int fGuardaProducto();
        [DllImport("MGW_SDK.DLL")]        static extern int fSetDatoProducto(string aCampo, string aValor);

        [DllImport("MGW_SDK.DLL")]        static extern int fInsertaAlmacen();
        [DllImport("MGW_SDK.DLL")]        static extern int fGuardaAlmacen();
        [DllImport("MGW_SDK.DLL")]        static extern int fSetDatoAlmacen(string aCampo, string aValor);

        [DllImport("MGW_SDK.DLL")] static extern int fInsertarDocumento();
        [DllImport("MGW_SDK.DLL")] static extern int fGuardaDocumento();
        [DllImport("MGW_SDK.DLL")] static extern int fSetDatoDocumento(string aCampo, string aValor);
        [DllImport("MGW_SDK.DLL")]        static extern int fInsertarMovimiento();
        [DllImport("MGW_SDK.DLL")]        static extern int fGuardaMovimiento();
        [DllImport("MGW_SDK.DLL")]        static extern int fSetDatoMovimiento(string aCampo, string aValor);
        [DllImport("MGW_SDK.DLL")]      static extern int fInsertaDireccion();

        [DllImport("MGW_SDK.DLL")]        static extern int fGuardaDireccion();
        [DllImport("MGW_SDK.DLL")]        static extern int fBorraDocumento();
        //[DllImport("MGW_SDK.DLL")]        static extern int fBorraMovimiento();


        [DllImport("MGW_SDK.DLL")]        static extern int fSetDatoDireccion(string aCampo, string aValor);
        [DllImport("MGW_SDK.DLL")]        static extern int fAfectaDocto_Param(string aConcepto, string aSerie, double aFolio, Boolean aAfecta);
        [DllImport("MGW_SDK.DLL")]        static extern long fError(long aNumErrror, string aError, long aLen);

        [DllImport("MGW_SDK.DLL")]
        static extern int fSiguienteFolio(string lCodigoConcepto, ref string lSerie, ref double lFolio);

        [DllImport("MGW_SDK.DLL")]
        static extern int fInicializaLicenseInfo(int aSistema);

        //Private Declare Function fEmitirDocumento Lib "MGW_SDK.DLL" (ByVal aCodigoConcepto As String, ByVal aNumSerie As String, ByVal aFolio As Double, ByVal aPassword As String, ByVal aArchivo As String) As Long
        [DllImport("MGW_SDK.DLL")]
        static extern int fEmitirDocumento(string aCodigoConcepto, string aNumSerie, double aFolio, string aPassword, string aArchivo);

        [DllImport("MGW_SDK.DLL")]
        static extern int fEntregEnDiscoXML(string aCodigoConcepto,string aNumSerie, double aFolio, int aFormato, ref string aFormatoamigo);

        //(aCodConcepto, aSerie, aFolio, aFormato, aFormatoAmig)
        //lError = fEntregEnDiscoXML (“4”, “B1”, 45, 1, “C:\Compacw\Empresas\Reportes\AdminPAQ\Plantilla_Factura_cfdi_1.html”)

        [DllImport("MGW_SDK.DLL")]
        static extern long fSaldarDocumento_Param(string lCodConcepto_Pagar, string lSerie_Pagar, double lFolio_Pagar,
string lCodConcepto_Pago, string lSerie_Pago, double lFolio_Pago, double lImporte, int lIdMoneda, string lFecha);


        [DllImport("MGW_SDK.DLL")]
        static extern long fRegresaExistencia(string lCodigoProducto, string lCodigoAlmacen, string lAnio, string lMes, string lDia, ref double lExistencia);


        // Need this DllImport statement to reset the floating point register below
        [DllImport("msvcr71.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int _controlfp(int n, int mask);

        [DllImport("KERNEL32.DLL")]
        static extern int SetCurrentDirectory(string pPtrDirActual);
        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fSetNombrePAQ(string aSistema);
        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fError(int aNumError, string aMensaje, int aLen);
        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fTerminaSDK();
        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fPosPrimerEmpresa(ref int aIdEmpresa, ref string aNombreEmpresa, ref string aDirectorioEmpresa);

        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fAbreEmpresa(string aDirectorioEmpresa);

        [DllImport("MGWSERVICIOS.DLL")]
        static extern void fCierraEmpresa();

        [DllImport("MGWSERVICIOS.DLL", EntryPoint="fSiguienteFolio")]
        static extern void fSiguienteFolioComercial (string aCodigoConcepto, ref string aSerie, ref double aFolio);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fInsertarDocumento")]
        static extern int fInsertarDocumentoComercial();

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fSetDatoDocumento")]
        static extern int fSetDatoDocumentoComercial(string aCampo, string aValor);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fGuardaDocumento")]
        static extern int fGuardaDocumentoComercial();

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fError")]
        static extern int fErrorComercial(int aNumError, StringBuilder aMensaje, int aLen);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fAltaDocumento")]
        static extern int fAltaDocumentoComercial(ref long aIdDocumento, TDocumento aDocumento);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fAltaMovimientoSeriesCapas_Param")] 
        static extern int fAltaMovimientoSeriesCapas_ParamComercial (string aIdMovimiento, string aUnidades, string  aTipoCambio,  string aSeries,
 string aPedimento,  string aAgencia,  string aFechaPedimento,  string aNumeroLote,  string aFechaFabricacion,  string aFechaCaducidad);


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fAltaMovimientoSeriesCapas_Param")]
        static extern int fAltaMovimientoSeriesCapas_ParamComercial1(string aIdMovimiento, string aUnidades, string aTipoCambio, string aSeries,
 string aPedimento, string aAgencia, string aFechaPedimento, string aNumeroLote, string aFechaFabricacion, string aFechaCaducidad);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fInsertaProducto")]
        static extern int fInsertaProductoComercial();

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fGuardaProducto")] 
        static extern int fGuardaProductoComercial();
        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fSetDatoProducto")] 
        static extern int fSetDatoProductoComercial(string aCampo, string aValor);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fBuscaProducto")]
        static extern int fBuscaProductoComercial(string aCampo);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fInsertaCteProv")]
        static extern int fInsertaCteProvComercial();

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fGuardaCteProv")]
        static extern int fGuardaCteProvComercial();
        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fSetDatoCteProv")]
        static extern int fSetDatoCteProvComercial(string aCampo, string aValor);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fBuscaCteProv")]
        static extern int fBuscaCteProvComercial(string aCampo);


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fInsertaDireccion")]
        static extern int fInsertaDireccionComercial();

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fGuardaDireccion")]
        static extern int fGuardaDireccionComercial();
        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fSetDatoDireccion")]
        static extern int fSetDatoDireccionComercial(string aCampo, string aValor);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fBuscaDireccion")]
        static extern int fBuscaDireccionComercial(string aCampo);


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fInsertarMovimiento")]
        static extern int fInsertarMovimientoComercial();
        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fGuardaMovimiento")]
        static extern int fGuardaMovimientoComercial();

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fEditarMovimiento")]
        static extern int fEditarMovimientoComercial();

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fSetDatoMovimiento")]
        static extern int fSetDatoMovimientoComercial(string aCampo, string aValor);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fBuscarDocumento")]
        static extern int fBuscarDocumentoComercial(string aCodConcepto, string aSerie, string aFolio);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fLeeDatoDocumento")]
        static extern int fLeeDatoDocumentoComercial(string aCampo, StringBuilder aMensaje, int aLen);


        //public static extern void fError(int NumeroError, StringBuilder Mensaje, int Longitud);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fLeeDatoMovimiento")]
        static extern int fLeeDatoMovimientoComercial(string aCampo, StringBuilder aMensaje, int aLen);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fModificaCostoEntrada")]
        static extern int fModificaCostoEntradaComercial(string aIdMovimiento, string aCostoEntrada);


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fBuscarIdMovimiento")]
        static extern int fBuscarIdMovimientoComercial(int aIdMovimeinto);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fBorraDocumento")]
        static extern int fBorraDocumentoComercial();


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fInicializaLicenseInfo")]
        static extern int fInicializaLicenseInfoComercial(int aSistema);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fEmitirDocumento")]
        static extern int fEmitirDocumentoComercial(string aCodigoConcepto, string aNumSerie, double aFolio, string aPassword, string aArchivo);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fLeeDatoCteProv")]
        static extern int fLeeDatoCteProvComercial(string aCampo, StringBuilder aMensaje, int aLen);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fEntregEnDiscoXML")]
        static extern int fEntregEnDiscoXMLComercial(string aCodigoConcepto, string aNumSerie, double aFolio, int aFormato, string aFormatoAmig);

        //fEntregEnDiscoXML (aCodConcepto, aSerie, aFolio, aFormato, aFormatoAmig)
        

        //protected ClassConexion miconexion;
         public ClassConexion  miconexion = new ClassConexion();
        public RegDocto _RegDoctoOrigen = new RegDocto();
        private string _rfc;
        private string _razonsocial;
        //public const string _NombreAplicacionCompleto = "Grid.exe";
        //public const string _NombreAplicacion = "Grid";
        public const string _NombreAplicacionCompleto = "InterfazAdmin.exe";
        public const string _NombreAplicacion = "InterfazAdmin";
        
        public List<RegDocto> _RegDoctos = new List<RegDocto>();
        //protected OleDbConnection _con;

        public List<RegOrigen> list2 = new List<RegOrigen>();


        RegDocto p777 = new RegDocto();
        RegDocto p888 = new RegDocto();
        RegDocto p999 = new RegDocto();


        protected OleDbConnection  _con;


        public void mLlenarinfoMicroplane()
        {

            //List<RegDocto> doctos = new List<RegDocto>();

            /*SqlConnection lconexionOrigen = new SqlConnection();

            lconexionOrigen = miconexion.mAbrirConexionSQLOrigen();
             */

            string dsn = GetSettingValueFromAppConfigForDLL("databaseOrigen"); 

            OdbcConnection DbConnection = new OdbcConnection(dsn);
            try
            {
                DbConnection.Open();
            }
            catch (Exception eeeee)
            {
                return;
            }
             
 

            Boolean noseguir = false;

            string ssql = " SELECT " +
            " inv_dt, " +
            " h.inv_no, " +
            " c.cmp_code,c.cmp_name, c.textfield1, c.TaxCode   " +
            " , h.curr_cd, h.curr_trx_rt " +
            " , h.bill_to_addr_1,  h.bill_to_city, h.bill_to_country,   h.bill_to_no, h.bill_to_state, h.bill_to_zip " +
            " , l.item_no, l.item_desc_1, l.unit_price, l.discount_pct, l.qty_ordered, l.qty_to_ship, p.item_note_1, p.item_note_5 " +
            " FROM oehdrhst_sql h " +
            " join cicmpy c on c.cmp_code = h.cus_no " +
            " join oelinhst_sql l on l.inv_no = h.inv_no " +
            " join imitmidx_sql p on p.item_no = l.item_no " +
            " where h.inv_no > 7130 " +
            " order by h.inv_no asc ";
            /*
            SqlCommand lsql = new SqlCommand(ssql, DbConnection);

            SqlDataReader dr = lsql.ExecuteReader();*/

            OdbcCommand DbCommand = DbConnection.CreateCommand();
            DbCommand.CommandText = ssql;
            OdbcDataReader dr = DbCommand.ExecuteReader();


            _RegDoctos.Clear();
            RegDocto lDocto = new RegDocto();
            if (dr.HasRows)
            {
                string clienteleido = "";
                long folioleido = 0;
                long lfolio = 0;
                string cserie = "";
                string lConcepto = GetSettingValueFromAppConfigForDLL("Concepto");
                //long lFoliox = mBuscarUltimoFolioConcepto("4", GetSettingValueFromAppConfigForDLL("Concepto"), ref cserie);
                while (noseguir == false)
                {

                    if (dr.Read() == true)
                    {

                        string lcliente = dr["cmp_code"].ToString();

                        if (lcliente == "")
                            break;

                        lfolio = long.Parse(dr["inv_no"].ToString());


                        //if (lcliente != clienteleido)
                        if (lfolio != folioleido)
                        {
                            if (lDocto.cCodigoCliente != "")
                            {
                                _RegDoctos.Add(lDocto);
                                lDocto = new RegDocto();
                            }


                            //lDocto.cSerie = cserie;
                            lDocto.cCodigoCliente = dr["cmp_code"].ToString();
                            lDocto.cRazonSocial = dr["cmp_code"].ToString();
                            lDocto._RegCliente.Codigo = dr["cmp_code"].ToString();
                            lDocto._RegCliente.RazonSocial = dr["cmp_name"].ToString();
                            lcliente = lDocto.cCodigoCliente;
                            lDocto.cCodigoConcepto = lConcepto;
                            //lDocto.cMetodoPago = "02";

                            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");
                            lDocto.cFolio = long.Parse(dr["inv_no"].ToString());

                            lDocto.cFecha = DateTime.Parse(dr["inv_dt"].ToString());

                            lDocto.cFecha = DateTime.Today;


                            clienteleido = lcliente;
                            folioleido = lfolio;
                            lDocto.cMoneda = dr["curr_cd"].ToString();
                            lDocto.cTipoCambio = decimal.Parse(dr["curr_trx_rt"].ToString());

                        }

                        RegMovto regmov = new RegMovto();
                        regmov.cCodigoProducto = dr["item_no"].ToString();
                        regmov._RegProducto.Nombre = dr["item_desc_1"].ToString();

                        regmov._RegProducto.noIdentificacion = dr["item_note_1"].ToString();
                        regmov._RegProducto.CodigoMedidaPesoSAT = dr["item_note_5"].ToString(); ;

                        regmov.cPorcent01 = decimal.Parse(dr["discount_pct"].ToString());
                        regmov.cUnidades = decimal.Parse(dr["qty_to_ship"].ToString());
                        regmov.cCodigoAlmacen = "1";
                        regmov.cPrecio = decimal.Parse(dr["unit_price"].ToString());
                        lDocto._RegMovtos.Add(regmov);
                    }
                    else
                        noseguir = true;

                }
            }
            dr.Close();

            
        }

        public void mLlenarinfoXML(string archivo)
        {

            //List<RegDocto> doctos = new List<RegDocto>();


            _RegDoctos.Clear();

            DirectoryInfo di = new DirectoryInfo(@archivo);
            //Console.WriteLine("No search pattern returns:");
            foreach (var fi in di.GetFiles("*.xml"))
            {

                RegDocto lDocto = new RegDocto();
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(archivo + "\\" + fi.Name);

                XmlNodeList xComprobante = xDoc.GetElementsByTagName("cfdi:Comprobante");

                foreach (XmlElement nodo in xComprobante)
                {
                    lDocto.cFecha = DateTime.Parse(nodo.GetAttribute("fecha"));
                    lDocto.cTipoCambio = Decimal.Parse(nodo.GetAttribute("TipoCambio"));
                    lDocto.cMoneda = nodo.GetAttribute("Moneda");
                    lDocto.cMetodoPago = nodo.GetAttribute("metodoDePago");
                }

                XmlNodeList xEmisor = ((XmlElement)xComprobante[0]).GetElementsByTagName("cfdi:Emisor");
                XmlNodeList xReceptor = ((XmlElement)xComprobante[0]).GetElementsByTagName("cfdi:Receptor");
                XmlNodeList xConceptos = ((XmlElement)xComprobante[0]).GetElementsByTagName("cfdi:Conceptos");


                foreach (XmlElement nodo in xEmisor)
                {
                    lDocto.cRFC = nodo.GetAttribute("rfc");
                    lDocto.cRazonSocial = nodo.GetAttribute("nombre");
                    //long lFoliox = mBuscarUltimoFolioConcepto("4", GetSettingValueFromAppConfigForDLL("Concepto"), ref cserie);
                    string cserie = "";
                    //long lFoliox = mBuscarUltimoFolioConcepto("4", "4", ref cserie);

                    lDocto.cCodigoCliente = nodo.GetAttribute("rfc");
                    lDocto.cCodigoConcepto = "4";
                    XmlNodeList xDomFiscal = ((XmlElement)nodo).GetElementsByTagName("cfdi:DomicilioFiscal");

                    foreach (XmlElement nodoDomFiscal in xDomFiscal)
                    {
                        lDocto._RegDireccion.cCodigoPostal = nodoDomFiscal.GetAttribute("codigoPostal");    
                    }

                    
                    XmlNodeList xRegFiscal = ((XmlElement)nodo).GetElementsByTagName("cfdi:RegimenFiscal");
                    foreach (XmlElement nodoRegFiscal in xRegFiscal)
                    {
                        lDocto.cRegimenFiscal = nodoRegFiscal.GetAttribute("Regimen");
                    }

                    
                    //lDocto.cFecha = 



                }

                foreach (XmlElement nodoReceptor in xReceptor)
                {
                    lDocto.cRFC = nodoReceptor.GetAttribute("rfc");
                    lDocto.cRazonSocial = nodoReceptor.GetAttribute("nombre");
                    XmlNodeList xDomicilioReceptor = ((XmlElement)nodoReceptor).GetElementsByTagName("cfdi:Domicilio");
                    foreach (XmlElement nodoDomicilioReceptor in xDomicilioReceptor)
                    {
                        lDocto._RegDireccion.cNombreCalle = nodoDomicilioReceptor.GetAttribute("calle");
                        lDocto._RegDireccion.cPais = nodoDomicilioReceptor.GetAttribute("pais");
                    }
                }


                foreach (XmlElement nodoConceptos in xConceptos)
                {

                    XmlNodeList xConcepto = ((XmlElement)nodoConceptos).GetElementsByTagName("cfdi:Concepto");
                    foreach (XmlElement nodoConcepto in xConcepto)
                    {
                        //RegMovto regmov = new RegMovto();
                        //regmov. = nodoConcepto.GetAttribute("importe");
                        //                        decimal xValorUnitario = decimal.Parse(nodoConcepto.GetAttribute("valorUnitario"));

                        RegMovto regmov = new RegMovto();
                        regmov._RegProducto.Nombre = nodoConcepto.GetAttribute("descripcion");
                        regmov.cUnidades = decimal.Parse(nodoConcepto.GetAttribute("cantidad"));
                        regmov.cCodigoAlmacen = "1";
                        regmov.cPrecio = decimal.Parse(nodoConcepto.GetAttribute("valorUnitario"));


                        int HashCode = regmov._RegProducto.Nombre.GetHashCode();

                        //regmov._RegProducto.Codigo = HashCode.ToString();

                        regmov._RegProducto.noIdentificacion = nodoConcepto.GetAttribute("noIdentificacion");
                        lDocto._RegMovtos.Add(regmov);


                    }
                }
                _RegDoctos.Add(lDocto);
            }
        }

        public List<RegConcepto> mCargarConceptosComercial(long aIdDocumentoDe, int aTipo, int cfdi)
        {
            List<RegConcepto> _RegFacturas = new List<RegConcepto>();
            SqlConnection lconexion = new SqlConnection();

            if (aTipo == 0)
                lconexion = miconexion.mAbrirConexionComercial(false);
            else
                lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {

                //OleDbCommand lsql = new OleDbCommand("select ccodigoc01,cnombrec01 from mgw10006 where ciddocum01 = " + aIdDocumentoDe + " and cescfd = 1 and cnombrec01 = 'CFDI'", lconexion);

                string sqlstring = "select ccodigoconcepto,cnombreconcepto,cverfacele from admConceptos where ciddocumentode = " + aIdDocumentoDe;
                if (cfdi == 1)
                    sqlstring = "select ccodigoconcepto,cnombreconcepto,cverfacele from admConceptos where ciddocumentode = " + aIdDocumentoDe + " and cescfd = 1";




                SqlCommand lsql = new SqlCommand();
                lsql.CommandText = sqlstring;
                lsql.Connection = lconexion;
                SqlDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _RegFacturas.Clear();
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegConcepto lRegConcepto = new RegConcepto();
                        lRegConcepto.Codigo = lreader[0].ToString();

                        lRegConcepto.Nombre = lreader[1].ToString();
                        //lRegConcepto.Tipocfd = lreader[2].ToString();
                        _RegFacturas.Add(lRegConcepto);
                    }
                }
                lreader.Close();
            }

            return _RegFacturas;



        }

        private string ObtenerNombreProducto(string codigo)
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            miconexion.mAbrirConexionDestino();
            lsql.CommandText = "select cnombrep01 from mgw10005 where ccodigop01 = '" + codigo+ "'";
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            string lregresa = "";
            if (lreader.HasRows)
            {
                string x;
                lreader.Read();
                try
                {
                    x = lreader[0].ToString();
                }
                catch (Exception ee)
                {
                    x = ee.Message;
                }
                lreader.Close();
                miconexion.mCerrarConexionDestino();
                lregresa = x;
            }
            return lregresa;

        }

        public RegAlmacen mBuscarAlmacenAsumidoComercial(string aCodigoConcepto)
        {
            miconexion.mAbrirConexionComercial(false);
            string lalmacen = "";
            RegAlmacen lAlmacen = new RegAlmacen();

            string lquery = "";

            SqlCommand lsql = new SqlCommand();
            SqlDataReader lreader;

            // miconexion.mAbrirConexionDestino();

            lsql.CommandText = "select CCODIGOALMACEN,a.cidalmacen from admconceptos c " +
            " join admAlmacenes a " +
            " on c.CIDALMASUM = a.CIDALMACEN " +
            " where CCODIGOCONCEPTO = '" + aCodigoConcepto + "'";
            lsql.Connection = miconexion._conexion1;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            string lregresa = "";
            string x = "";

            if (lreader.HasRows)
            {
                lreader.Read();
                try
                {
                    lAlmacen.Codigo = lreader[0].ToString();
                    lAlmacen.Id = long.Parse(lreader[1].ToString());
                }
                catch (Exception ee)
                {
                    //                    lreader.Close();
                }
                lreader.Close();
            }
            miconexion.mCerrarConexionOrigenComercial();
            return lAlmacen  ;
        }
        private string ObtenercodigoAlmacen(string textoextra1)
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            miconexion.mAbrirConexionDestino();
            lsql.CommandText = "select ccodigoa01 from mgw10003 where ctextoex01 = '" + textoextra1 + "'";
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            string lregresa = "";
            if (lreader.HasRows)
            {
                string x;
                lreader.Read();
                try
                {
                    x = lreader[0].ToString();
                }
                catch (Exception ee)
                {
                    x = ee.Message;
                }
                lreader.Close();
                miconexion.mCerrarConexionDestino();
                lregresa = x;
            }
            return lregresa;
            
        }
        /*
        OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + aNombreArchivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");

            conn.Open();
            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            cmd.CommandText = "SELECT * FROM Hoja1$]";
            cmd.ExecuteNonQuery();

            System.Data.OleDb.OleDbDataReader dr;
            dr = cmd.ExecuteReader();
            Boolean noseguir = false;
            _RegDoctos.Clear();
            if (dr.HasRows)
            {
                while (noseguir == false)
                {
                    dr.Read();
                }
            }*/

        public void mLlenarinfoFacturacionMasiva(string archivo)
        {
            //string archivo1 = @archivo;
            OleDbConnection conn1 = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @archivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");

            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @archivo + ";Extended Properties=" + Convert.ToChar(34).ToString() + @"Excel 8.0" + Convert.ToChar(34).ToString() + ";");

            conn.Open();
            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            cmd.CommandText = "SELECT * FROM [Datos$]";
            cmd.ExecuteNonQuery();

            System.Data.OleDb.OleDbDataReader dr;
            dr = cmd.ExecuteReader();
            Boolean noseguir = false;
            _RegDoctos.Clear();
            List<RegDocto> doctos = new List<RegDocto>();
            RegDocto lDocto = new RegDocto();
            if (dr.HasRows)
            {
                string clienteleido = "";
                string cserie = "";
                long lFoliox = mBuscarUltimoFolioConcepto("4",GetSettingValueFromAppConfigForDLL("Concepto"), ref cserie);
                while (noseguir == false)
                {

                    dr.Read();
                    string lcliente = dr["CODIGO DEL CLIENTE"].ToString();
                    if (lcliente == "")
                        break;
                    
                    
                    if (lcliente != clienteleido)
                    {
                        if (lDocto.cCodigoCliente != "")
                        {
                            _RegDoctos.Add(lDocto);
                            lDocto = new RegDocto();
                        }
                        
                        
                        lDocto.cSerie = cserie;
                        lDocto.cCodigoCliente = dr["CODIGO DEL CLIENTE"].ToString();
                        lcliente = lDocto.cCodigoCliente;
                        lDocto.cCodigoConcepto = "4";

                        lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");
                        lDocto.cFolio = lFoliox;
                        lFoliox++;
                        lDocto.cFecha = DateTime.Parse(DateTime.Today.ToString());
                        clienteleido = lcliente;
                    }
                    
                    RegMovto regmov = new RegMovto();
                    regmov.cCodigoProducto = dr["CODIGO DEL SERVICIO A FACTURAR"].ToString();
                    regmov.cUnidades = decimal.Parse(dr["CANTIDAD"].ToString());
                    regmov.cCodigoAlmacen = "1";
                    regmov.cPrecio = decimal.Parse(dr["PRECIO UNIT"].ToString());
                    regmov.cObservaciones = dr["DESCRIPCION"].ToString();
                    lDocto._RegMovtos.Add(regmov);

                }

                if (lDocto.cCodigoCliente != "")

                    _RegDoctos.Add(lDocto);
                 
            }

        }

        public int mLlenarinfoAmcoPedidos(string url)
        {
            List<RegElemento> lista = new List<RegElemento>();
            List<RegElemento> listanueva = new List<RegElemento>();
            WebClient wc = new WebClient();
            string ll = wc.DownloadString(url);
            string lnuevacadena = ll.Substring(1, ll.Length - 1);
            string lelemento = "";
            int lposactual = 0;
            int lposcoma = -1;
            int lseguir = 1;
            _RegDoctos.Clear();
            long lfolioleido = 0;
            //List<RegDocto> doctos = new List<RegDocto>();
            RegDocto lDocto = new RegDocto();
            while (lseguir == 1)
            {
                lposcoma = lnuevacadena.IndexOf("},", lposactual);

                try
                {
                    lelemento = lnuevacadena.Substring(lposactual, lposcoma - lposactual);
                }
                catch (Exception eee)
                {
                    lelemento = lnuevacadena.Substring(lposactual);
                    lseguir = 0;
                }

                RegElemento newelemento = new RegElemento();
                int linicio = lelemento.IndexOf("id") + 4;
                int lfin = lelemento.IndexOf(",");
                newelemento.id = lelemento.Substring(linicio, lfin - linicio);

                linicio = lelemento.IndexOf("school_id") + 11;
                lfin = lelemento.IndexOf(",", linicio);
                newelemento.school_id = lelemento.Substring(linicio, lfin - linicio);

                linicio = lelemento.IndexOf("code") + 7;
                lfin = lelemento.IndexOf(",", linicio);
                newelemento.code = lelemento.Substring(linicio, lfin - linicio - 1);

                linicio = lelemento.IndexOf("amount") + 8;
                lfin = lelemento.IndexOf(",", linicio);
                newelemento.amount = lelemento.Substring(linicio, lfin - linicio);

                linicio = lelemento.IndexOf("type") + 7;
                lfin = lelemento.IndexOf(",", linicio);
                newelemento.type = lelemento.Substring(linicio, lfin - linicio - 1);

                linicio = lelemento.IndexOf("unit") + 7;
                lfin = lelemento.IndexOf(",", linicio);
                newelemento.unit = lelemento.Substring(linicio, lfin - linicio - 1);

                linicio = lelemento.IndexOf("date") + 7;
                lfin = lelemento.IndexOf("\\", linicio);
                newelemento.date = lelemento.Substring(linicio);
                lposactual += linicio + newelemento.date.Length + 4;
                newelemento.date = newelemento.date.Substring(0, newelemento.date.Length - 1);

                listanueva.Add(newelemento);
            }

            List<RegElemento> listanueva2 = new List<RegElemento>();
            listanueva.Sort(delegate(RegElemento x, RegElemento y)
            {
                return x.id.CompareTo(y.id);
            });


            var result1 = listanueva.OrderBy(a => a.id);

            listanueva2 = result1.ToList();

            //listanueva2 = listanueva.Sort();

            foreach (RegElemento newelemento in listanueva2)
            {
                long lfolio = long.Parse(newelemento.id.ToString());
                if (lfolio != lfolioleido)
                {

                    if (lfolioleido != 0)
                    {
                        _RegDoctos.Add(lDocto);
                        lDocto = new RegDocto();
                    }

                    lDocto.cSerie = "";
                    lDocto.cCodigoCliente = newelemento.school_id;

                    lDocto.cCodigoCliente = "001";

                    // lDocto.cCodigoConcepto = "4";
                    // normal, muestra, extemporáneo, reposicion, devolucion, tiendita, faltante, sctijuana, brilliant
                    //lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");
                    string tipo = newelemento.type.Substring(0,5);
                    switch (tipo)
                    {
                        case "brill":
                            lDocto.cCodigoConcepto = "ped9";
                            break;
                        case "extem":
                            lDocto.cCodigoConcepto = "ped3";
                            break;
                        case "muest":
                            lDocto.cCodigoConcepto = "ped2";
                            break;
                        case "repos":
                            lDocto.cCodigoConcepto = "ped4";
                            break;
                        case "norma":
                            lDocto.cCodigoConcepto = "ped1";
                            break;
                        case "devol":
                            lDocto.cCodigoConcepto = "ped5";
                            break;
                        case "tiend":
                            lDocto.cCodigoConcepto = "ped6";
                            break;
                        case "falta":
                            lDocto.cCodigoConcepto = "ped7";
                            break;
                        case "sctij":
                            lDocto.cCodigoConcepto = "ped6";
                            break;    
                    }
                    //lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");
                    
                    lDocto.cFolio = long.Parse(newelemento.id.ToString());

                    lDocto.cFecha = DateTime.Parse(newelemento.date.Substring(0, 10));
                    lfolioleido = lDocto.cFolio;

                    
                }



                RegMovto regmov = new RegMovto();
                regmov.cCodigoProducto = newelemento.code;
                regmov.cUnidades = decimal.Parse(newelemento.amount.ToString());
                regmov.cCodigoAlmacen = "1";
                regmov.cPrecio = decimal.Parse("0");
                //regmov.cObservaciones = dr["DESCRIPCION"].ToString();
                lDocto._RegMovtos.Add(regmov);
                

                // id folio del pedido
                // school_id codigo cliente
                // code codigo del producto
                // amount cantidad
                // unit id unidad de medida y peso
                // date fecha del pedido
            }
            _RegDoctos.Reverse();
            return _RegDoctos.Count;

        }

        public void mLlenarinfo(string archivo, string Observaciones777, string Observaciones888, string txtObservaciones999, string Referencia, string ObservacionesMov, string refmovto777, string textoextra1777, string refmovto888, string textoextra1888, string refmovto999, string textoextra1999)
        {

            // crear 3 documetnos 1 para salida 777, salida 888 y remision 999

            p777 = new RegDocto();
            p888 = new RegDocto();
            p999 = new RegDocto();
            _RegDoctos.Clear();
            string cserie = "";
            p777.cFolio = mBuscarUltimoFolioConcepto("33", "506", ref cserie);
            p777.cSerie = cserie;
            p777.cCodigoCliente = "0";
            p777.cFecha = DateTime.Parse(DateTime.Today.ToString());
            p777.cCodigoConcepto = "506";
            p777.cCodigoCliente = "(Ninguno)";
            p777.cTextoExtra3 = Observaciones777;
            _RegDoctos.Add(p777);

            p888.cFolio = mBuscarUltimoFolioConcepto("33", "35", ref cserie) + 1;
            p888.cSerie = cserie;
            p888.cCodigoCliente = "0";
            p888.cFecha = DateTime.Parse(DateTime.Today.ToString());
            p888.cCodigoConcepto = "35";
            p888.cCodigoCliente = "(Ninguno)";
            p888.cTextoExtra3 = Observaciones888;
            _RegDoctos.Add(p888);

            p999.cFolio = mBuscarUltimoFolioConcepto("3", "3", ref cserie);
            p999.cSerie = cserie;
            p999.cCodigoCliente = "75";
            p999.cFecha = DateTime.Parse(DateTime.Today.ToString());
            p999.cCodigoConcepto = "1";
            p999.cTextoExtra3 = txtObservaciones999;
            p999.cReferencia = Referencia;
            
            _RegDoctos.Add(p999);
            string archivo1 = @archivo;
            OleDbConnection connect = new OleDbConnection();
            connect.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.GetDirectoryName(archivo1) +";Extended Properties='Text;HDR=Yes;FMT=Delimited;IMEX=1';Persist Security Info=False";
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [" + Path.GetFileName(archivo1) + "]", connect);
            DataTable dt = new DataTable();
            da.Fill(dt);


            
            string almacen777 = ObtenercodigoAlmacen("777");
            string almacen888 = ObtenercodigoAlmacen("888");
            string almacen999 = ObtenercodigoAlmacen("999");
        

            foreach (DataRow row in dt.Rows)
            {
                RegMovto regmov = new RegMovto();
                regmov.cCodigoProducto = row[1].ToString();
                regmov.cNombreProducto = ObtenerNombreProducto(regmov.cCodigoProducto);
                regmov.cUnidades = decimal.Parse(row[2].ToString());
                switch (row[0].ToString())
                {
                    case "777":
                        regmov.cCodigoAlmacen = almacen777;
                        regmov.cReferencia = refmovto777;
                        regmov.ctextoextra1 = textoextra1777;
                        p777._RegMovtos.Add(regmov);
                        break;
                    case "888":
                        regmov.cCodigoAlmacen = almacen888;
                        regmov.cReferencia = refmovto888;
                        regmov.ctextoextra1 = textoextra1888;
                        p888._RegMovtos.Add(regmov);
                        break;
                    case "999":
                        regmov.ctextoextra3 = ObservacionesMov;
                        regmov.cCodigoAlmacen = almacen999; // salida
                        regmov.cAlmacenEntrada = "999"; // entrada
                        regmov.cReferencia = refmovto999;
                        regmov.ctextoextra1 = textoextra1999;
                        p999._RegMovtos.Add(regmov);
                        break;

                }

                
            }
           
        }
        private string mGrabarEncabezado(double aFolio, string lCodigoConcepto)
        {
            long lret, lidconce = 0, tipocfd = 0;
            string cserie = "";

            //lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            int lnat = 0;
            mRegresarPrincipales(lCodigoConcepto, ref lidconce, ref tipocfd, ref cserie,ref lnat);
            string lresp = mValidarExisteDoc(lidconce, cserie, aFolio);
            if (lresp != "")
                return lresp;

            fInsertarDocumento();
            lret = fSetDatoDocumento("cCodigoConcepto", lCodigoConcepto);
            if (_RegDoctoOrigen.cSerie != "")
                lret = fSetDatoDocumento("cSerieDocumento", _RegDoctoOrigen.cSerie);
            else
                lret = fSetDatoDocumento("cSerieDocumento", "");
            //            lret = fSetDatoDocumento("CCODIGOCLIENTE", _RegDoctoOrigen.cCodigoCliente);
            lret = fSetDatoDocumento("cCodigoCteProv", _RegDoctoOrigen.cCodigoCliente);
            lresp = "";
            if (lret != 0)
            {
                lresp = mGrabarCliente();
            }
            if (lresp != "")
                return lresp;

            lret = fSetDatoDocumento("cRazonSocial", _RegDoctoOrigen.cRazonSocial);
            lret = fSetDatoDocumento("cRFC", _RegDoctoOrigen.cRFC);
            if (_RegDoctoOrigen.cMoneda != "Pesos")
                lret = fSetDatoDocumento("cIdMoneda", "2");
            else
                lret = fSetDatoDocumento("cIdMoneda", "1");
            lret = fSetDatoDocumento("cTipoCambio", _RegDoctoOrigen.cTipoCambio.ToString());
            lret = fSetDatoDocumento("cReferencia", "Por Programa");
            //lret = fSetDatoDocumento("cObservaciones", _RegDoctoOrigen.cTextoExtra1 );
            lret = fSetDatoDocumento("cFolio", _RegDoctoOrigen.cFolio.ToString().Trim());


            try
            {
                //lret = fSetDatoDocumento("cReferencia", _RegDoctoOrigen.cFolio.ToString());
                //lret = fSetDatoDocumento("cTextoExtra1", _RegDoctoOrigen.cReferencia);
                lret = fSetDatoDocumento("cReferencia", _RegDoctoOrigen.cReferencia);
                lret = fSetDatoDocumento("cTextoExtra2", _RegDoctoOrigen.cTextoExtra2);
                lret = fSetDatoDocumento("cTextoExtra3", _RegDoctoOrigen.cTextoExtra3);
            }
            catch (Exception ee)
            {
            }


            DateTime lFechaVencimiento;
            lFechaVencimiento = _RegDoctoOrigen.cFecha.AddDays(int.Parse("0"));
            //lFechaVencimiento = DateTime.Today.AddDays(int.Parse(_RegDoctoOrigen.cCond) );

            string lfechavenc = "";
            lfechavenc = String.Format("{0:MM/dd/yyyy}", lFechaVencimiento); ;  // "8 08 008 2008"   year
            lret = fSetDatoDocumento("cFechaVencimiento", lfechavenc);
            try
            {
                lret = fSetDatoDocumento("cCodigoAgente", _RegDoctoOrigen.cAgente);
            }
            catch (Exception e)
            {

            }
            /*
            if (lret != 0)
            {
                miconexion.mCerrarConexionOrigen(1);
                _controlfp(0x9001F, 0xFFFFF);
                // barra.Asignar(100);
                return "Fecha Incorrecta";
            }
             * 
             */




            string lfechadocto = "";
            lfechadocto = _RegDoctoOrigen.cFecha.ToString();
            DateTime lFechaDocto;
            lFechaDocto = _RegDoctoOrigen.cFecha;

            lfechadocto = "";


            lfechadocto = String.Format("{0:MM/dd/yyyy}", lFechaDocto); ;  // "8 08 008 2008"   year
            lret = fSetDatoDocumento("cFecha", lfechadocto);

            lret = fSetDatoDocumento("cFechaVencimiento", lfechadocto);
            lret = fSetDatoDocumento("cTipoCambio", "1");
            lret = fGuardaDocumento();
            string serror = "";
            if (lret != 0)
            {
                //fError(lret, serror, 255);
                _controlfp(0x9001F, 0xFFFFF);
                miconexion.mCerrarConexionOrigen(1);
                return lret.ToString() + " Documento ya Existe";


            }
            return "";



        }

        public string mGrabarAdm(string afolioant, double afolionuevo, int opcion, int tipo)
        {
            miconexion.mAbrirConexionDestino(1);
            string lCodigoConcepto;
            //lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            lCodigoConcepto = _RegDoctoOrigen.cCodigoConcepto;

            string lresp1 = mGrabarEncabezado(afolionuevo, lCodigoConcepto);
            if (lresp1 != "")
                return lresp1;

            //OleDbCommand lsql = new OleDbCommand();
            //OleDbDataReader lreader;

            string cserie;
            cserie = _RegDoctoOrigen.cSerie;
            long lIdDocumento = mBuscarIdDocumento(lCodigoConcepto, 0, cserie, long.Parse(afolionuevo.ToString().Trim()));


            if (lIdDocumento == 0)
            {

                _controlfp(0x9001F, 0xFFFFF);
                miconexion.mCerrarConexionOrigen(1);
                return "no se encontro documento " + lCodigoConcepto + " " +
                    long.Parse(GetSettingValueFromAppConfigForDLL("FolioDestino").ToString().Trim());
            }

            string lresp = mGrabarDireccion(lIdDocumento);
            lresp = mGrabarMovimientos(lIdDocumento, opcion, tipo);

            string lrespuestas = mGrabarExtras(lIdDocumento, 1, afolionuevo);

            int lret = fAfectaDocto_Param(lCodigoConcepto, cserie, afolionuevo, true);
            string lCodigoConceptoPago = "";
            if (_RegDoctoOrigen.cContado == 1)
            {
                lCodigoConceptoPago = "10";
            }
            if (_RegDoctoOrigen.cCodigoConcepto == "503")
            {
                lCodigoConceptoPago = "504";
            }
            lCodigoConceptoPago = "300";
            if (lCodigoConceptoPago != "300")
            {
                _RegDoctoOrigen.cCodigoConcepto = lCodigoConcepto;
                mGrabarEncabezado(afolionuevo, lCodigoConceptoPago);
                lIdDocumento = mBuscarIdDocumento(lCodigoConceptoPago, 0, cserie, long.Parse(afolionuevo.ToString().Trim()));
                _RegDoctoOrigen._RegMovtos.Clear();
                lresp = mGrabarMovimientos(lIdDocumento, 3,0);
                lrespuestas = mGrabarExtras(lIdDocumento, 1, afolionuevo);

                /*
                string lfechavenc = "";
                lfechavenc = String.Format("{0:MM/dd/yyyy}", _RegDoctoOrigen.cFecha); ;  // "8 08 008 2008"   year


                double importe = _RegDoctoOrigen.cNeto;
                string otroconcepto = "10";
                string sFolo = afolionuevo.ToString();

                lret = fAfectaDocto_Param("10", cserie, afolionuevo, true);
                long lret1 = fSaldarDocumento_Param (lCodigoConcepto, cserie, afolionuevo,
otroconcepto, cserie, afolionuevo, importe, 1, lfechavenc);
                 */
            }




            miconexion.mCerrarConexionOrigen(1);
            //miconexion.mCerrarConexionDestino(1);

            try
            {
                _controlfp(0x9001F, 0xFFFFF);
            }
            catch (Exception eee)
            {
                lrespuestas = eee.Message;
            }
            // barra.Asignar(100);
            return lrespuestas;
        }

        public List<string> mGrabarDoctos(bool incluyetimbrado, int tipo)
        {
            List<string> lista = new List<string>();
            miconexion.mAbrirConexionDestino(1);
            string lresp2 = "";
            foreach (RegDocto x in _RegDoctos)
            {
                if (x._RegMovtos.Count > 0)
                {
                    _RegDoctoOrigen = x;
                    lresp2 = mGrabarAdmNew(x.cFolio, 1, incluyetimbrado, tipo);
                    if (lresp2 != "")
                        lista.Add(lresp2);
                }
            }
            miconexion.mCerrarConexionOrigen(1);
            return lista;
        }

        public string mGrabarDescuentos()
        {

            return "";
        }

        //CONTPAQiComercial.Comercial comComercialMain ;
        //CONTPAQiComercial.TTInterfazTabla gTablas;

        public void mCargaCom()
        {
           /* string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");


            string lruta2 = Directory.GetCurrentDirectory();

            comComercialMain = new CONTPAQiComercial.Comercial();
            gTablas = new CONTPAQiComercial.TTInterfazTabla();
            string sURLSACI = "http://127.0.0.1:9080/saci/adminpaq";

            string szRegKeyAdminPAQ2001 = lruta2;
            string lRuta = lruta2 + @"\ContPAQiComercial.exe";

            string gGuidCom = "";
            string lConfig = "";
            int lError1 = gTablas.InicializarComunicacion(sURLSACI, "CONTPAQ I Comercial", 10, "", out gGuidCom, out lConfig);

            string kTokenSeparadorSACICOM = "¬";
            string lstrREsultado = "";
            int lResultado = 0;

            comComercialMain.ProcesarUnaFuncion("SetUrl", sURLSACI + kTokenSeparadorSACICOM + gGuidCom, "", out lstrREsultado, out lResultado);


            string sUsuario = "SUPERVISOR";
            string sPassword = "";
            int aIdUsuario = 0;
            string aNombreUsuario = "";
            int aPerfilUsuario = 0;
            string aListaEstados = "";
            string aListaPersmisos = "";
            string aListaDescripciones = "";
            int lResultado1 = 0;

            comComercialMain.seguridadValidaUsuario(1, sUsuario, sPassword, out aIdUsuario, out aNombreUsuario, out aPerfilUsuario, out aListaPersmisos, out aListaEstados, out aListaDescripciones, out lResultado1);
            //comComercialMain
            */
        }

        public ClassBD()
        {
          //  miconexion = new ClassConexion();
           _con = new OleDbConnection ();

           

        }

        public string mGrabarAbono(string lConcepto, int lDocumentoModelo)
        {
            
            miconexion.mAbrirConexionDestino(1);

            double afolionuevo;
            afolionuevo = mRegresarFolio(lDocumentoModelo);
            _RegDoctoOrigen.cFolio = long.Parse(afolionuevo.ToString());
            string lresp1 = mGrabarEncabezado(afolionuevo, lConcepto, "0");
            if (lresp1 != "")
                return lresp1;
            long lIdDocumento = mBuscarIdDocumento(lConcepto, 0, "", long.Parse(afolionuevo.ToString().Trim()));
            _RegDoctoOrigen._RegMovtos.Clear();
            string lresp = mGrabarMovimientos(lIdDocumento, 3,0);
            string lrespuestas = mGrabarExtras(lIdDocumento, 3, afolionuevo);

            string lresp10 = mGrabarTablaAdicional();
            return lrespuestas;

          
        }

        public string mGrabarTablaAdicional()
        {
                //string lcadena212 = "insert into ncprod values (''," + _RegDoctoOrigen.cFolio.ToString() + "," +   ;

                //OleDbCommand lsql414 = new OleDbCommand(lcadena212, miconexion._conexion);
                //lsql4.ExecuteNonQuery();

            string lcadena212 = "";
            foreach (RegOrigen regs in list2)
            {

                lcadena212 = "insert into ncprod values (''," + _RegDoctoOrigen.cFolio.ToString() + "," + regs.cidproducto.ToString() + "," + regs.Cantidad.ToString() + "," + regs.Precio.ToString() + "," + regs.Precio2.ToString() + "," + regs.IEPS.ToString() + "," + regs.IEPS2.ToString() + "," + regs.Descuento.ToString() + "," + regs.cIdClien01.ToString() + "," + regs.TotalMov.ToString() + "," + regs.DescuentoAplicar.ToString() + ')';
                miconexion.mAbrirConexionDestino();
                OleDbCommand lsql414 = new OleDbCommand(lcadena212, miconexion._conexion);
                lsql414.ExecuteNonQuery();
                miconexion.mCerrarConexionDestino();

            }
            return "";
                
        }

        private double mRegresarFolio(int lDocumentoModelo )
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            
        //    miconexion.mAbrirConexionDestino();

            lsql.CommandText = "select max(cfolio) as zzz from mgw10008 where ciddocum02 = " + lDocumentoModelo;
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            double lregresa = 0;
            if (lreader.HasRows)
            {
                long x;
                lreader.Read();
                try
                {
                  x = long.Parse(lreader[0].ToString());
                }
                catch(Exception ee)
                {
                    x = 0;
                }
                lreader.Close();
                    
                    lregresa = x + 1;
            }
            return lregresa;
        }

        public void mAsignaRuta(string aRuta)
        {
            aRutaExe = aRuta;
            miconexion.aRutaExe = aRuta;
        }

        protected virtual string mRegresarConsultaMovimientos(string aFuente, string lfolio, int ltipo)
        {
            string lregresa = "";
            switch (aFuente)
            {
                case "Flex":
                    lregresa = "select f.itemcode as ccodigop01,FCUnitPrice  as cprecioc01, " +
                    " BillTaxPerc as cporcent01,  '1' as ccodigoa01, p.ItemDesc  as cnombrep01, f.priceunitcode  as Unidad, " +
                    " IVAxLin as cimpuesto1, TotxLinea as cneto,  TotalxLineaIVA as ctotal, Cantidad as unidades   " +
                    " , isnull(f.itemdesc,'') as ctextoextra2, isnull(f.itemcode,'') as ctextoextra3, isnull(f.SHIPPINGRE,'') as creferen01 , isnull(f.CUSTITEMREF,'') as ctextoextra1" +
                    " from facturacione f join PM_Item p " +
                    " on f.ItemCode = p.ItemCode " +
                    " where f.billnum = " + lfolio;
                    break;
                case "Mercado":
                    lregresa = " select vd.articulo as ccodigop01,  " +
                    " Precio as cprecioc01, Cantidad as cunidades, vd.Impuesto1 as cimpuesto1,  Almacen as ccodigoa01, a.Descripcion1 as cnombrep01, a.Unidad " +
                    " from VentaD VD  join Art a  " +
                    " on VD.Articulo = a.Articulo   " +
                    " where ID = " + lfolio;

                    lregresa = " select vd.articulo as ccodigop01,  " +
                    " 'cprecioc01' =  case  " +
                    " when vd.impuesto1 <> 0 then round((Precio / (1 + (vd.impuesto1/100) )),4)  " +
                    " when vd.impuesto1 = 0 then Precio  " +
                    " end   " +
                    " , Cantidad as cunidades, vd.Impuesto1 as cimpuesto1,  Almacen as ccodigoa01, a.Descripcion1 as cnombrep01, a.Unidad, vd.impuesto1 as cPorcent01 " +
                    " from VentaD VD  join Art a  " +
                    " on VD.Articulo = a.Articulo   " +
                    " where ID = " + lfolio;



                    break;

            }
            return lregresa;
        }

        protected virtual Boolean mchecarvalido()
        {
            return true;
            //if (_RegDoctoOrigen.cFecha > DateTime.Parse("2011/08/01"))
            //    return false;
        }

        protected virtual string mModificaDatosCliente()
        {
            return "";
        }
        protected  string mModificaDatosClienteFlexo()
        {
            //return "";
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            string lrespuesta = "";
            long lidcliente = 0;
            miconexion.mAbrirConexionDestino();

            lsql.CommandText = "select * from mgw10002 where ccodigoc01 = '" + _RegDoctoOrigen.cCodigoCliente + "'";
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                lreader.Read();
                _RegDoctoOrigen.cCodigoCliente = lreader["ccodigoc01"].ToString();
                _RegDoctoOrigen.cRazonSocial = lreader["crazonso01"].ToString();
                _RegDoctoOrigen.cRFC = lreader["cRFC"].ToString();
                _RegDoctoOrigen.cCond = lreader["cdiascre01"].ToString();
                lidcliente = long.Parse(lreader["cidclien01"].ToString());



                lsql.CommandText = "select * from mgw10001 where cidagente = " + lreader["cidagent01"].ToString();
                lreader.Close();
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    _RegDoctoOrigen.cAgente = lreader["ccodigoa01"].ToString();
                }

                // ahora checar si tiene direccion fiscal si no la tiene avisar, si la tiene asignarla de adminpaq
                lreader.Close();


                lsql.CommandText = "select * from mgw10011 where ctipocat01 = 1 and cidcatal01 = " + lidcliente + " and ctipodir01 = 0";
                //lreader.Close();
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    _RegDoctoOrigen._RegDireccion.cNombreCalle = lreader["cnombrec01"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cNumeroExterior = lreader["cnumeroe01"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cNumeroInterior = lreader["cnumeroi01"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cColonia = lreader["ccolonia"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cCodigoPostal = lreader["ccodigop01"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cEstado = lreader["cestado"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cPais = lreader["cpais"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cCiudad = lreader["cciudad"].ToString().Trim();
                    lrespuesta = "";
                }

                else
                    lrespuesta = "Cliente sin direccion fiscal en ADMINPAQ";
            }
            else
                lrespuesta = "Cliente no existe en ADMINPAQ";

            lreader.Close();
            miconexion.mCerrarConexionDestino();
            return lrespuesta;


            //else
            //{
            //    lRespuesta = "El cliente no se ha dado de alta en Adminpaq "; // documento no encontrado
            //}

            //_con.Close();

            return "";
        }

        protected  string mLlenarDoctos(OleDbDataReader aReader)
        {
            _RegDoctos.Clear() ;
            string lfolio = "";
            aReader.Read();
            int lbandera = 1;
            while (lbandera == 1 && aReader.HasRows )
            {
                RegDocto x = new RegDocto ();
                List<RegMovto> movtos = new List<RegMovto> ();
                
                x.cAgente = "(Ninguno)";
                try
                {
                    x.cReferencia = aReader["cReferen01"].ToString();
                }
                catch (Exception dddd)
                {
 
                }
                x.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento");
                x.cFecha = DateTime.Parse (aReader["cfecha"].ToString());
                x.sMensaje = "";
                x.cMoneda = "Pesos";
                x.cTextoExtra1 = aReader["cObserva01"].ToString(); 

                string sfoliodocto = aReader["cfolio"].ToString();
                long lfoliodocto = 0 ;
                string lserie = "";
                try
                {
                    lfoliodocto = long.Parse( aReader["cfolio"].ToString());

                }
                catch (Exception eee)
                {
                    lserie = sfoliodocto.Substring(sfoliodocto.Length-1);
                    sfoliodocto = sfoliodocto.Substring(0, sfoliodocto.Length - 1);
                }

                
 
                x.cFolio = long.Parse (sfoliodocto );
                x.cSerie = lserie; 
                lfolio = aReader["cfolio"].ToString();
                while (lfolio == aReader["cfolio"].ToString())
                {
                    RegMovto mov = new RegMovto();
                    mov.cCodigoProducto = aReader["ccodigop01"].ToString();
                    mov.cNombreProducto = aReader["cnombrep01"].ToString();
                    mov.cCodigoAlmacen = "1";
                    mov.cUnidad = "PZA";
                    //" o.id_productos as ccodigop01, o.importe as cprecioc01, pr.nombreproducto as cnombrep01, o.cantidad as cunidades " +
                    mov.cUnidades = decimal.Parse(aReader["cunidades"].ToString());
                    mov.cPrecio = decimal.Parse(aReader["cprecioc01"].ToString());
                    movtos.Add(mov);
                    if (aReader.Read() == false)
                    {
                        lbandera = 0;
                        break;
                    }
                    //else
                        //lfolio = aReader["cfolio"].ToString();
                }
                x._RegMovtos = movtos;
                _RegDoctos.Add(x);
            }
            return "";

        }

        protected virtual string mLlenarDocto(OleDbDataReader aReader, int atipo, string aFolio, string aFuente)
        {
            string lrespuesta = "";
            string lfolio= "0";
            if (atipo == 1 || atipo == 2)
            {
                lfolio = aReader["cfolio"].ToString();
                _RegDoctoOrigen.cFolio  = long.Parse (lfolio);
            }
            if (aReader["cliente"].ToString() == string.Empty )
                return "Falta Codigo de cliente en documento " + aFolio ;
            else
                _RegDoctoOrigen.cCodigoCliente = aReader["cliente"].ToString();

            _RegDoctoOrigen.cFecha = DateTime.Parse(aReader["cfecha"].ToString());
            _RegDoctoOrigen.cFecha = DateTime.Parse(DateTime.Today.ToString ());
            if (mchecarvalido() == false)
                return "";


            
            //_RegDoctoOrigen.cFolio = long.Parse (aReader["cfolio"].ToString()) ;
            if (aReader["cRFC"].ToString() == string.Empty )
                return "Cliente sin RFC en documento " + aFolio;
            else
                if (!(aReader["cRFC"].ToString().Length == 12 ||  aReader["cRFC"].ToString().Length == 13))
                    return "El RFC tiene una longitud incorrecta en el documento " + aFolio;
                else
                    _RegDoctoOrigen.cRFC = aReader["cRFC"].ToString();
            
            
            if (atipo == 1)
            {
                _RegDoctoOrigen.cAgente = aReader["Agente"].ToString();
                _RegDoctoOrigen.cCond  = aReader["condpago"].ToString();

            }
            if (aReader["cRazonso01"].ToString() == string.Empty)
                return "Cliente sin Razon Social en documento " + aFolio;
            else
                _RegDoctoOrigen.cRazonSocial = aReader["cRazonso01"].ToString();

            //IsDBNull(
            //aReader["cTextoExtra1"].isnull
               // if(!aReader.IsDBNull(18))
                 //   _RegDoctoOrigen.cTextoExtra1 = aReader[18].ToString();



            // UNA modificacion que aplica para flexo es que los datos del cliente se toman de adminpaq
            lrespuesta = mModificaDatosCliente();
            //lrespuesta = mModificaDatosClienteFlexo();
            if (lrespuesta != string.Empty)
                return lrespuesta;


            _RegDoctoOrigen.cMoneda = aReader["Moneda"].ToString();
            _RegDoctoOrigen.cTipoCambio  = decimal.Parse (aReader["TipoCambio"].ToString());

            if (atipo != 1)
                _RegDoctoOrigen.cReferencia = aReader["cReferen01"].ToString();
            else
                _RegDoctoOrigen.cReferencia = aReader["cReferen01"].ToString();



            if (aReader["cnombrec01"].ToString().Trim() == string.Empty )
                _RegDoctoOrigen._RegDireccion.cNombreCalle = "Ninguna";
            else
                _RegDoctoOrigen._RegDireccion.cNombreCalle = aReader["cnombrec01"].ToString().Trim();

            _RegDoctoOrigen._RegDireccion.cNumeroExterior = aReader["cnumeroe01"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cNumeroInterior = aReader["cnumeroi01"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cColonia = aReader["ccolonia"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cCodigoPostal = aReader["ccodigop01"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cEstado = aReader["cestado"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cPais = aReader["cpais"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cCiudad = aReader["cciudad"].ToString().Trim();
            if (atipo == 3 || atipo == 4)
            {
                _RegDoctoOrigen.cNeto = double.Parse(aReader["importe"].ToString());
                _RegDoctoOrigen.cImpuestos = double.Parse(aReader["impuestos"].ToString().Trim());
            }

            
            OleDbCommand  lsql = new OleDbCommand ();
            OleDbDataReader   lreader;

            lsql.CommandText = mRegresarConsultaMovimientos(aFuente, lfolio, atipo );

            
            lsql.Connection = (OleDbConnection  )_con;
            aReader.Close();
            lreader = lsql.ExecuteReader();
            _RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                while (lreader.Read())
                {
                    RegMovto lRegmovto = new RegMovto();
                    lRegmovto.cCodigoProducto = lreader["ccodigop01"].ToString();
                    lRegmovto.cNombreProducto = lreader["cnombrep01"].ToString();
                    lRegmovto.cIdDocto = long.Parse(_RegDoctoOrigen.cIdDocto.ToString());
                    lRegmovto.cPrecio = decimal.Parse(lreader["cprecioc01"].ToString());
                    
                    lRegmovto.cImpuesto = decimal.Parse(lreader["cimpuesto1"].ToString());
                    lRegmovto.cPorcent01 = decimal.Parse(lreader["cPorcent01"].ToString());
                    if (aFuente != "Mercado")
                    {
                        lRegmovto.cUnidades = decimal.Parse(lreader["unidades"].ToString());
                        lRegmovto.cTotal = decimal.Parse(lreader["cTotal"].ToString());
                        lRegmovto.cneto = decimal.Parse(lreader["cneto"].ToString());
                        lRegmovto.cReferencia = lreader["creferen01"].ToString();
                        lRegmovto.ctextoextra1 = lreader["ctextoextra1"].ToString();
                        lRegmovto.ctextoextra2 = lreader["ctextoextra2"].ToString();
                        lRegmovto.ctextoextra3 = lreader["ctextoextra3"].ToString();
                        
                    }
                    else
                        lRegmovto.cUnidades = decimal.Parse(lreader["cunidades"].ToString());
                    lRegmovto.cCodigoAlmacen = lreader["ccodigoa01"].ToString();
                    lRegmovto.cNombreAlmacen = lreader["ccodigoa01"].ToString();
                    lRegmovto.cUnidad = lreader["unidad"].ToString();
                    _RegDoctoOrigen._RegMovtos.Add(lRegmovto);
                }

            }
            else
            { 
                
            }
            lreader.Close();
            return lrespuesta;
            //miconexion.mCerrarConexionOrigen(); 
        }




        //public boolean mBuscar(long aFolio, long aIdDocum02)
        public Boolean  mBuscar(long aFolio, string aConcepto, string aSerie, int aTipo)
        {
            Boolean lRespuesta = false;
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader ;
            OleDbParameter lparametrofolio = new OleDbParameter ("@p2",aFolio );
            OleDbParameter lparametrodocumentode = new OleDbParameter("@p1", aConcepto);

            lsql.CommandText = "Select m2.ccodigoc01 as cliente,m6.ccodigoc01 as concepto, m6.cidconce01, m8.cfecha,m8.cfolio, m8.ciddocum01 " +
                " from mgw10008 m8 join mgw10002 m2 on m8.cidclien01 = m2.cidclien01 " +
                " join mgw10006 m6 on m8.cidconce01 = m6.cidconce01 " +
                " and m6.ccodigoc01 =  '" + aConcepto + "'" +
                " where cfolio = " + aFolio +
            " and cseriedo01 = '" + aSerie + "'";
            //lsql.Parameters.Add(lparametrodocumentode);
            //lsql.Parameters.Add(lparametrofolio);
            if (aTipo==0)
                lsql.Connection = miconexion.mAbrirConexionOrigen();
            else
                lsql.Connection = miconexion.mAbrirConexionDestino();
            
            
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows )
            {
                lreader.Read ();
                //mLlenarDocto(lreader);

                lRespuesta = true;

            }
            miconexion.mCerrarConexionOrigen();
            lreader.Close();
            return lRespuesta ;


 

            

        }



        public string mGrabarDestinos()
        {
            ClassConexion miconexion = new ClassConexion();
            string lregresa = "";
            miconexion.aRutaExe = aRutaExe;
            miconexion.mAbrirConexionDestino (1); 
            lregresa = mGrabarCompra();
            if (lregresa == "")
            {
                miconexion.mAbrirConexionOrigen(1);
                // Grabar Factura
                mGrabarFactura();
            }

            return lregresa;
        }

        private string mGrabarCompra()
        {
            //barra.Avanzar();
            long lret;
            string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoCompra");
            long lIdDocumento;

            RegProveedor lRegProveedor = new RegProveedor();
            lRegProveedor = mBuscarCliente(GetSettingValueFromAppConfigForDLL("Proveedor").ToString().Trim(), 1, 1);

            fInsertarDocumento();
            
            

            // lret = fSetDatoDocumento("cFecha", DateTime.Today.ToString()); 
            

            lret = fSetDatoDocumento("cCodigoConcepto", lCodigoConcepto);

            lret = fSetDatoDocumento("cSerieDocumento", GetSettingValueFromAppConfigForDLL("SerieCompra").ToString().Trim());
            string lproveedor = GetSettingValueFromAppConfigForDLL("Proveedor");
            lret = fSetDatoDocumento("cCodigoCteProv", lproveedor);

            //buscar el rfc y la razon social de proveedor
            lret = fSetDatoDocumento("cRazonSocial", lRegProveedor.RazonSocial);
            lret = fSetDatoDocumento("cRFC", lRegProveedor.RFC);

            //lret = fSetDatoDocumento("cRazonSocial", ldr["crazonso01"].ToString());
            //lret = fSetDatoDocumento("cRFC", ldr["crfc"].ToString());
            lret = fSetDatoDocumento("cIdMoneda", "1");
            //barra.Avanzar();
            //lret = fSetDatoDocumento("cTipoCambio", z.Cells[21].Value.ToString());
            //lret = fSetDatoDocumento("cReferencia", z.Cells[18].Value.ToString());
            lret = fSetDatoDocumento("cFolio", GetSettingValueFromAppConfigForDLL("FolioFactura").ToString().Trim());
            lret = fSetDatoDocumento("cReferencia", "Por Programa");
            //lret = fSetDatoDocumento("cSerieDocumento", GetSettingValueFromAppConfigForDLL("SerieCompra"));
            DateTime lFecha;
            lFecha = DateTime.Today;
            
            string lfechavenc = "";
            lfechavenc = String.Format("{0:MM/dd/yyyy}", lFecha); ;  // "8 08 008 2008"   year
            lret = fSetDatoDocumento("cFecha", lfechavenc );
            DateTime lFechaVencimiento;
            lFechaVencimiento = DateTime.Today.AddDays(lRegProveedor.DiasCredito);
            lfechavenc = "";
            lfechavenc = String.Format("{0:MM/dd/yyyy}", lFechaVencimiento); ;  // "8 08 008 2008"   year
            lret = fSetDatoDocumento("cFechaVencimiento", lfechavenc);
            lret = fGuardaDocumento();
            //barra.Avanzar();

            if (lret != 0)
            {
                miconexion.mCerrarConexionOrigen(1);
                return "El documento de compra ya existe con el folio y serie de la compra por lo que no se grabara";
            }
            // buscar el id del documento generado
            lIdDocumento = mBuscarIdDocumento(lCodigoConcepto, 1, GetSettingValueFromAppConfigForDLL("SerieCompra").ToString().Trim(), long.Parse(GetSettingValueFromAppConfigForDLL("FolioFactura").ToString().Trim()));
            long lNumeroMov = 100;
            string lregresa = "";
            lret = fInsertarMovimiento();
            productos = "";
            almacenes = "";

            foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
            {
                
                lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
                if (lret != 0)
                {
                    lregresa += "@" + x.cCodigoProducto.Trim();
                    productos += x.cCodigoProducto.Trim();
                }

                lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);
                if (lret != 0)
                {
                    lregresa += "!" + x.cCodigoAlmacen.Trim();
                    almacenes += x.cCodigoAlmacen.Trim();
                }

            }
            if (lregresa == "")
            {
                decimal lprecioconmargen = 0;
                foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
                {
                    //barra.Avanzar();
                    lret = fSetDatoMovimiento("cIdDocumento", lIdDocumento.ToString());
                    lret = fSetDatoMovimiento("cNumeroMovimiento", lNumeroMov.ToString());

                    lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
                    if (lret != 0)
                        lregresa += "#&" + x.cCodigoProducto;

                    lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);


                    lret = fSetDatoMovimiento("cUnidades", x.cUnidades.ToString());

                    //lprecioconmargen = x.cPrecio * (1 + (x.cMargenUtilidad / 100));

                    lprecioconmargen = x.cMargenUtilidad;

                    lret = fSetDatoMovimiento("cPrecio", lprecioconmargen.ToString());
                    //lret = fSetDatoMovimiento("cPorcentajeImpuesto1", z.Cells[17].Value.ToString());
                    //w = decimal.Parse(z.Cells[4].Value.ToString()) * decimal.Parse(z.Cells[6].Value.ToString());
                    lret = fSetDatoMovimiento("cImpuesto1", x.cImpuesto.ToString());

                    lret = fGuardaMovimiento();
                    lNumeroMov += 100;
                    lret = fInsertarMovimiento();

                }
                long lrespuesta = 0;
                if (lret == 0)
                    lrespuesta = fAfectaDocto_Param(lCodigoConcepto, GetSettingValueFromAppConfigForDLL("SerieCompra").ToString().Trim(), double.Parse(_RegDoctoOrigen.cFolio.ToString()), true);
            }
            else
                fBorraDocumento();
            miconexion.mCerrarConexionDestino();
            //barra.Asignar(50);
            return lregresa;
        }

        private string mGrabarFactura()
        {
            //barra.Avanzar();
            //return "";
            long lret;
            string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoFactura").ToString().Trim() ;
            long lIdDocumento;
            RegProveedor lRegProveedor = new RegProveedor ();
            lRegProveedor = mBuscarCliente(GetSettingValueFromAppConfigForDLL("Cliente").ToString ().Trim() , 0, 0);
            

            fInsertarDocumento();
            lret = fSetDatoDocumento("cFecha", DateTime.Today.ToString ()   );
            lret = fSetDatoDocumento("cCodigoConcepto", lCodigoConcepto);
            lret = fSetDatoDocumento("cSerieDocumento", GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim());
            lret = fSetDatoDocumento("cCodigoCteProv", lRegProveedor.Codigo );
            lret = fSetDatoDocumento("cRazonSocial", lRegProveedor.RazonSocial );
            lret = fSetDatoDocumento("cRFC", lRegProveedor.RFC );
            lret = fSetDatoDocumento("cIdMoneda", "1");
            lret = fSetDatoDocumento("cTipoCambio", "1");
            lret = fSetDatoDocumento("cReferencia", "Por Programa");
            lret = fSetDatoDocumento("cFolio", GetSettingValueFromAppConfigForDLL("FolioFactura").ToString().Trim()) ; 
            //lret = fSetDatoDocumento("cSerieDocumento", GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim ()   );
            lret = fGuardaDocumento();
            if (lret != 0)
            {
                miconexion.mCerrarConexionOrigen(1); 
                return "El documento de factura ya existe con el folio y serie mostrados en pantalla";
            }

            // buscar el id del documento generado
            lIdDocumento = mBuscarIdDocumento(lCodigoConcepto, 0, GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim(), long.Parse(GetSettingValueFromAppConfigForDLL("FolioFactura").ToString().Trim()));

            RegDireccion lRegDireccion = new RegDireccion();
            // la direccion del cliente pasarla a la direccion de la factura
            lRegDireccion = mBuscarDireccion(lRegProveedor.Id ,0);
            
            if (!string.IsNullOrEmpty (lRegDireccion.cNombreCalle))
            {
                lret = fInsertaDireccion();
                lret = fSetDatoDireccion("cIdCatalogo", lIdDocumento.ToString ());
                lret = fSetDatoDireccion("cTipoCatalogo", "3");
                lret = fSetDatoDireccion("cTipoDireccion", "0");
                lret = fSetDatoDireccion("cNombreCalle", lRegDireccion.cNombreCalle );
                lret = fSetDatoDireccion("cNumeroExterior", lRegDireccion.cNumeroExterior );
                lret = fSetDatoDireccion("cNumeroInterior", lRegDireccion.cNumeroInterior );
                lret = fSetDatoDireccion("cColonia", lRegDireccion.cColonia  );
                lret = fSetDatoDireccion("cCodigoPostal", lRegDireccion.cCodigoPostal  );
                lret = fSetDatoDireccion("cEstado", lRegDireccion.cEstado );
                lret = fSetDatoDireccion("cPais", lRegDireccion.cPais );
                lret = fSetDatoDireccion("cCiudad", lRegDireccion.cCiudad );
                lret = fGuardaDireccion();
            }

            
            long lNumeroMov = 100;

            foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
            {
                //barra.Avanzar();
                lret = fInsertarMovimiento();
                lret = fSetDatoMovimiento("cIdDocumento", lIdDocumento.ToString());
                lret = fSetDatoMovimiento("cNumeroMovimiento", lNumeroMov.ToString());

                lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto );
                lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen );
                lret = fSetDatoMovimiento("cUnidades", x.cUnidades.ToString () );
                lret = fSetDatoMovimiento("cPrecio", x.cPrecio.ToString () );
                //lret = fSetDatoMovimiento("cPorcentajeImpuesto1", z.Cells[17].Value.ToString());
                //w = decimal.Parse(z.Cells[4].Value.ToString()) * decimal.Parse(z.Cells[6].Value.ToString());
                lret = fSetDatoMovimiento("cImpuesto1", x.cImpuesto .ToString());

                lret = fGuardaMovimiento();
                lNumeroMov += 100;

            }
            long lrespuesta = 0;
            if (lret == 0)
                lrespuesta = fAfectaDocto_Param(lCodigoConcepto, GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim(), double.Parse (_RegDoctoOrigen.cFolio.ToString()), true);
            miconexion.mCerrarConexionOrigen(1);
            //barra.Asignar(100);
            return "";
                    

        }

        //public long mBuscarDocumento(string aConcepto, long aFolio)
        //{
            //_RegDoctoOrigen.cFolio = aFolio;
            //'return mBuscarIdDocumento(aConcepto, 0);
        //}

        private long mBuscarIdDocumento(string aConcepto, int aTipo, string aSerie, long afolio)
        {
            OleDbConnection lconexion= new OleDbConnection ();
            if (aTipo == 0)
                lconexion = miconexion.mAbrirConexionOrigen();
            else
                lconexion = miconexion.mAbrirConexionDestino();

            string lcadena = "select m8.ciddocum01,m2.crazonso01, m2.crfc from mgw10008 m8 " +
            " join mgw10002 m2 on m8.cidclien01 = m2.cidclien01 " +
            " join mgw10006 m6 on m8.cidconce01 = m6.cidconce01 " +
            " where m6.ccodigoc01 = '" + aConcepto + "' and m8.cfolio = " + afolio.ToString() +
            " and cseriedo01 = '" + aSerie + "'";

            OleDbCommand lsql = new OleDbCommand (lcadena ,lconexion );
            OleDbDataReader lreader;
            long lIdDocumento = 0;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lreader.Read();
                lIdDocumento = long.Parse(lreader["ciddocum01"].ToString());
                _rfc = lreader["crfc"].ToString(); 
                _razonsocial = lreader["crazonso01"].ToString();
            }
            lreader.Close();

            return lIdDocumento ;
 
        }
        private RegDireccion  mBuscarDireccion(long  aCliente, int aTipo)
        {
            string sql;
            OleDbConnection lconexion = new OleDbConnection();
            RegDireccion lreg = new RegDireccion ();
            lconexion = miconexion.mAbrirConexionOrigen();
            sql = "select * from mgw10011 where cidcatal01 = " + aCliente + 
                        " and ctipocat01 = 1 and ctipodir01 = " + aTipo ;
            OleDbCommand lsql = new OleDbCommand(sql, lconexion);
            OleDbDataReader lreader;
            //long lIdDocumento = 0;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lreader.Read();
                lreg.cNombreCalle = lreader["cnombrec01"].ToString().Trim();
                lreg.cNumeroExterior = lreader["cnumeroe01"].ToString().Trim();
                lreg.cNumeroInterior = lreader["cnumeroi01"].ToString().Trim();
                lreg.cColonia = lreader["ccolonia"].ToString().Trim();
                lreg.cCodigoPostal = lreader["ccodigop01"].ToString().Trim();
                lreg.cEstado = lreader["cestado"].ToString().Trim();
                lreg.cPais = lreader["cpais"].ToString().Trim();
                lreg.cCiudad = lreader["cciudad"].ToString().Trim();
            }
            lreader.Close();

            return lreg ;

        }

        protected virtual  string GetSettingValueFromAppConfigForDLL(string aNombreSetting)
        {
            string lrutadminpaq = Directory.GetCurrentDirectory();
            if (Directory.GetCurrentDirectory() != aRutaExe)
                Directory.SetCurrentDirectory(aRutaExe);

            string value = "";
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);
            ClientSettingsSection userSettingsSection = (ClientSettingsSection)config.SectionGroups["userSettings"].Sections[_NombreAplicacion + ".Properties.Settings"];
            //SettingElement elemToDelete = null;

            foreach (SettingElement connStr in userSettingsSection.Settings)
            {
                if (connStr.Name == aNombreSetting)
                {
                    value = connStr.Value.ValueXml.InnerText;
                    break;
                }
            }
            if (lrutadminpaq != aRutaExe)
                Directory.SetCurrentDirectory( lrutadminpaq);
            return value;
        }

        public List<RegConcepto> mCargarConceptosComercial(long aIdDocumentoDe, int aTipo)
        {
            List<RegConcepto> _RegFacturas = new List<RegConcepto>();
            SqlConnection lconexion = new SqlConnection();
             
            if (aTipo == 0)
                lconexion = miconexion.mAbrirConexionComercial(false);
            else
                lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {

                //OleDbCommand lsql = new OleDbCommand("select ccodigoc01,cnombrec01 from mgw10006 where ciddocum01 = " + aIdDocumentoDe + " and cescfd = 1 and cnombrec01 = 'CFDI'", lconexion);
                // este es para flexo
                SqlCommand lsql = new SqlCommand("select ccodigoconcepto,cnombreconcepto from admConceptos where ciddocumentode = " + aIdDocumentoDe, lconexion);
                SqlDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _RegFacturas.Clear();
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegConcepto lRegConcepto = new RegConcepto();
                        lRegConcepto.Codigo = lreader[0].ToString();
                        
                        lRegConcepto.Nombre = lreader[1].ToString();
                        //lRegConcepto.Tipocfd = lreader[2].ToString();
                        _RegFacturas.Add(lRegConcepto);
                    }
                }
                lreader.Close();
            }

            return _RegFacturas;



        }

        public List<RegConcepto> mCargarConceptos(long aIdDocumentoDe, int aTipo, int cfdi)
        {
            List<RegConcepto > _RegFacturas = new List<RegConcepto >(); 
            OleDbConnection lconexion = new OleDbConnection();
            if (aTipo == 0)
                lconexion = miconexion.mAbrirConexionOrigen();
            else
                lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {

                //OleDbCommand lsql = new OleDbCommand("select ccodigoc01,cnombrec01 from mgw10006 where ciddocum01 = " + aIdDocumentoDe + " and cescfd = 1 and cnombrec01 = 'CFDI'", lconexion);
                // este es para flexo
                string sqlstring = "select ccodigoc01,cnombrec01,cverfacele from mgw10006 where ciddocum01 = " + aIdDocumentoDe ;
                if (cfdi == 1)
                    sqlstring = "select ccodigoc01,cnombrec01,cverfacele from mgw10006 where ciddocum01 = " + aIdDocumentoDe + " and cescfd = 1";

                OleDbCommand lsql = new OleDbCommand(sqlstring, lconexion);
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _RegFacturas.Clear();
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegConcepto lRegConcepto = new RegConcepto();
                        lRegConcepto.Codigo = lreader[0].ToString();
                        lRegConcepto.Nombre = lreader[1].ToString();
                        lRegConcepto.Tipocfd = lreader[2].ToString();
                        _RegFacturas.Add(lRegConcepto);
                    }
                }
                lreader.Close();
            }
            
            return _RegFacturas;

                  

        }

        public List<RegOrigen> mCargarDocumentos(int aDocumentoDe, int aFolio, string aSerie)
        {
            List<RegOrigen> _RegOrigenes = new List<RegOrigen>();
            OleDbConnection lconexion = new OleDbConnection();
            lconexion = miconexion.mAbrirConexionOrigen();
            if (lconexion != null)
            {

                //OleDbCommand lsql = new OleDbCommand("select ccodigoc01,cnombrec01 from mgw10006 where ciddocum01 = " + aIdDocumentoDe + " and cescfd = 1 and cnombrec01 = 'CFDI'", lconexion);
                // este es para flexo
                OleDbCommand lsql = new OleDbCommand("select m5.cidprodu01, m5.ccodigop01, m5.cnombrep01,m5.cimpuesto2,m5.cimpuesto3, m5.cimporte01, m2.crazonso01, m2.cidclien01, m8.ctotal,m2.ccodigoc01 as cliente, m10.cprecioc01 as precio, m2.crfc as rfc, m10.cunidades as unidades, m10.cneto as TotalMov from mgw10008 m8 " +
" join mgw10010 m10 on m10.ciddocum01 = m8.ciddocum01 " +
" join mgw10005 m5 on m5.cidprodu01 = m10.cidprodu01 " +
" join mgw10002 m2 on m2.cidclien01 = m8.cidclien01 " +
" where m8.cfolio = " + aFolio.ToString() +
" and trim(m8.cseriedo01) = '" + aSerie + "'" + 
" and m8.ciddocum02 = " + aDocumentoDe.ToString(), lconexion);

                //lsql.Parameters.Add("@folio",
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _RegOrigenes.Clear();
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegOrigen lRegOrigen = new RegOrigen();
                        lRegOrigen.CodigoProducto = lreader["ccodigop01"].ToString().Trim();
                        lRegOrigen.NombreProducto = lreader["cnombrep01"].ToString().Trim();
                        lRegOrigen.cidproducto = int.Parse(lreader["cidprodu01"].ToString());
                        lRegOrigen.IEPS = decimal.Parse( lreader["cimpuesto2"].ToString());
                        lRegOrigen.IEPS2 = decimal.Parse(lreader["cimpuesto3"].ToString());
                        lRegOrigen.Descuento = decimal.Parse(lreader["cimporte01"].ToString());

                        lRegOrigen.cTotal = Math.Round(decimal.Parse(lreader["cTotal"].ToString()), 2);

                        

                        lRegOrigen.cIdClien01 = long.Parse(lreader["cidclien01"].ToString());
                        lRegOrigen.RazonSocial = lreader["cRazonSo01"].ToString();
                        lRegOrigen.CodigoCliente = lreader["cliente"].ToString();
                        lRegOrigen.Precio = Math.Round( decimal.Parse(lreader["precio"].ToString()),2);
                        lRegOrigen.Precio2 = Math.Round( decimal.Parse(lreader["precio"].ToString()),2);
                        lRegOrigen.TotalMov = Math.Round( decimal.Parse(lreader["TotalMov"].ToString()),2);
                        lRegOrigen.Cantidad = Math.Round( decimal.Parse(lreader["Unidades"].ToString()),2);
                        //lRegOrigen.TotalMov2 = decimal.Parse(lreader["TotalMov"].ToString());

                        if (lRegOrigen.Descuento == 0)
                        {
                            // precio facturado - precio capturado * unidades facturadas 

                            lRegOrigen.DescuentoAplicar = 0;

                        }
                        else
                        {
                            // precio facturado * unidades facturadas * descuento
                            lRegOrigen.DescuentoAplicar = Math.Round( lRegOrigen.Precio * lRegOrigen.Cantidad * (lRegOrigen.Descuento / 100));
                        }
                        

                        lRegOrigen.CodigoCliente = lreader["cliente"].ToString();
                        lRegOrigen.RFC = lreader["rfc"].ToString();

                        _RegOrigenes.Add(lRegOrigen);
                    }
                }
                lreader.Close();
            }

            return _RegOrigenes;



        }


        public List<RegProveedor> mCargarClientes()
        {
            List<RegProveedor> _RegProveedores = new List<RegProveedor>();
            OleDbConnection lconexion = new OleDbConnection();
            
            lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {

                //OleDbCommand lsql = new OleDbCommand("select ccodigoc01,cnombrec01 from mgw10006 where ciddocum01 = " + aIdDocumentoDe + " and cescfd = 1 and cnombrec01 = 'CFDI'", lconexion);
                // este es para flexo

                //string lstring =  "select ccodigoc01,rtrim(crazonso01)+' ('+rtrim(ccodigoc01) + ')'"  +
                //" from mgw10002 where ctipocli01 < 2 and cidclien01 > 0";

                OleDbCommand lsql = new OleDbCommand("select ccodigoc01,rtrim(ccodigoc01)+' ('+rtrim(crazonso01) + ')'" +
                " from mgw10002 where ctipocli01 < 2 and cidclien01 > 0 order by ccodigoc01 ", lconexion);
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _RegProveedores.Clear();
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegProveedor lRegCliente = new RegProveedor();
                        lRegCliente.Codigo = lreader[0].ToString();
                        lRegCliente.RazonSocial = lreader[1].ToString();
                        //lRegCliente.Tipocfd = lreader[2].ToString();
                        _RegProveedores .Add(lRegCliente);
                    }
                }
                lreader.Close();
            }

            return _RegProveedores;



        }

        public RegProveedor mBuscarCliente(string aCliente, int aTipo, int aTipoCliente )
        {
            OleDbConnection lconexion = new OleDbConnection();
            RegProveedor lReg = new RegProveedor ();
            string lcadena;
            if (aTipo == 0)
                lconexion = miconexion.mAbrirConexionOrigen();
            else
                lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {
                if (aTipoCliente == 0)
                    lcadena = "select ccodigoc01,crazonso01, cidclien01, crfc, cdiascre01 from mgw10002 where ctipocli01 < 2 and ccodigoc01 = '" + aCliente + "'";
                else
                    lcadena = "select ccodigoc01,crazonso01, cidclien01, crfc, cdiascre02 from mgw10002 where ctipocli01 > 1 and ccodigoc01 = '" + aCliente + "'";


                OleDbCommand lsql = new OleDbCommand(lcadena, lconexion);
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    lReg.Codigo = lreader[0].ToString();
                    lReg.RazonSocial = lreader[1].ToString();
                    lReg.Id = long.Parse(lreader[2].ToString());
                    lReg.RFC = lreader[3].ToString();
                    lReg.DiasCredito = int.Parse ( lreader[4].ToString());
                }
                lreader.Close();
            }
            return lReg;
                
                
        }

        public List<RegEmpresas> mCargarEmpresasAccess(out string amensaje)
        {

            OleDbConnection lconexion = new OleDbConnection();

            lconexion = miconexion.mAbrirConexionAccess(out amensaje);
            
            List<RegEmpresas> _RegEmpresas = new List<RegEmpresas >();
            //amensaje = lconexion.ConnectionString;

            if (amensaje == "")
            {
                //lconexion = miconexion.mAbrirConexionDestino();
                try
                {

                    OleDbCommand lsql = new OleDbCommand("SELECT distinct(Empresa) from tbl_puntosdeventa order by Empresa ", lconexion);
                    OleDbDataReader lreader;
                    //long lIdDocumento = 0;
                    lreader = lsql.ExecuteReader();
                    _RegEmpresas.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegEmpresas lRegEmpresas= new RegEmpresas();
                            lRegEmpresas.cEmpresa = lreader[0].ToString();
                            
                            _RegEmpresas.Add(lRegEmpresas);
                        }
                    }
                    lreader.Close();

                }
                catch (Exception eeeee)
                {
                    amensaje = eeeee.Message;
                }

            }


            
            return _RegEmpresas;




        }



        public List<RegPuntodeVenta> mCargarPuntoVenta(string aEmpresa, out string amensaje)
        {

            OleDbConnection lconexion = new OleDbConnection();

            lconexion = miconexion.mAbrirConexionAccess (out amensaje);

            List<RegPuntodeVenta> _RegPUntosVenta = new List<RegPuntodeVenta>();
            //amensaje = lconexion.ConnectionString;

            if (amensaje == "")
            {
                //lconexion = miconexion.mAbrirConexionDestino();
                try
                {

                    OleDbCommand lsql = new OleDbCommand("SELECT Nombre from tbl_puntosdeventa  where Empresa ='" + aEmpresa + "'", lconexion);
                    OleDbDataReader lreader;
                    //long lIdDocumento = 0;
                    lreader = lsql.ExecuteReader();
                    _RegPUntosVenta.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegPuntodeVenta lRePuntodeVenta = new RegPuntodeVenta();
                            //lRePuntodeVenta.cEmpresa = lreader[0].ToString();
                            lRePuntodeVenta.cNombre  = lreader[0].ToString();
                            _RegPUntosVenta.Add(lRePuntodeVenta );
                        }
                    }
                    lreader.Close();

                }
                catch (Exception eeeee)
                {
                    amensaje = eeeee.Message;
                }

            }



            return _RegPUntosVenta ;




        }


        public List<RegEmpresa> mCargarEmpresas(out string amensaje)
        {
            
            OleDbConnection lconexion = new OleDbConnection();
            
            lconexion = miconexion.mAbrirRutaGlobal (out amensaje);

            List<RegEmpresa> _RegEmpresas = new List<RegEmpresa>();
            //amensaje = lconexion.ConnectionString;
            
            if (amensaje == "")
            {
                //lconexion = miconexion.mAbrirConexionDestino();
                try
                {
                    
                    OleDbCommand lsql = new OleDbCommand("select cnombree01,crutadatos from mgw00001 where cidempresa > 1 ", lconexion);
                    OleDbDataReader lreader;
                    //long lIdDocumento = 0;
                    lreader = lsql.ExecuteReader();
                    _RegEmpresas.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegEmpresa lRegEmpresa = new RegEmpresa();
                            lRegEmpresa.Nombre = lreader[0].ToString();
                            lRegEmpresa.Ruta = lreader[1].ToString();
                            _RegEmpresas.Add(lRegEmpresa);
                        }
                    }
                    lreader.Close();

                }
                catch (Exception eeeee)
                {
                    amensaje = eeeee.Message;
                }
                
            }

                
            
            return _RegEmpresas;




        }

        public virtual bool mValidarConexionIntell(string aServidor, string aBd, string ausu, string apwd)
        {
            string Cadenaconexion = "data source =" + aServidor+ ";initial catalog =" + aBd   + ";user id = " + ausu  + "; password = " +  apwd  + ";";
            
            _con.ConnectionString = Cadenaconexion;
            try
            {
                _con.Open();
                // si se conecto grabar los datos en el cnf
                _con.Close();
                return true;
            }
            catch (Exception ee)
            {
                return false;
            }
        }

        public virtual bool mValidarConexionIntell(string aRuta)
        {
            //string Cadenaconexion = "data source =" + aServidor + ";initial catalog =" + aBd + ";user id = " + ausu + "; password = " + apwd + ";";

            ClassConexion x = new ClassConexion();

            string lmsg = "'";

            //_con = miconexion.mAbrirConexionAccess (out lmsg);
            return true;

            /*
            _con.ConnectionString = Cadenaconexion;
            try
            {
                _con.Open();
                // si se conecto grabar los datos en el cnf
                _con.Close();
                return true;
            }
            catch (Exception ee)
            {
                return false;
            }
             */
        }

        

        public virtual string mBuscarDoctos(long aFolio, long afoliofinal, int aTipo, Boolean aRevisar)
        {

            

            string lrespuesta = "";
            _RegDoctos.Clear();
            lrespuesta = mBuscarDoctoArchivo(aRevisar);
            return lrespuesta;
        /*

            for (long i = aFolio; i <= afoliofinal; i++)
            {
                RegDocto lDocto = new RegDocto();
                _RegDoctoOrigen = null;
                _RegDoctoOrigen = new RegDocto();
                lrespuesta = mBuscarDoctoAccess( aRevisar);
                if (lrespuesta == string.Empty)
                {
                    _RegDoctoOrigen.sMensaje = "";
                    _RegDoctoOrigen.cFolio = i;
                }
                else
                {
                    _RegDoctoOrigen.sMensaje = lrespuesta;
                    _RegDoctoOrigen.cFolio = i;
                }
                lDocto = _RegDoctoOrigen;
                _RegDoctos.Add(lDocto );
            }
            return lrespuesta;*/
        }

        
        public string mBuscarDocto (string aFolio, int aTipo, Boolean aRevisar)
        {
            OleDbCommand  lcmd = new OleDbCommand ();
            OleDbDataReader lreader;
            string lRespuesta= "";
            if (aTipo == 0)
                return lRespuesta;
            _con.Open();
            lcmd.Connection = _con;
            if (aTipo == 1 || aTipo == 2)
            {
                lcmd.CommandText = "select v.cliente as cliente, FechaEmision as cfecha, " +
                " ID as cfolio, c.Direccion as cnombrec01, c.DireccionNumero as cnumeroe01, c.DireccionNumeroInt as cnumeroi01, c.Colonia as ccolonia, c.Poblacion as cciudad, c.Estado as cestado, c.Pais as cpais " +
                " , c.RFC as crfc, c.Nombre as crazonso01, c.CodigoPostal as ccodigop01, v.moneda, v.tipocambio ";

                if (aTipo == 1)

                    lcmd.CommandText +=  ", case " +
                                        " when  v.condicion = '' then '0' " +
                                        " when  v.condicion = 'Contado' then '0'" + 
                                        " when  isnull(v.condicion,0) = '0' then '0'" + 
                                        " else left(v.condicion, isnull(charindex(' DIAS CREDITO',v.condicion,1),0)) " + 
                                        " end as condpago, v.agente " ;

                lcmd.CommandText += " from venta v join Cte c " +
                " on v.Cliente = c.Cliente " +
                "where MovID = '" + aFolio + "'" + 
                " and v.Estatus <> 'CANCELADO'";
            }
            if (aTipo == 3 || aTipo == 4)
            {
                lcmd.CommandText = "select " +
                " v.cliente as cliente, FechaEmision as cfecha, " +
                " v.MovID as cfolio, c.Direccion as cnombrec01, c.DireccionNumero as cnumeroe01, c.DireccionNumeroInt as cnumeroi01, c.Colonia as ccolonia, c.Poblacion as cciudad, c.Estado as cestado, c.Pais as cpais " +
                " , c.RFC as crfc, c.Nombre as crazonso01, c.CodigoPostal as ccodigop01, v.moneda, v.tipocambio, v.importe, v.impuestos " +
                " from cxc v  join Cte c " +
                " on v.Cliente = c.Cliente " +
                " where v.Estatus  = 'CONCLUIDO' " +
                " and MovID = '" + aFolio + "'";
            }


            switch (aTipo)
            {
                case 1:
                    lcmd.CommandText += " and (Mov = 'Factura' or Mov = 'Factura Global')";
                    break;
                case 2:
                    lcmd.CommandText += " and Mov = 'Devolucion Venta'";
                    break;
                case 3:
                    lcmd.CommandText += " and Mov = 'Nota Cargo'";
                    break;
                case 4:
                    lcmd.CommandText += " and Mov = 'Nota Credito'";
                    lcmd.CommandText += " and origen <> 'Devolucion Venta'";
                    break;

            }
                    

            
            lreader = lcmd.ExecuteReader();
            if (lreader.HasRows)
            {
                if (aRevisar == true)
                {
                    if (mBuscarGeneradoADM(aFolio, aTipo) == true)
                    {
                        _con.Close();
                        return "Documento ya existe en Adminpaq";
                    }
                }
                lreader.Read();
                lRespuesta = mLlenarDocto(lreader,aTipo, aFolio,"Mercado"  );
//                if (lRespuesta != string.Empty)
  //                  lRespuesta = "";
            }

             else
                {
                    //lreader.Read();
                    //mLlenarDocto(lreader);
                    lRespuesta = "Documento No Existe";
                }

            lreader.Close();
            
            _con.Close();
            return lRespuesta;
        }

        protected virtual string mConsultaEncabezado(int aTipo, string aFolio)
        {

            string aEmpresa = GetSettingValueFromAppConfigForDLL("Empresa");
            
            string aNombre = GetSettingValueFromAppConfigForDLL("Nombre");
            string aFecha = GetSettingValueFromAppConfigForDLL("Fecha");
            


            string lregresa = "";
            
            /*
                    lregresa = " select top 1 isnull(CustCode,'') as cliente, BillDate as cfecha, " +
                    " isnull(BillNum,0) as cfolio, isnull(billaddrname,'') as cnombrec01, isnull(billaddress2,'') as cnumeroe01, '' as cnumeroi01, " +
                    " isnull(billaddress3,'') as ccolonia, isnull(billtown,'') as cciudad, isnull(billcounty,'') as cestado, isnull(billcountry,'') as cpais " +
                    " , isnull(VatRegNo,'') as crfc, isnull(CustName,'') as crazonso01, isnull(BillPostCode,0) as ccodigop01, " +
                    " 'moneda' = case when currCode = 'MN' then 'Pesos' else '0' end,  " +
                    " '1' as tipocambio, 0 as condpago, '(Ninguno)' as agente, isnull(OCCLIENTE,'') as creferen01 " +
                    " from facturacione " +
                    " where billnum = '" + aFolio + "'";
            */
            // "SELECT C.id_cliente as cliente, o.id_punto_de_venta as cfecha, o.fechamov, o.puntodeventa, p.nombre, o.cantidad, o.importe, pr.nombreproducto, o.id_productos, c.nombrefiscal, c.direccion, c.colonia, m.nombre, e.nombreestados, c.codigopostal " +
            // , c.nombrefiscal
            DateTime lfecha = DateTime.Parse(aFecha);
            string sfecha = lfecha.Month.ToString().Trim().PadLeft(2, '0') + "/" + lfecha.Day.ToString().Trim().PadLeft(2, '0') + "/" + lfecha.Year;
            aFecha = sfecha;
            lregresa = "SELECT C.id_cliente as cliente, o.fechamov as cfecha, 1 as cfolio, c.direccion as cnombrec01, c.celular as cnumeroe01, '' as cnumeroi01, " + 
                    " c.colonia as ccolonia, m.nombre as cciudad, e.nombreestados as cestado, 'Mexico' as cpais, " +
                    " c.fax as crfc, c.nombrefiscal as crazonso01,   c.codigopostal as ccodigop01, 'Pesos'  as moneda," +
                    " '1' as tipocambio, 0 as condpago, '(Ninguno)' as agente, '' as creferen01 " +
                   " FROM ((((tbl_operaciones AS o INNER JOIN tbl_puntosdeventa AS p ON o.id_punto_de_venta =p.id_pv02) INNER JOIN tbl_productos01 AS pr ON o.id_productos = pr.id_productos01) INNER JOIN tbl_clientes01 AS C ON p.empresa = C.nombrefiscal) INNER JOIN tbl_municipios AS m ON c.municipio = m.id_municipios) INNER JOIN tbl_estados AS e ON e.id_estados = c.estado " +
                   " WHERE o.id_punto_de_venta <> '' " +
                   " and o.id_punto_de_venta = p.id_pv02  " +
                   " and o.fechamov = #" + aFecha + "#" +
                   " and  " +
                   " p.Empresa ='" + aEmpresa + "'" +
                   " and p.nombre = '" + aNombre + "'" +
                   "ORDER BY o.fechamov DESC ";

            lregresa = "SELECT o.fechamov as cfecha, referencia_documento as cfolio, p.nombre as cReferen01, " +
                        " o.id_productos as ccodigop01, o.importe as cprecioc01, pr.nombreproducto as cnombrep01, o.cantidad as cunidades, o.observaciones01 as cobserva01 " +
                       "from  " + 
                        " (tbl_operaciones as o " + 
                        " INNER JOIN tbl_puntosdeventa AS p ON o.id_punto_de_venta =p.id_pv02) " + 
                        " INNER JOIN tbl_productos01 AS pr ON o.id_productos = pr.id_productos01 " + 
                        " where " + 
                        " o.id_punto_de_venta <> ''   " + 
                        " and  p.Empresa ='" + aEmpresa + "'" +
                        " and o.fechamov = #" + aFecha + "# " + 
                        " and referencia_documento <> '' " +
                        " ORDER BY referencia_documento, o.fechamov DESC ";
            
            return lregresa;

        }

 private string mProcesaItem(ref int aInicio,string sLine)
 {
     int lfin = sLine.IndexOf("|", aInicio) - aInicio;
     string lRegresa = sLine.Substring(aInicio, lfin);
     aInicio = sLine.IndexOf("|", aInicio) + 1;
     return lRegresa;
 }

 public string mBuscarDoctoArchivo(Boolean aRevisar)
 {
     string lrespuesta = ""; 
     string mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
     StringBuilder sb = new StringBuilder();
     List <RegDocto> misdoctos = new List<RegDocto>();
     _RegDoctos.Clear();
     string lrutacarpeta = @GetSettingValueFromAppConfigForDLL("RutaCarpeta");
     //lrutacarpeta = @lrutacarpeta;
     foreach (string txtName in Directory.GetFiles(lrutacarpeta , "*.txt"))
     {
         StreamReader objReader = new StreamReader(txtName);
         string sLine = "";
         
         ArrayList arrText = new ArrayList();
         RegDocto midocto = new RegDocto();
         int linicio = 0;
         while (sLine != null)
         {
             sLine = objReader.ReadLine();
             if (sLine != null)
             {
                 if (sLine != "")
                 {
                     if (sLine.Substring(0, 1) == "S")
                     {
                         linicio = 2;
                         string x  = mProcesaItem(ref linicio, sLine);
                         x = mProcesaItem(ref linicio, sLine);
                         x = mProcesaItem(ref linicio, sLine);
                         midocto.cNeto = double.Parse(x);
                         //midocto.cTextoExtra1 = "";
                         _RegDoctos.Add(midocto);
                         midocto = new RegDocto();
                     }
                     if (sLine.Substring(0, 2) == "H1")
                     {
                         //buscar contado
                         midocto.cContado = 0;
                         if (sLine.IndexOf("CONTADO") != -1)
                             midocto.cContado = 1;
                         linicio = 3;
                         midocto.cNombreArchivo = txtName;
                         midocto.cSerie = mProcesaItem(ref linicio, sLine);
                         midocto.cFolio = int.Parse(mProcesaItem(ref linicio, sLine));
                         midocto.cTextoExtra1 = mProcesaItem(ref linicio, sLine);
                         midocto.cMoneda = "Pesos";
                         midocto.cTipoCambio = 1;

                         // fecha
                         string y = mProcesaItem(ref linicio, sLine);
                         // 03/10/12
                         y = y.Substring(0, 6) + "20" + y.Substring(8, 2);
                         DateTime dt2 = DateTime.ParseExact(y, "dd/MM/yyyy", null);
                         midocto.cFecha = dt2;

                         //midocto.cTextoExtra1 = mProcesaItem(ref linicio, sLine);
                         string tempo;
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine); //cliente
                         midocto.cCodigoCliente = tempo;
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine); //moneda
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         //midocto.cTextoExtra1 = tempo; // para movimiento

                         midocto.cReferencia = tempo; // para movimiento

                         tempo = mProcesaItem(ref linicio, sLine);//carga y placas 
                         //midocto.cReferencia = tempo.Substring(12, tempo.IndexOf("PLACAS") - 12);

                         midocto.cTextoExtra1 = tempo.Substring(12, tempo.IndexOf("PLACAS") - 12);
                         tempo = tempo.Substring(tempo.IndexOf("PLACAS:") + 8);


                         midocto.cTextoExtra1 = tempo.Substring(0, tempo.IndexOf("RUTA"));
                         midocto.cTextoExtra2 = tempo.Substring(0, tempo.IndexOf("RUTA"));

                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine); 
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         midocto.cAgente = mProcesaItem(ref linicio, sLine); //agente
                         midocto.cCodigoConcepto = mProcesaItem(ref linicio, sLine); //concepto

                         midocto.cCodigoConcepto = midocto.cCodigoConcepto.Trim();

                         // fecha vencimiento
                         string yy = mProcesaItem(ref linicio, sLine);
                         // 03/10/12
                         yy = yy.Substring(0, 6) + "20" + y.Substring(8, 2);
                         DateTime dt3 = DateTime.ParseExact(yy, "dd/MM/yyyy", null);
                         midocto.cFechaVencimiento = dt3;




                     }
                     if (sLine.Substring(0, 1) == "D")
                     {
                         RegMovto movto = new RegMovto();
                         linicio = 2;
                         movto.cNombreProducto = mProcesaItem(ref linicio, sLine);

                         movto.cUnidades = decimal.Parse(mProcesaItem(ref linicio, sLine));

                         //movto.cUnidades = decimal.Parse(mProcesaItem(ref linicio, sLine));
                         //MessageBox.Show(movto.cUnidades.ToString());
                         movto.cUnidad = mProcesaItem(ref linicio, sLine);

                         movto.cPrecio = decimal.Parse(mProcesaItem(ref linicio, sLine));

                         movto.cneto = decimal.Parse(mProcesaItem(ref linicio, sLine));

                         //movto.cPrecio = decimal.Parse(mProcesaItem(ref linicio, sLine));

                         movto.cPorcent01 = decimal.Parse(mProcesaItem(ref linicio, sLine));

                         movto.cImpuesto = decimal.Parse(mProcesaItem(ref linicio, sLine));

                         movto.cCodigoProducto = mProcesaItem(ref linicio, sLine);
                         movto.cCodigoAlmacen = "1";
                         movto.ctextoextra3 = midocto.cTextoExtra1;
                         
                         //MessageBox.Show(movto.cImpuesto.ToString());


                         midocto._RegMovtos.Add(movto);
                     }
                     if (sLine.Substring(0, 2) == "H2")
                     {
                         linicio = 3;
                         string tempo;
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine); //cliente
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine); //moneda
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         midocto.cTextoExtra3 = mProcesaItem(ref linicio, sLine);
                     }
                     if (sLine.Substring(0, 2) == "H4")
                     {
                         linicio = 3;
                         midocto.cRazonSocial = mProcesaItem(ref linicio, sLine);
                         midocto.cRFC = mProcesaItem(ref linicio, sLine);
                         //midocto.cCodigoCliente = midocto.cRFC;
                         // linicio = lfin + 1;
                         //lfin = sLine.IndexOf("|", linicio) - linicio;
                         midocto._RegDireccion.cNombreCalle = mProcesaItem(ref linicio, sLine);
                         midocto._RegDireccion.cNumeroExterior = mProcesaItem(ref linicio, sLine); ;
                         midocto._RegDireccion.cNumeroInterior = mProcesaItem(ref linicio, sLine); ;
                         midocto._RegDireccion.cColonia = mProcesaItem(ref linicio, sLine); ;
                         midocto._RegDireccion.cCiudad = mProcesaItem(ref linicio, sLine); ;
                         midocto._RegDireccion.cEstado = mProcesaItem(ref linicio, sLine); ;
                         midocto._RegDireccion.cPais = mProcesaItem(ref linicio, sLine); ;
                         midocto._RegDireccion.cCodigoPostal = mProcesaItem(ref linicio, sLine); ;
                         midocto.sMensaje = "";
                     }
                 }
             }
         }

         objReader.Close();    
     }
     
     return lrespuesta;
 }

        public  string mBuscarDoctoAccess(Boolean aRevisar)
        {
            OleDbCommand lcmd = new OleDbCommand();
            OleDbDataReader lreader;
            
             
            string lRespuesta = "";
            if (_con.State != 0)
                _con.Close();

            _con.Open();
            lcmd.Connection = _con;

            lcmd.CommandText = mConsultaEncabezado(0, "1");

            //OleDbDataAdapter lda = new OleDbDataAdapter(lcmd );
            //System.Data.DataSet xxx = new System.Data.DataSet ();
            //lda.Fill(xxx);

            

            

            try
            {
                lreader = null ;
                //lreader.Close();
                lreader = lcmd.ExecuteReader();
            }
            catch (Exception e)
            {
                lRespuesta = e.Message;
                _con.Close();
                return lRespuesta;
            }
            if (lreader.HasRows)
            {
                if (aRevisar == true)
                {

                    //if (mBuscarADM(aFolio, aTipo) == true)
                    //{
                    //    _con.Close();
                    //    return "Documento Ya existe en Adminpaq"; // documento ya existe
                    //}
                }
                //lreader.Read();
                lRespuesta = mLlenarDoctos(lreader);
            }

            else
            {
                lRespuesta = "Documento no Encontrado"; // documento no encontrado
            }

            _con.Close();
            return lRespuesta;
        }

        private Boolean mBuscarGeneradoADM(string aFolio, int aTipo )
        {
            OleDbConnection lconexion = new OleDbConnection();
            string amensaje="";
            lconexion = miconexion.mAbrirRutaGlobal(out amensaje);
            bool lrespuesta = false;

            //lconexion = miconexion.mAbrirConexionDestino();

            OleDbCommand lsql = new OleDbCommand("select * from interfaz where folioi = '" + aFolio + "' and tipodoc = " + aTipo, lconexion);
            OleDbDataReader lreader;
            //long lIdDocumento = 0;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lrespuesta = true;
            }
            lreader.Close();

            return lrespuesta ;
        }


        protected Boolean mBuscarADM(string aFolio, int aTipo)
        {
            bool lrespuesta = false;
            string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();

            
            miconexion.mAbrirConexionDestino();
            string lcadena = "select cfolio from mgw10008 m8 join mgw10006 m6 on m6.cidconce01 = m8.cidconce01 where m8.cfolio = " + aFolio + " and m6.ccodigoc01 = '" + lCodigoConcepto + "'";
            OleDbCommand lsql = new OleDbCommand(lcadena, miconexion._conexion);
            OleDbDataReader  lreader;
            //long lIdDocumento = 0;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lrespuesta = true;
            }
            lreader.Close();
            miconexion.mCerrarConexionDestino ();
            return lrespuesta;
        }
        public string mGrabarAdm1()
        {

            
                miconexion.mAbrirConexionDestino(1);
                bool lentre = true;
            
            //miconexion.mCerrarConexionDestino();
            miconexion.mCerrarConexionOrigen(1);
            _controlfp(0x9001F, 0xFFFFF); 
            // barra.Asignar(100);
            return "";
        }

        public List<string> mGrabarAdms(int opcion, int tipo)
        {
            string lrespuesta = "";
            string lcadena =  "";

            //List<string> lvar = new List<string>();

            lvar.Clear();
            int lcuantos = _RegDoctos.Count;
            int lindice = 1;

            if (_RegDoctos.Count == 0 )
            {
                lvar.Add("No existe documentos con los filtros seleccionados");
                return lvar; 
            }


            foreach (RegDocto _reg in _RegDoctos)
            {
                _RegDoctoOrigen = null;
                _RegDoctoOrigen = new RegDocto();
                _RegDoctoOrigen = _reg;
                string lCodigoConcepto ;
                //if (opcion != 5)
                //    lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
                //else
                    lCodigoConcepto = _reg.cCodigoConcepto ;

                //lrespuesta = _RegDoctoOrigen.sMensaje;
                //if (_RegDoctoOrigen.sMensaje == string.Empty)
                //{
                  lrespuesta = mGrabarAdm(_reg.cFolio.ToString(), _RegDoctoOrigen.cFolio, opcion, tipo);
                
                //}
                //mActualizarBarra((double)lindice / lcuantos);
                //lporcentaje = 0.0D;
                //lporcentaje = (double)lindice / lcuantos;
                //Notificar();
                Notificar((double)(lindice*100) / lcuantos);
                


                lindice++;
                if (lrespuesta != "")
                {
                    switch (opcion)
                    {
                        case 1:
                            lcadena = "La factura";
                            break;
                        case 2:
                            lcadena = "El pedido";
                            break;
                        case 3:
                            lcadena = "La nota de credito";
                            break;
                        case 4:
                            lcadena = "La nota de cargo";
                            break;
                    }
                    lcadena += " con folio " + _reg.cFolio.ToString() + " presento el siguiente problema " + lrespuesta + Convert.ToChar(13);

                    lvar.Add(lcadena);
                }
                else
                { 
                    //copiar el archivo
                    string lrutaorigen = GetSettingValueFromAppConfigForDLL("RutaCarpeta");
                    lrutaorigen += "\\"  + _RegDoctoOrigen.cNombreArchivo;
                    string lrutadestino = GetSettingValueFromAppConfigForDLL("RutaCarpetaBackup") ;
                    string larchivo = System.IO.Path.GetFileName(_RegDoctoOrigen.cNombreArchivo);
                    lrutadestino = lrutadestino + "\\" + larchivo;


                   // File.Move(_RegDoctoOrigen.cNombreArchivo, lrutadestino);
                    
                }

                    

            }

            return lvar;
        }
          
        protected virtual void mActualizarBarra(double valor)
        {
            return;
            //lporcentaje = 0.0D;
                //lporcentaje = (double)lindice / lcuantos;
                //Notificar(lporcentaje);
        }

        private void mRegresarPrincipales(string lCodigoConcepto,ref long lidconce, ref long tipocfd, ref string cserie, ref int cnaturaleza)
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            lidconce = 0;
            tipocfd = 0;
            lsql.CommandText = "select cidconce01, ciddocum01,cseriepo01,cverfacele,cnatural01 from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                lreader.Read();
                cserie = lreader["cseriepo01"].ToString();
                lidconce = long.Parse(lreader["cidconce01"].ToString());
                tipocfd = long.Parse(lreader["cverfacele"].ToString());
                cnaturaleza = int.Parse(lreader["cnatural01"].ToString());
            }
            else
                cserie = "";
            lreader.Close();
        }
        private string mValidarExisteDoc(long lidconce, string cserie, double afolionuevo)
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            lsql.CommandText = "select count(*) as cuantos from mgw10008 where cidconce01 = " + lidconce + " and cseriedo01 = '" + cserie + "' and cfolio = " + afolionuevo.ToString().Trim();
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                lreader.Read();
                long cuantos = 0;
                cuantos = long.Parse(lreader["cuantos"].ToString());
                lreader.Close();
                if (cuantos > 0)
                {
                    _controlfp(0x9001F, 0xFFFFF);
                    //miconexion.mCerrarConexionOrigen(1);

                    return "Documento ya existe en ADMINPAQ";
                }
            }
            lreader.Close();
            return "";
        }

        private string mGrabarEncabezado(double aFolio, string lCodigoConcepto, string aImpreso)
        {
            long lret,lidconce=0,tipocfd=0;
            string cserie="";
            
            int naturaleza=0;
            mRegresarPrincipales(lCodigoConcepto,ref lidconce, ref tipocfd, ref cserie, ref naturaleza);
            string lresp = mValidarExisteDoc(lidconce,cserie,aFolio);
            if (lresp != "")
                return lresp;

            fInsertarDocumento();
            lret = fSetDatoDocumento("cCodigoConcepto", lCodigoConcepto);
            lret = fSetDatoDocumento("cSerieDocumento", _RegDoctoOrigen.cSerie);
            lret = fSetDatoDocumento("cCodigoCteProv", _RegDoctoOrigen.cCodigoCliente);
            //lresp = "";
            //if (lret!=0)
            //{
            //    lresp=mGrabarCliente();
            //}
            //if (lresp != "")
            //    return lresp;

            //lret = fSetDatoDocumento("cRazonSocial", _RegDoctoOrigen.cRazonSocial);
            //lret = fSetDatoDocumento("cRFC", _RegDoctoOrigen.cRFC);
            lret = fSetDatoDocumento("cIdMoneda", "1");
            //lret = fSetDatoDocumento("cReferencia", "Por Programa");
            lret = fSetDatoDocumento("cFolio", aFolio.ToString().Trim());
            lret = fSetDatoDocumento("cTipoCambio", "1");

            try
            {
                //lret = fSetDatoDocumento("cReferencia", _RegDoctoOrigen.cFolio.ToString());
                //lret = fSetDatoDocumento("cTextoExtra1", _RegDoctoOrigen.cReferencia);
                lret = fSetDatoDocumento("cReferencia", _RegDoctoOrigen.cReferencia);
                //lret = fSetDatoDocumento("cObservaciones", _RegDoctoOrigen.cTextoExtra3);
                //lret = fSetDatoDocumento("cTextoExtra2", _RegDoctoOrigen.cTextoExtra2);
                //lret = fSetDatoDocumento("cTextoExtra3", _RegDoctoOrigen.cTextoExtra3);
                //lret = fSetDatoDocumento("cImpreso", aImpreso);
            }
            catch (Exception ee)
            {
            }

         /*
            DateTime lFechaVencimiento;
            lFechaVencimiento = _RegDoctoOrigen.cFecha.AddDays(int.Parse("0"));
            lFechaVencimiento = _RegDoctoOrigen.cFechaVencimiento;
            //lFechaVencimiento = DateTime.Today.AddDays(int.Parse(_RegDoctoOrigen.cCond) );

            string lfechavenc = "";
            lfechavenc = String.Format("{0:MM/dd/yyyy}", lFechaVencimiento); ;  // "8 08 008 2008"   year
           */ 


            

            string lfechadocto = "";
            lfechadocto = _RegDoctoOrigen.cFecha.ToString();
            DateTime lFechaDocto;
            lFechaDocto = _RegDoctoOrigen.cFecha;

            lfechadocto = "";


            lfechadocto = String.Format("{0:MM/dd/yyyy}", lFechaDocto); ;  // "8 08 008 2008"   year


            lret = fSetDatoDocumento("cFecha", lfechadocto);

            
            
            lret = fGuardaDocumento();
            
            if (lret != 0)
            {
                //fError(lret, serror, 255);
                _controlfp(0x9001F, 0xFFFFF);
                //miconexion.mCerrarConexionOrigen(1);
                return lret.ToString()+ " Documento ya Existe";
                

            }
            return "";



        }

        private string  mGrabarCliente()
        { 
            long lret = 0;
            fInsertaCteProv();
                lret = fSetDatoCteProv("CCODIGOCLIENTE", _RegDoctoOrigen.cCodigoCliente);
                lret = fSetDatoCteProv("cRazonSocial", _RegDoctoOrigen.cRazonSocial);
                if (lret != 0)
                {
                    _controlfp(0x9001F, 0xFFFFF);
                    miconexion.mCerrarConexionOrigen(1);
                    return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cRazonSocial;
                }
                lret = fSetDatoCteProv("cRFC", _RegDoctoOrigen.cRFC);
                if (lret != 0)
                {
                    _controlfp(0x9001F, 0xFFFFF);
                    miconexion.mCerrarConexionOrigen(1);
                    return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cRFC;
                }
                lret = fSetDatoCteProv("CLISTAPRECIOCLIENTE", "1");
                lret = fSetDatoCteProv("CIDMONEDA", "1");

                string lfecha = _RegDoctoOrigen.cFecha.ToString();
                DateTime ldate = DateTime.Parse(lfecha);
                lfecha = ldate.ToString("MM/dd/yyyy");
                lret = fSetDatoCteProv("CFECHAALTA", lfecha);
                if (lret != 0)
                {
                    _controlfp(0x9001F, 0xFFFFF);
                    miconexion.mCerrarConexionOrigen(1);
                    return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cFecha.ToString();
                }
                lret = fSetDatoCteProv("CTIPOCLIENTE", "1");
                lret = fSetDatoCteProv("CESTATUS", "1");
                lret = fSetDatoCteProv("CIDADDENDA", "-1");

                lret = fSetDatoCteProv("CEMAIL1", _RegDoctoOrigen._RegDireccion.cEmail);
                lret = fSetDatoCteProv("CEMAIL2", _RegDoctoOrigen._RegDireccion.cEmail2);
                lret = fSetDatoCteProv("CBANCFD", "1");
                lret = fSetDatoCteProv("CTIPOENTRE", "6");





                lret = fGuardaCteProv();
                if (lret == 0)
                    return "";
                else
                    return "Error dar de alta Cliente";
        }

        public string mGrabarDireccion(long lIdDocumento)
        {
            long lret=0;
            mLeerDireccion();

            RegDireccion lRegDireccion = new RegDireccion();
            // la direccion del cliente pasarla a la direccion de la factura
            lRegDireccion = _RegDoctoOrigen._RegDireccion;
            if (lRegDireccion.cNombreCalle != null )
            {
                lret = fInsertaDireccion();
                lret = fSetDatoDireccion("cIdCatalogo", lIdDocumento.ToString());
                lret = fSetDatoDireccion("cTipoCatalogo", "3");
                lret = fSetDatoDireccion("cTipoDireccion", "0");
                lret = fSetDatoDireccion("cNombreCalle", lRegDireccion.cNombreCalle);
                if (lRegDireccion.cNumeroExterior == string.Empty)
                    lret = fSetDatoDireccion("cNumeroExterior", "0");
                else
                    lret = fSetDatoDireccion("cNumeroExterior", lRegDireccion.cNumeroExterior);
                lret = fSetDatoDireccion("cNumeroInterior", lRegDireccion.cNumeroInterior);
                lret = fSetDatoDireccion("cColonia", lRegDireccion.cColonia);
                lret = fSetDatoDireccion("cCodigoPostal", lRegDireccion.cCodigoPostal);
                lret = fSetDatoDireccion("cEstado", lRegDireccion.cEstado);
                lret = fSetDatoDireccion("cPais", lRegDireccion.cPais);
                lret = fSetDatoDireccion("cCiudad", lRegDireccion.cCiudad);
                lret = fSetDatoDireccion("cEmail", lRegDireccion.cEmail);
                lret = fGuardaDireccion();
                if (lret != 0)
                {

                    _controlfp(0x9001F, 0xFFFFF);
                    miconexion.mCerrarConexionOrigen(1);
                    return "Se presento el error direccion" + lret.ToString();

                }
            }
            return "";
        }

        private long mBuscarUltimoFolioConcepto(string aIdDocumentoModelo, string aConcepto, ref string cserie)
        {
            long x;
            miconexion.mAbrirConexionDestino();
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            
            // checar si el concepto tiene serie si la tiene buscarlo por concpeto sino por documento modelo

            string cad1 = "select m6.cseriepo01 from mgw10006 m6 where m6.ccodigoc01 = '" + aConcepto + "'";
            
            lsql.CommandText = cad1;
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lreader.Read();
                cserie = lreader[0].ToString();
                
            }
            lreader.Close();
            string cad2 ="";
            if (cserie.Trim()!="")
                cad2 = "select max(cfolio)+1  from mgw10008 m8 join mgw10006 m6 on m8.cidconce01 = m6.cidconce01 and m6.ccodigoc01 = '" + aConcepto + "'";
            else
                cad2 = "select max(cfolio)+1 from mgw10008 m8 join mgw10007 m7 on m8.ciddocum02 = m7.ciddocum01 and m7.ciddocum01 = " + aIdDocumentoModelo ;

            lsql.CommandText = cad2;
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            x = 1;
            if (lreader.HasRows)
            {
                lreader.Read();
                if (lreader[0].ToString() != "")
                    x = long.Parse(lreader[0].ToString());
            }
            else
            {
                x = 1;
            }
            lreader.Close();
            miconexion.mCerrarConexionDestino();
            return x;
        }

        protected virtual bool mActualizaDocumento5(long liddocum, int aopcion, double afolionuevo)
        {
            miconexion.mAbrirConexionDestino();
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            long cidfoldig;
            long cidconce;
            long ciddocum01 = 0;
            string cserie = "";
            double ctotal = 0;
            bool lrespuesta = false;

            //OleDbParameter lparametroIdDocumento = new OleDbParameter("@p1", _RegDoctoOrigen.cIdDocto);
            string lcadena = "update mgw10008 set cescfd = 1 where ciddocum01 = " + liddocum;

            OleDbCommand lsql1 = new OleDbCommand(lcadena, miconexion._conexion);
            try
            {

                lsql1.ExecuteNonQuery();
                string lfecha = _RegDoctoOrigen.cFecha.ToString();
                DateTime ldate = DateTime.Parse(lfecha);
                lfecha = ldate.ToString("MM/dd/yyyy");


                long ctipocfd = 0;
                //string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
                //"ConceptoDocumento"
                string lCodigoConcepto = _RegDoctoOrigen.cCodigoConcepto;
                lsql.CommandText = "select cidconce01, ciddocum01,cseriepo01, cverfacele from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    cidconce = long.Parse(lreader["cidconce01"].ToString());
                    ciddocum01 = long.Parse(lreader["ciddocum01"].ToString());
                    cserie = lreader["cseriepo01"].ToString();
                    ctipocfd = long.Parse(lreader["cverfacele"].ToString());

                }
                else
                    cidconce = 1;
                lreader.Close();

                lsql.CommandText = "select ctotal from mgw10008 where ciddocum01 = " + liddocum;
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    ctotal = double.Parse(lreader["ctotal"].ToString());

                }
                else
                    cidconce = 1;
                lreader.Close();


                double x = double.Parse(afolionuevo.ToString().Trim());

                lsql.CommandText = "select max(cidfoldig) + 1 as cidclien01 from mgw10045";
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    cidfoldig = long.Parse(lreader["cidclien01"].ToString());
                }
                else
                    cidfoldig = 1;
                lreader.Close();

                lcadena = "insert into mgw10045 (cidfoldig,ciddoctode,cidcptodoc,ciddocto,cserie,cfolio,cestado,centregado, cfechaemi,cestrad,ctotal) " +
                                 " values (" + liddocum + "," + ciddocum01 + "," + cidconce + "," + liddocum + ",'" + cserie.Trim() + "'," + x + ",1, 0, ctod('" + lfecha + "'),3," + ctotal + ")";
                //lcadena = "insert into mgw10045 (cidfoldig,ciddoctode,cidcptodoc,ciddocto,cserie,cfolio,cestado, cfechaemi,cestrad) " +
                //                 "values (8,4,3001,11,'B',444,1,ctod('" + lfecha + "'),3)";
                OleDbCommand lsql2 = new OleDbCommand(lcadena, miconexion._conexion);
                lsql1.CommandText = "SET NULL OFF";
                lsql1.ExecuteNonQuery();

                lsql2.ExecuteNonQuery();
                lrespuesta = true;
            }
            catch (Exception eee)
            {
                lrespuesta = true;
            }
            finally
            {
                miconexion.mCerrarConexionDestino();
            }
            //.mCerrarConexionDestino ();



            return lrespuesta;

        }

        private void mGrabarDirecciones(long lIdDocumento)
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            lsql.CommandText = "select alltrim(cnombrec01) + ', '" +
                                " + alltrim(cnumeroe01) + ', '" +
                                " + alltrim(ccolonia) + ', '" +
                                " + alltrim(cciudad) + ', '" +
                                " + alltrim(cestado) + ', '" +
                                " + alltrim(cpais) " +
                               " from mgw10011 where ctipocat01 = 4";

            //miconexion.mAbrirConexionDestino();
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            string ldireccion = "";
            if (lreader.HasRows)
            {
                lreader.Read();
                ldireccion = lreader[0].ToString().Trim();
                lreader.Close();
            }



            string lcadena2 = "update mgw10008 set clugarexpe = '" + ldireccion.Trim() + "'  where ciddocum01 = " + lIdDocumento;
            OleDbCommand lsql4 = new OleDbCommand(lcadena2, miconexion._conexion);
            lsql4.ExecuteNonQuery();

            if (_RegDoctoOrigen._RegDireccion.cEmail != "")
            {
                string lcadena21 = "update mgw10002 set cemail1 = '" + _RegDoctoOrigen._RegDireccion.cEmail + "', cemail2 = '" + _RegDoctoOrigen._RegDireccion.cEmail2 + "', cbancfd = 1, ctipoentre=6 where ccodigoc01 = '" + _RegDoctoOrigen.cCodigoCliente + "'";
                OleDbCommand lsql212 = new OleDbCommand(lcadena21, miconexion._conexion);
                lsql212.ExecuteNonQuery();
            }

        }

        public string mGrabarAdmNew(long afolionuevo, int opcion, bool incluyetimbrado, int tipo)
        {
            //miconexion.mAbrirConexionDestino(1);
            string lCodigoConcepto;
            //lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            lCodigoConcepto = _RegDoctoOrigen.cCodigoConcepto;
            
            string lresp1 = mGrabarEncabezado(afolionuevo, lCodigoConcepto,"0");
            if (lresp1 != "")
                return lresp1;

            string cserie;
            cserie = _RegDoctoOrigen.cSerie;
            long lIdDocumento = mBuscarIdDocumento(lCodigoConcepto, 0, cserie, long.Parse(afolionuevo.ToString().Trim()));


            if (lIdDocumento == 0)
            {

                _controlfp(0x9001F, 0xFFFFF);
                miconexion.mCerrarConexionOrigen(1);
                return "no se encontro documento " + lCodigoConcepto + " " +
                    long.Parse(GetSettingValueFromAppConfigForDLL("FolioDestino").ToString().Trim());
            }

            string lresp = mGrabarDireccion(lIdDocumento);
            lresp = mGrabarMovimientos(lIdDocumento, opcion,tipo);

            if (_RegDoctoOrigen.cCodigoConcepto == "1")
            {
                mGrabarRemision(lIdDocumento);
                lCodigoConcepto = "3";
            }

            string lrespuestas = mGrabarExtrasObservaciones(lIdDocumento );



            int lret = fAfectaDocto_Param(lCodigoConcepto, cserie, afolionuevo, true);

            //mImprimir(lIdDocumento);
            long lresp10 = 0;

            int lNumeroMov = 100;
            foreach (RegMovto mov in _RegDoctoOrigen._RegMovtos)
            {
                if (mov.cObservaciones != "" && mov.cObservaciones != null)
                {
                    string lcadenaA = "update mgw10010 set cobserva01= '" + mov.cObservaciones + "' where ciddocum01 = " + lIdDocumento.ToString() + " and cnumerom01 = " + lNumeroMov.ToString();
                    OleDbCommand lsqlA = new OleDbCommand(lcadenaA, miconexion._conexion);
                    int xxx = 0;
                    xxx = lsqlA.ExecuteNonQuery();
                }
                lNumeroMov += 100;
            }

            if (incluyetimbrado == true)
            {
                mActualizaDocumento(lIdDocumento, 1, afolionuevo);

                mGrabarDirecciones(lIdDocumento);
        
                lresp10 = fInicializaLicenseInfo(0);
                if (lresp10 == 0)
                {
                    //int lresp20 = fEmitirDocumento(string aCodigoConcepto, string aNumSerie, double aFolio, string aPassword, string aArchivo);

                    //Properties.Settings.Default.Pass = textBox3.Text;
                    string lpass = "";
                    lpass = GetSettingValueFromAppConfigForDLL("Pass").ToString().Trim();


                    int lresp20 = fEmitirDocumento(_RegDoctoOrigen.cCodigoConcepto, _RegDoctoOrigen.cSerie, _RegDoctoOrigen.cFolio, lpass, "");
                    /*
                    string lformatin = @"\\TOSHIBA-PC\Empresas\Reportes\AdminPAQ\Plantilla_Factura_CFDi_1.htm";
                    lformatin = @"c:\compacw\Empresas\Reportes\AdminPAQ\Plantilla_Factura_CFDi_1.htm";
                    string lserie = _RegDoctoOrigen.cSerie;
                    string lcodigo = _RegDoctoOrigen.cCodigoConcepto;
                    //int lError = fEntregEnDiscoXML (lcodigo, lserie, _RegDoctoOrigen.cFolio, 1, ref lformatin);
                    */
                }
                //miconexion.mCerrarConexionOrigen(1);

                //miconexion.mCerrarConexionDestino(1);
            }

            try
            {
                _controlfp(0x9001F, 0xFFFFF);
            }
            catch (Exception eee)
            {
                lrespuestas = eee.Message;
            }
            // barra.Asignar(100);
            return lrespuestas;
        }

        private string mGrabarRemision(long aIdDocumento)
        {
            long x;
        //    miconexion.mAbrirConexionDestino();
            string cad = "select cidmovim01, cunidades from mgw10010 where ciddocum01 = " + aIdDocumento + " order by cidmovim01 ";

            OleDbCommand lsql = new OleDbCommand(cad, miconexion._conexion);
            OleDbDataReader lreader;
            
            lreader = lsql.ExecuteReader();
            x = 1;
            long idmov1=0;
            long idmov2=0;
            decimal lunidades = 0;
            OleDbCommand lsql4 = new OleDbCommand();
            string lcadena2 = "";
            if (lreader.HasRows)
            {
                while (lreader.Read())
                {
                    if (x % 2 == 0) // movto 2
                    {
                        idmov2 = long.Parse(lreader[0].ToString());
                        lunidades += decimal.Parse(lreader[1].ToString());
                        lcadena2 = "update mgw10010 set ciddocum02 =3, ciddocum01 = 0, cnumerom01=0,cafectae01 = 1, cafectad01 = 0, cmovtooc01 = 1, cidmovto01 = " + idmov1 + " where cidmovim01 = " + idmov2;
                        lsql4.CommandText = lcadena2;
                        lsql4.Connection = miconexion._conexion;
                        lsql4.ExecuteNonQuery();
            
                    }
                    else
                    {
                        idmov1 = long.Parse(lreader[0].ToString());
                        
                        lcadena2 = "update mgw10010 set ciddocum02 = 3, cafectae01 = 2, cafectad01 = 0, cmovtooc01 = 0  where cidmovim01 = " + idmov1;

                        lsql4.CommandText = lcadena2;
                        lsql4.Connection = miconexion._conexion;
                        lsql4.ExecuteNonQuery();
                    }
                    x++;
                    
                }
            }
            lreader.Close();

            lcadena2 = "update mgw10008 set ciddocum02 = 3, cidconce01= 3, ctotalun01 = " + lunidades + ", cunidade01 = " + lunidades+ " where ciddocum01 = " + aIdDocumento;
            //lcadena2 = "update mgw10008 set ciddocum02 = 3, cidconce01= 3 where ciddocum01 = " + aIdDocumento;
            // cambiar el concepto al documento
            lsql4.CommandText = lcadena2;
            lsql4.Connection = miconexion._conexion;
            lsql4.ExecuteNonQuery();
          //  miconexion.mCerrarConexionDestino();
            return "";
        }

        private void mImprimir(long aIdDocumento)
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;

            string lcadena2 = "update mgw10008 set cimpreso = 1 where ciddocum01 = " + aIdDocumento;

                OleDbCommand lsql4 = new OleDbCommand(lcadena2, miconexion._conexion);
                lsql4.ExecuteNonQuery();
                //lrespuesta = fAfectaDocto_Param(lCodigoConcepto, cserie , x, true);

            
        }

        private string mGrabarExtrasObservaciones(long lIdDocumento)
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;

            string lresp = "";
            long lrespuesta = 0;
            string lrespuestas = "";
            
            //miconexion.mAbrirConexionDestino();

               


                string lcadena2 = "update mgw10008 set cobserva01 = '" + _RegDoctoOrigen.cTextoExtra3 + "' where ciddocum01 = " + lIdDocumento;

                OleDbCommand lsql4 = new OleDbCommand(lcadena2, miconexion._conexion);
                lsql4.ExecuteNonQuery();

                if (_RegDoctoOrigen._RegMovtos.Count > 0)
                {

                    string lcadenaA = "update mgw10010 set cobserva01= '" + _RegDoctoOrigen._RegMovtos[0].ctextoextra3 + "' where ciddocum01 = " + lIdDocumento;
                    OleDbCommand lsqlA = new OleDbCommand(lcadenaA, miconexion._conexion);
                    lsqlA.ExecuteNonQuery();
                }

                //miconexion.mCerrarConexionDestino();

            
            return lrespuestas;
            

        }

        private string mGrabarExtras(long lIdDocumento,int opcion, double afolionuevo)
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            
            string lresp= "";
            long lrespuesta = 0;
            string lrespuestas = "";
            if (lresp == "")
            {
                //double x = double.Parse(afolionuevo.ToString () );
                if (opcion == 2)
                    mActualizaDocumento(lIdDocumento, opcion, afolionuevo);

                lsql.CommandText = "select alltrim(cnombrec01) + ', '" +
                                    " + alltrim(cnumeroe01) + ', '" +
                                    " + alltrim(ccolonia) + ', '" +
                                    " + alltrim(cciudad) + ', '" +
                                    " + alltrim(cestado) + ', '" +
                                    " + alltrim(cpais) " +
                                   " from mgw10011 where ctipocat01 = 4";

                //miconexion.mAbrirConexionDestino();
                miconexion.mAbrirConexionDestino();
            
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                string ldireccion = "";
                if (lreader.HasRows)
                {
                    lreader.Read();
                    ldireccion = lreader[0].ToString().Trim();
                    lreader.Close();
                }



                string lcadena2 = "update mgw10008 set cobserva01 = '" + _RegDoctoOrigen.cTextoExtra1 + "', clugarexpe = '" + ldireccion.Trim() + "'  where ciddocum01 = " + lIdDocumento;

                OleDbCommand lsql4 = new OleDbCommand(lcadena2, miconexion._conexion);
                lsql4.ExecuteNonQuery();
                //lrespuesta = fAfectaDocto_Param(lCodigoConcepto, cserie , x, true);

                if (opcion == 1 && _RegDoctoOrigen.cTipoCambio != 1)
                {

                    //miconexion.mAbrirConexionDestino(1);
                    //                    string lcadena1 = "update mgw10008 set cpendiente = ctotal where ciddocum01 = " + lIdDocumento;

                    double ltotal = _RegDoctoOrigen.cImpuestos + _RegDoctoOrigen.cNeto;
                    string lcadena1 = "update mgw10008 set ctipocam01 = " + _RegDoctoOrigen.cTipoCambio + "  where ciddocum01 = " + lIdDocumento;

                    OleDbCommand lsql3 = new OleDbCommand(lcadena1, miconexion._conexion);
                    // lsql3.ExecuteNonQuery();



                }

                /* actualizar observaciones del movimiento */

                string lcadenaA = "update mgw10010 set cobserva01= '" + _RegDoctoOrigen.cTextoExtra1 + "' where ciddocum01 = " + lIdDocumento;
                    OleDbCommand lsqlA= new OleDbCommand(lcadenaA, miconexion._conexion);
                    lsqlA.ExecuteNonQuery();
                


                
                if (_RegDoctoOrigen._RegDireccion.cEmail != "")
                {

                    string lcadena21 = "update mgw10002 set cemail1 = '" + _RegDoctoOrigen._RegDireccion.cEmail + "', cemail2 = '" + _RegDoctoOrigen._RegDireccion.cEmail2 + "', cbancfd = 1, ctipoentre=6 where ccodigoc01 = '" + _RegDoctoOrigen.cCodigoCliente + "'";
                    OleDbCommand lsql212 = new OleDbCommand(lcadena21, miconexion._conexion);
                    lsql212.ExecuteNonQuery();
                }


                if (opcion == 3 || opcion == 4)
                {

                    //miconexion.mAbrirConexionDestino(1);
                    //                    string lcadena1 = "update mgw10008 set cpendiente = ctotal where ciddocum01 = " + lIdDocumento;

                    double ltotal = _RegDoctoOrigen.cImpuestos + _RegDoctoOrigen.cImpuesto2 + _RegDoctoOrigen.cNeto;
                    decimal limpuestos = decimal.Parse(_RegDoctoOrigen.cImpuestos.ToString());
                    decimal limpuestos2 = decimal.Parse(_RegDoctoOrigen.cImpuesto2.ToString());
                    limpuestos = decimal.Round(limpuestos, 4);
                    string lcadena1 = "update mgw10008 set cneto = " + _RegDoctoOrigen.cNeto.ToString() + ", cimpuesto2 = " + limpuestos + ", cimpuesto3 = " + limpuestos2 + ",ctotal = " + ltotal.ToString() + ",cpendiente = " + ltotal.ToString() + ",ctipocam01 = " + _RegDoctoOrigen.cTipoCambio + " where ciddocum01 = " + lIdDocumento;

                    OleDbCommand lsql3 = new OleDbCommand(lcadena1, miconexion._conexion);
                    lsql3.ExecuteNonQuery();
                    //miconexion.mCerrarConexionDestino();
                }

                


                miconexion.mCerrarConexionDestino();

                //mGrabarInterfaz(afolioant, opcion);
            }
            else
            {
                lrespuestas = "ocurrio error";
            }
            //miconexion.mCerrarConexionDestino ();
            return lrespuestas;
            // antes de cerrar grabar en la tabla de interfaz

        }


        public string mGrabarAdm5(string afolioant, double afolionuevo, int opcion)
        {

            //mGrabarEncabezado(afolionuevo);


            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            string cserie;

            long lret;
            string lCodigoConcepto;
                lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            
            miconexion.mAbrirConexionDestino(1);
            
            //long lidconce = 0;
            //long tipocfd = 0;
            //lsql.CommandText = "select cidconce01, ciddocum01,cseriepo01,cverfacele from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
            //lsql.Connection = miconexion._conexion;
            //lreader = lsql.ExecuteReader();
            ////_RegDoctoOrigen._RegMovtos.Clear();
            //if (lreader.HasRows)
            //{
            //    lreader.Read();
            //    cserie = lreader["cseriepo01"].ToString();
            //    lidconce = long.Parse(lreader["cidconce01"].ToString());
            //    tipocfd = long.Parse(lreader["cverfacele"].ToString());
            //}
            //else
            //    cserie = "";
            //lreader.Close();


            //lsql.CommandText = "select count(*) as cuantos from mgw10008 where cidconce01 = " + lidconce + " and cseriedo01 = '" + cserie + "' and cfolio = " + afolionuevo.ToString().Trim();
            //lsql.Connection = miconexion._conexion;
            //lreader = lsql.ExecuteReader();
            ////_RegDoctoOrigen._RegMovtos.Clear();
            //if (lreader.HasRows)
            //{
            //    lreader.Read();
            //    long cuantos=0;
            //    cuantos = long.Parse(lreader["cuantos"].ToString());
            //    lreader.Close();
            //    if (cuantos > 0)
            //    {
            //        _controlfp(0x9001F, 0xFFFFF);
            //        miconexion.mCerrarConexionOrigen(1);
            //        return "Documento ya existe en ADMINPAQ";
            //    }
            //}
            //lreader.Close();



            //fInsertarDocumento();
            //lret = fSetDatoDocumento("cCodigoConcepto", lCodigoConcepto);

            
            //if (_RegDoctoOrigen.cSerie != "")
            //    lret = fSetDatoDocumento("cSerieDocumento",  _RegDoctoOrigen.cSerie  );
            //else
            //    lret = fSetDatoDocumento("cSerieDocumento", "");
            //lret = fSetDatoDocumento("CCODIGOCLIENTE", _RegDoctoOrigen.cCodigoCliente);
            //lret = fSetDatoDocumento("cCodigoCteProv", _RegDoctoOrigen.cCodigoCliente);
            //if (lret != 0)
            //{
            //    fInsertaCteProv();
            //    lret = fSetDatoCteProv("CCODIGOCLIENTE", _RegDoctoOrigen.cCodigoCliente);
            //    lret = fSetDatoCteProv("cRazonSocial", _RegDoctoOrigen.cRazonSocial);
            //    if (lret != 0)
            //    {
            //        _controlfp(0x9001F, 0xFFFFF);
            //        miconexion.mCerrarConexionOrigen(1);
            //        return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cRazonSocial;
            //    }
            //    lret = fSetDatoCteProv("cRFC", _RegDoctoOrigen.cRFC);
            //    if (lret != 0)
            //    {
            //        _controlfp(0x9001F, 0xFFFFF);
            //        miconexion.mCerrarConexionOrigen(1);
            //        return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cRFC;
            //    }
            //    lret = fSetDatoCteProv("CLISTAPRECIOCLIENTE", "1");
            //    lret = fSetDatoCteProv("CIDMONEDA", "1");

            //    string lfecha = _RegDoctoOrigen.cFecha.ToString();
            //    DateTime ldate = DateTime.Parse(lfecha);
            //    lfecha = ldate.ToString("MM/dd/yyyy");
            //    lret = fSetDatoCteProv("CFECHAALTA", lfecha);
            //    if (lret != 0)
            //    {
            //        _controlfp(0x9001F, 0xFFFFF);
            //        miconexion.mCerrarConexionOrigen(1);
            //        return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cFecha.ToString();
            //    }
            //    lret = fSetDatoCteProv("CTIPOCLIENTE", "1");
            //    lret = fSetDatoCteProv("CESTATUS", "1");
            //    lret = fSetDatoCteProv("CIDADDENDA", "-1");

            //    lret = fSetDatoCteProv("CEMAIL1", _RegDoctoOrigen._RegDireccion.cEmail);
            //    lret = fSetDatoCteProv("CEMAIL2", _RegDoctoOrigen._RegDireccion.cEmail2);
            //    lret = fSetDatoCteProv("CBANCFD", "1");
            //    lret = fSetDatoCteProv("CTIPOENTRE", "6");





            //    lret = fGuardaCteProv();
            //    if (lret != 0)
            //    {
            //        _controlfp(0x9001F, 0xFFFFF);
            //        bool sigue = false;
            //        sigue = mDarAltaCliente();
            //        if (sigue == true)
            //            lret = fSetDatoDocumento("cCodigoCteProv", _RegDoctoOrigen.cCodigoCliente);
            //        else
            //        {
            //            miconexion.mCerrarConexionOrigen(1);
            //            return "Se presento el error en clientes111 " + lret.ToString();
            //        }

            //    }
            //    else
            //        lret = fSetDatoDocumento("cCodigoCteProv", _RegDoctoOrigen.cCodigoCliente);

            //    if (lret != 0)
            //    {
            //        _controlfp(0x9001F, 0xFFFFF);
            //        miconexion.mCerrarConexionOrigen(1);
            //        return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cCodigoCliente;
            //    }
            
            //}
            /*
            if (primerdocto.cCodigoCliente == null)
            {
                mModificaDatosClienteFlexo();
                primerdocto.cCodigoCliente = _RegDoctoOrigen.cCodigoCliente;
                primerdocto.cRazonSocial = _RegDoctoOrigen.cRazonSocial;
                primerdocto.cRFC = _RegDoctoOrigen.cRFC;
                primerdocto.cCond = _RegDoctoOrigen.cCond;
                primerdocto.cAgente = _RegDoctoOrigen.cAgente;
                primerdocto._RegDireccion = _RegDoctoOrigen._RegDireccion;
            }

            else
            {
                _RegDoctoOrigen.cCodigoCliente = primerdocto.cCodigoCliente;
                _RegDoctoOrigen.cRazonSocial = primerdocto.cRazonSocial;
                _RegDoctoOrigen.cRFC = primerdocto.cRFC;
                _RegDoctoOrigen.cCond = primerdocto.cCond;
                _RegDoctoOrigen.cAgente = primerdocto.cAgente;
                _RegDoctoOrigen._RegDireccion = primerdocto._RegDireccion;

            }
             */
            
            
            //lret = fSetDatoDocumento("cRazonSocial", _RegDoctoOrigen.cRazonSocial );
            //lret = fSetDatoDocumento("cRFC", _RegDoctoOrigen.cRFC );
            //if (_RegDoctoOrigen.cMoneda != "Pesos") 
            //    lret = fSetDatoDocumento("cIdMoneda", "2");
            //else
            //    lret = fSetDatoDocumento("cIdMoneda", "1");
            //lret = fSetDatoDocumento("cTipoCambio", _RegDoctoOrigen.cTipoCambio.ToString ());
            //lret = fSetDatoDocumento("cReferencia", "Por Programa");
            ////lret = fSetDatoDocumento("cObservaciones", _RegDoctoOrigen.cTextoExtra1 );
            //lret = fSetDatoDocumento("cFolio", _RegDoctoOrigen.cFolio.ToString().Trim());


            //try
            //{
            //    lret = fSetDatoDocumento("cReferencia", _RegDoctoOrigen.cFolio.ToString ());
            //    lret = fSetDatoDocumento("cTextoExtra1", _RegDoctoOrigen.cReferencia);
            //}
            //catch (Exception ee)
            //{ 
            //}

            ////lret = fSetDatoDocumento("cEsCFD", "1");
            ////lret = fSetDatoDocumento("cSerieDocumento", GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim ()   );
            //string lfechavenc = "";
            //if (opcion == 1)
            //{
            //    DateTime lFechaVencimiento;
            //    lFechaVencimiento = _RegDoctoOrigen.cFecha.AddDays(int.Parse("0"));

            //    //lFechaVencimiento = DateTime.Today.AddDays(int.Parse(_RegDoctoOrigen.cCond) );
                
            //    lfechavenc = "";
            //    lfechavenc = String.Format("{0:MM/dd/yyyy}", lFechaVencimiento); ;  // "8 08 008 2008"   year
            //    lret = fSetDatoDocumento("cFechaVencimiento", lfechavenc);
            //    //lret = fSetDatoDocumento("cCodigoAgente", _RegDoctoOrigen.cAgente );
            //    if (lret != 0)
            //    {
            //        miconexion.mCerrarConexionOrigen(1);
            //        _controlfp(0x9001F, 0xFFFFF);
            //        // barra.Asignar(100);
            //        return "Agente no existe";
            //    }


            //}
            ////lret = fSetDatoDocumento("cImpuesto1", _RegDoctoOrigen.cImpuestos.ToString ());
            

            //string lfechadocto = "";
            //lfechadocto = _RegDoctoOrigen.cFecha.ToString();
            //DateTime lFechaDocto;
            //lFechaDocto = _RegDoctoOrigen.cFecha;

            //lfechadocto = "";
            

            //lfechadocto = String.Format("{0:MM/dd/yyyy}", lFechaDocto); ;  // "8 08 008 2008"   year
            ////if (opcion == 3 || opcion == 4)
            ////    lfechadocto = String.Format("{0:MM/dd/yyyy}", DateTime.Today); ;  // "8 08 008 2008"   year

            ////lfechadocto = String.Format("{0:MM/dd/yyyy}", lFechaDocto); ;  
            
            //lret = fSetDatoDocumento("cFecha", lfechadocto);

            //if (opcion == 1)
            //    lret = fSetDatoDocumento("cFechaVencimiento", lfechadocto);
            //lret = fSetDatoDocumento("cTipoCambio", "1");
            //lret = fGuardaDocumento();
            //if (lret != 0)
            //{

            //    _controlfp(0x9001F, 0xFFFFF); 
            //    miconexion.mCerrarConexionOrigen(1);
            //    return "Se presento el error " + lret.ToString () ;
                
            //}

            //lret = fSetDatoDocumento("cCodigoConcepto", "10");
            //lret = fGuardaDocumento();
            

            // buscar el id del documento generado
            //long lIdDocumento = mBuscarIdDocumento(lCodigoConcepto, 0, GetSettingValueFromAppConfigForDLL("SerieDestino").ToString().Trim(), long.Parse(GetSettingValueFromAppConfigForDLL("FolioDestino").ToString().Trim()));
            cserie = _RegDoctoOrigen.cSerie;
            long lIdDocumento = mBuscarIdDocumento(lCodigoConcepto, 0, cserie , long.Parse (afolionuevo.ToString().Trim()));
                

            if (lIdDocumento == 0)
            {

                _controlfp(0x9001F, 0xFFFFF);
                miconexion.mCerrarConexionOrigen(1);
                return "no se encontro documento " + lCodigoConcepto + " " +
                    long.Parse(GetSettingValueFromAppConfigForDLL("FolioDestino").ToString().Trim()); 
//                +" " + 
  //                  lret.ToString();

            }

            string lresp = mGrabarDireccion(lIdDocumento);

            //mLeerDireccion();

            //RegDireccion lRegDireccion = new RegDireccion();
            //// la direccion del cliente pasarla a la direccion de la factura
            //lRegDireccion = _RegDoctoOrigen._RegDireccion;
            //if (lRegDireccion.cNombreCalle != null )
            //{
            //    lret = fInsertaDireccion();
            //    lret = fSetDatoDireccion("cIdCatalogo", lIdDocumento.ToString());
            //    lret = fSetDatoDireccion("cTipoCatalogo", "3");
            //    lret = fSetDatoDireccion("cTipoDireccion", "0");
            //    lret = fSetDatoDireccion("cNombreCalle", lRegDireccion.cNombreCalle);
            //    if (lRegDireccion.cNumeroExterior == string.Empty)
            //        lret = fSetDatoDireccion("cNumeroExterior", "0");
            //    else
            //        lret = fSetDatoDireccion("cNumeroExterior", lRegDireccion.cNumeroExterior);
            //    lret = fSetDatoDireccion("cNumeroInterior", lRegDireccion.cNumeroInterior);
            //    lret = fSetDatoDireccion("cColonia", lRegDireccion.cColonia);
            //    lret = fSetDatoDireccion("cCodigoPostal", lRegDireccion.cCodigoPostal);
            //    lret = fSetDatoDireccion("cEstado", lRegDireccion.cEstado);
            //    lret = fSetDatoDireccion("cPais", lRegDireccion.cPais);
            //    lret = fSetDatoDireccion("cCiudad", lRegDireccion.cCiudad);
            //    lret = fSetDatoDireccion("cEmail", lRegDireccion.cEmail);
            //    lret = fGuardaDireccion();
            //    if (lret != 0)
            //    {

            //        _controlfp(0x9001F, 0xFFFFF);
            //        miconexion.mCerrarConexionOrigen(1);
            //        return "Se presento el error direccion" + lret.ToString();

            //    }
            //}

            lresp = mGrabarMovimientos(lIdDocumento,opcion,0);


            //long lNumeroMov = 100;
            //if (_RegDoctoOrigen._RegMovtos.Count == 0  && (opcion ==3 || opcion == 4))
            //{

            //    RegMovto lRegmovto = new RegMovto();
            //    lRegmovto.cCodigoProducto = "(Ninguno)";
            //    lRegmovto.cNombreProducto = "(Ninguno)";
            //    lRegmovto.cIdDocto = long.Parse(_RegDoctoOrigen.cIdDocto.ToString());
            //    lRegmovto.cSubtotal  = decimal.Parse(_RegDoctoOrigen.cNeto.ToString () ); 
            //    //lRegmovto.cTotal = decimal.Parse(lreader["cunidades"].ToString());
            //    lRegmovto.cImpuesto = decimal.Parse(_RegDoctoOrigen.cImpuestos.ToString());
            //    lRegmovto.cTotal = decimal.Parse(_RegDoctoOrigen.cNeto.ToString());
            //    lRegmovto.cCodigoAlmacen = "(Ninguno)";
            //    lRegmovto.cNombreAlmacen = "(Ninguno)";
            //    lRegmovto.cUnidad = "";
            //    lRegmovto.cUnidades = 1;
            //    lRegmovto.cReferencia = "";
            //    lRegmovto.ctextoextra1 = "";
            //    lRegmovto.ctextoextra2 = "";
            //    lRegmovto.ctextoextra3 = "";
            //    _RegDoctoOrigen._RegMovtos.Add(lRegmovto); 
            //}
            //foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
            //{
            //    lret = fInsertarMovimiento();
            //    lret = fSetDatoMovimiento("cIdDocumento", lIdDocumento.ToString());
            //    lret = fSetDatoMovimiento("cNumeroMovimiento", lNumeroMov.ToString());
            //    lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
            //    if (lret != 0)
            //    {
            //        fInsertaProducto();
            //        lret = fSetDatoProducto("CCODIGOPRODUCTO", x.cCodigoProducto);
            //        lret = fSetDatoProducto("CNOMBREPRODUCTO", x.cNombreProducto );
            //        lret = fSetDatoProducto("CTIPOPRODUCTO", "1");
            //        lret = fSetDatoProducto("CMETODOCOSTEO", "1");
            //        lret = fSetDatoProducto("CCONTROLEXISTENCIA", "1");
            //        lret = fSetDatoProducto("CIMPUESTO1", x.cPorcent01.ToString () );
            //        OleDbCommand cmdunidad = new OleDbCommand();
            //        cmdunidad.CommandText = "select * from mgw10026 where cnombreu01 = '" + x.cUnidad.ToUpper() + "'";
            //        miconexion.mAbrirConexionDestino();
            //        cmdunidad.Connection = miconexion._conexion  ;
            //        OleDbDataReader ldr = cmdunidad.ExecuteReader() ;
            //        int lidunidad ;
            //        if (ldr.HasRows == false)
            //        {
            //            ldr.Read();
            //            ldr.Close();
            //            lret = fSetDatoProducto("CCODIGOUNIDADBASE", x.cUnidad.ToUpper());
            //            if (lret != 0)
            //            {
            //                // dar de alta la unicad de medida y peso
            //                cmdunidad.CommandText = "select max(cidunidad) + 1 from mgw10026";
            //                ldr = cmdunidad.ExecuteReader();
            //                ldr.Read();

            //                 lidunidad = int.Parse(ldr[0].ToString());
            //                ldr.Close();
            //                cmdunidad.CommandText = "insert into mgw10026 values (" + lidunidad + ",'" + x.cUnidad.ToUpper() + "','" + x.cUnidad.ToUpper() + "','" + x.cUnidad.ToUpper() + "','')";
            //                cmdunidad.ExecuteNonQuery();
            //                lret = fSetDatoProducto("CIDUNIDADBASE", lidunidad.ToString());
            //            }
            //        }
            //        else
            //        {
            //            ldr.Read();
                        
            //            lidunidad = int.Parse(ldr[0].ToString());
            //            ldr.Close();
            //            lret = fSetDatoProducto("CIDUNIDADBASE", lidunidad.ToString());
            //        }
            //        lret = fGuardaProducto();
            //        lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
            //    }
            //    lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);
            //    if (lret != 0)
            //    {
            //        fInsertaAlmacen();
            //        lret = fSetDatoAlmacen("CCODIGOALMACEN", x.cCodigoAlmacen);
            //        lret = fSetDatoAlmacen("CNOMBREALMACEN", x.cNombreAlmacen);
            //        lret = fGuardaAlmacen();
            //        lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);

            //    }
            //    decimal total ;
            //    if (opcion == 3 || opcion == 4)
            //    {
            //        lret = fGuardaMovimiento();
            //        if (opcion == 4)
            //        {
            //            //lret = fSetDatoMovimiento("cNETO", x.cSubtotal.ToString());
            //            //lret = fSetDatoMovimiento("cTotal", x.cSubtotal.ToString());
            //            //string lcadena = "update mgw10010 set cneto = " + x.cSubtotal + ", ctotal = " + x.cSubtotal + " where ciddocum01 = " + lIdDocumento;
            //            total = x.cSubtotal + x.cImpuesto;
            //            string lcadena55 = "update mgw10010 set cneto = " + x.cSubtotal + ", cimpuesto1 = " + x.cImpuesto + ", ctotal = " + total +"  where ciddocum01 = " + lIdDocumento;
            //            OleDbCommand lsql22 = new OleDbCommand(lcadena55, miconexion._conexion);
            //            lsql22.ExecuteNonQuery();
            //        }
            //        else
            //        {
            //             total = x.cSubtotal + x.cImpuesto;
            //             string lcadena44 = "update mgw10010 set cneto = " + x.cSubtotal + ", cimpuesto1 = " + x.cImpuesto + ", ctotal = " + total + " where ciddocum01 = " + lIdDocumento;
            //            OleDbCommand lsql3 = new OleDbCommand(lcadena44, miconexion._conexion);
            //            lsql3.ExecuteNonQuery();
            //        }

            //        //lret = fSetDatoMovimiento("cImpuesto1", x.cImpuesto.ToString());
            //    }
            //    else
            //    {
                    
            //        lret = fSetDatoMovimiento("cUnidades", x.cUnidades.ToString());
            //        lret = fSetDatoMovimiento("cPrecio", x.cPrecio.ToString());
            //        lret = fSetDatoMovimiento("cporcentajeimpuesto1", x.cPorcent01.ToString());

            //        try
            //        {
            //            lret = fSetDatoMovimiento("ctextoextra1", x.ctextoextra1);
            //            lret = fSetDatoMovimiento("cReferencia", x.cReferencia);
            //            lret = fSetDatoMovimiento("ctextoextra2", x.ctextoextra2);
            //            lret = fSetDatoMovimiento("ctextoextra3", x.ctextoextra3);
            //        }
            //        catch (Exception ee)
            //        { }

            //        lret = fGuardaMovimiento();
            //    }

                
            //    lNumeroMov += 100;

            //}
            long lrespuesta = 0;
            string lrespuestas = "";
            if (lresp == "")
            {
                //double x = double.Parse(afolionuevo.ToString () );
                    mActualizaDocumento(lIdDocumento, opcion, afolionuevo );

                    lsql.CommandText = "select alltrim(cnombrec01) + ', '" + 
                                        " + alltrim(cnumeroe01) + ', '" +  
                                        " + alltrim(ccolonia) + ', '" +  
                                        " + alltrim(cciudad) + ', '" + 
                                        " + alltrim(cestado) + ', '" + 
                                        " + alltrim(cpais) " +   
                                       " from mgw10011 where ctipocat01 = 4";

                    miconexion.mAbrirConexionDestino();
                    lsql.Connection = miconexion._conexion;
                    lreader = lsql.ExecuteReader();
                    string ldireccion= "" ;
                    if (lreader.HasRows)
                    {
                        lreader.Read();
                        ldireccion = lreader[0].ToString().Trim();
                        lreader.Close();
                    }


                    
                    string lcadena2 = "update mgw10008 set cobserva01 = '" + _RegDoctoOrigen.cTextoExtra1 + "', clugarexpe = '"+ ldireccion.Trim () + "'  where ciddocum01 = " + lIdDocumento;

                    OleDbCommand lsql4 = new OleDbCommand(lcadena2, miconexion._conexion);
                    lsql4.ExecuteNonQuery();
                //lrespuesta = fAfectaDocto_Param(lCodigoConcepto, cserie , x, true);
                    
                if (opcion == 1 && _RegDoctoOrigen.cTipoCambio != 1 )
                {

                    //miconexion.mAbrirConexionDestino(1);
                    //                    string lcadena1 = "update mgw10008 set cpendiente = ctotal where ciddocum01 = " + lIdDocumento;
                    
                    double ltotal = _RegDoctoOrigen.cImpuestos + _RegDoctoOrigen.cNeto;
                    string lcadena1 = "update mgw10008 set ctipocam01 = " + _RegDoctoOrigen.cTipoCambio + "  where ciddocum01 = " + lIdDocumento;

                    OleDbCommand lsql3 = new OleDbCommand(lcadena1, miconexion._conexion);
                   // lsql3.ExecuteNonQuery();

                    
                    
                }
                if (_RegDoctoOrigen._RegDireccion.cEmail != "")
                {
                
                    string lcadena21 = "update mgw10002 set cemail1 = '" + _RegDoctoOrigen._RegDireccion.cEmail + "', cemail2 = '" + _RegDoctoOrigen._RegDireccion.cEmail2 + "', cbancfd = 1, ctipoentre=6 where ccodigoc01 = '" + _RegDoctoOrigen.cCodigoCliente + "'";
                    OleDbCommand lsql212 = new OleDbCommand(lcadena21, miconexion._conexion);
                    lsql212.ExecuteNonQuery();
                }
                

                if (opcion == 3 || opcion == 4)
                {
                    
                    //miconexion.mAbrirConexionDestino(1);
//                    string lcadena1 = "update mgw10008 set cpendiente = ctotal where ciddocum01 = " + lIdDocumento;

                    double ltotal = _RegDoctoOrigen.cImpuestos + _RegDoctoOrigen.cNeto;
                    decimal limpuestos = decimal.Parse(_RegDoctoOrigen.cImpuestos.ToString ());
                    limpuestos = decimal.Round(limpuestos,4);
                    string lcadena1 = "update mgw10008 set cneto = "  +  _RegDoctoOrigen.cNeto.ToString () + ", cimpuesto1 = " + limpuestos  + ",ctotal = " + ltotal.ToString () + ",cpendiente = " + ltotal.ToString () + ",ctipocam01 = " + _RegDoctoOrigen.cTipoCambio + " where ciddocum01 = " + lIdDocumento;

                    OleDbCommand lsql3 = new OleDbCommand(lcadena1, miconexion._conexion);
                    lsql3.ExecuteNonQuery();
                    //miconexion.mCerrarConexionDestino();
                }
                miconexion.mCerrarConexionDestino();

                //mGrabarInterfaz(afolioant, opcion);
            }
            else
            {
                lrespuestas = "ocurrio error";
            }
            //miconexion.mCerrarConexionDestino ();

            // antes de cerrar grabar en la tabla de interfaz




            miconexion.mCerrarConexionOrigen(1);
            //miconexion.mCerrarConexionDestino(1);

            try
            {
                _controlfp(0x9001F, 0xFFFFF);
            }
            catch (Exception eee)
            {
                lrespuestas = eee.Message;
            }
           // barra.Asignar(100);
            return lrespuestas;
        }

        private string mGrabarMovimientos(long lIdDocumento, int opcion, int tipo)
        {
            long lret = 0;
            long lNumeroMov = 100;
            if (_RegDoctoOrigen._RegMovtos.Count == 0 && (opcion == 3 || opcion == 4))
            {

                RegMovto lRegmovto = new RegMovto();
                lRegmovto.cCodigoProducto = "(Ninguno)";
                lRegmovto.cNombreProducto = "(Ninguno)";
                //lRegmovto.cIdDocto = long.Parse(_RegDoctoOrigen.cIdDocto.ToString());
                lRegmovto.cIdDocto = lIdDocumento;
                lRegmovto.cSubtotal = decimal.Parse(_RegDoctoOrigen.cNeto.ToString());
                //lRegmovto.cTotal = decimal.Parse(lreader["cunidades"].ToString());
                lRegmovto.cImpuesto = decimal.Parse(_RegDoctoOrigen.cImpuestos.ToString());
                lRegmovto.cImpuesto2 = decimal.Parse(_RegDoctoOrigen.cImpuesto2.ToString());
                
                lRegmovto.cTotal = decimal.Parse(_RegDoctoOrigen.cNeto.ToString());
                lRegmovto.cCodigoAlmacen = "(Ninguno)";
                lRegmovto.cNombreAlmacen = "(Ninguno)";
                lRegmovto.cUnidad = "";
                lRegmovto.cUnidades = 1;
                lRegmovto.cReferencia = "";
                lRegmovto.ctextoextra1 = "";
                lRegmovto.ctextoextra2 = "";
                lRegmovto.ctextoextra3 = "";
                _RegDoctoOrigen._RegMovtos.Add(lRegmovto);
            }
            foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
            {
                lret = fInsertarMovimiento();
                lret = fSetDatoMovimiento("cIdDocumento", lIdDocumento.ToString());
                lret = fSetDatoMovimiento("cNumeroMovimiento", lNumeroMov.ToString());
                lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
                //lret = fSetDatoMovimiento("cObserva01", x.ctextoextra3);
                if (lret != 0)
                {
                    fInsertaProducto();
                    lret = fSetDatoProducto("CCODIGOPRODUCTO", x.cCodigoProducto);
                    lret = fSetDatoProducto("CNOMBREPRODUCTO", x.cNombreProducto);
                    lret = fSetDatoProducto("CTIPOPRODUCTO", "3");
                    lret = fSetDatoProducto("CMETODOCOSTEO", "1");
                    lret = fSetDatoProducto("CCONTROLEXISTENCIA", "1");
                    x.ctextoextra1 = "";
                    lret = fSetDatoDocumento("COBSERVACIONES", x.cObservaciones);
                    //lret = fSetDatoProducto("CIMPUESTO1", x.cPorcent01.ToString());

                    lret = fSetDatoProducto("CIMPUESTO1", "0");
                    lret = fSetDatoProducto("CIMPUESTO2", x.cImpuesto.ToString());
                    OleDbCommand cmdunidad = new OleDbCommand();
                    cmdunidad.CommandText = "select * from mgw10026 where cnombreu01 = '" + x.cUnidad.ToUpper() + "'";
                    miconexion.mAbrirConexionDestino();
                    cmdunidad.Connection = miconexion._conexion;
                    OleDbDataReader ldr = cmdunidad.ExecuteReader();
                    int lidunidad;
                    if (ldr.HasRows == false)
                    {
                        ldr.Read();
                        ldr.Close();
                        lret = fSetDatoProducto("CCODIGOUNIDADBASE", x.cUnidad.ToUpper());
                        if (lret != 0)
                        {
                            // dar de alta la unicad de medida y peso
                            cmdunidad.CommandText = "select max(cidunidad) + 1 from mgw10026";
                            ldr = cmdunidad.ExecuteReader();
                            ldr.Read();

                            lidunidad = int.Parse(ldr[0].ToString());
                            ldr.Close();
                            cmdunidad.CommandText = "insert into mgw10026 values (" + lidunidad + ",'" + x.cUnidad.ToUpper() + "','" + x.cUnidad.ToUpper() + "','" + x.cUnidad.ToUpper() + "','')";
                            cmdunidad.ExecuteNonQuery();
                            lret = fSetDatoProducto("CIDUNIDADBASE", lidunidad.ToString());
                        }
                    }
                    else
                    {
                        ldr.Read();

                        lidunidad = int.Parse(ldr[0].ToString());
                        ldr.Close();
                        lret = fSetDatoProducto("CIDUNIDADBASE", lidunidad.ToString());
                    }
                    //lret = fGuardaProducto();
                    lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
                }
                lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);
                if (lret != 0)
                {
                    fInsertaAlmacen();
                    lret = fSetDatoAlmacen("CCODIGOALMACEN", x.cCodigoAlmacen);
                    lret = fSetDatoAlmacen("CNOMBREALMACEN", x.cNombreAlmacen);
                    lret = fGuardaAlmacen();
                    lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);

                }
                decimal total;
                

                    lret = fSetDatoMovimiento("cUnidades", x.cUnidades.ToString());
                    lret = fSetDatoMovimiento("cPrecio", x.cPrecio.ToString());
                    lret = fSetDatoMovimiento("cporcentajeimpuesto1", x.cPorcent01.ToString());

                    try
                    {
                        lret = fSetDatoMovimiento("COBSERVACIONES", x.cObservaciones);
                        lret = fSetDatoMovimiento("ctextoextra1", x.ctextoextra1);
                        lret = fSetDatoMovimiento("cReferencia", x.cReferencia);
                        lret = fSetDatoMovimiento("ctextoextra2", x.ctextoextra2);
                        lret = fSetDatoMovimiento("ctextoextra3", x.ctextoextra3);
                    }
                    catch (Exception ee)
                    { }

                    // checar si hay existencias del producto para sacarlo;


                    if (tipo == 0)
                    {
                        string lAnio = DateTime.Today.Year.ToString();
                        string lMes = DateTime.Today.Month.ToString().PadLeft(2, '0');
                        string lDia = DateTime.Today.Day.ToString();

                        //lMes = "02";


                        decimal lExistencia = 0;
                        lExistencia = mRegresarExistencia(x.cCodigoProducto, x.cCodigoAlmacen.Trim(), lAnio, lMes, lDia, lIdDocumento);

                        if (lExistencia <= 0 || lExistencia < x.cUnidades)
                            mMandarBitacora1(x, lExistencia, false);
                        else
                        {
                            mMandarBitacora1(x, lExistencia, true);
                            lret = fGuardaMovimiento();
                            if (_RegDoctoOrigen.cCodigoConcepto == "1")
                            {
                                lret = fInsertarMovimiento();
                                lret = fSetDatoMovimiento("cIdDocumento", lIdDocumento.ToString());
                                lret = fSetDatoMovimiento("cNumeroMovimiento", lNumeroMov.ToString());
                                lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
                                lret = fSetDatoMovimiento("cCodigoAlmacen", x.cAlmacenEntrada);
                                lret = fSetDatoMovimiento("cUnidades", x.cUnidades.ToString());
                                lret = fSetDatoMovimiento("cPrecio", x.cPrecio.ToString());
                                lret = fGuardaMovimiento();
                            }
                            
                        }
                    }
                    if (tipo == 1)
                    {
                        lret = fGuardaMovimiento();
                        if (x.cObservaciones != "")
                        {
                            
                            string lcadenaA = "update mgw10010 set cobserva01= '" + x.cObservaciones + "' where ciddocum01 = " + lIdDocumento.ToString() + " and cnumerom01 = " + lNumeroMov.ToString();
                            OleDbCommand lsqlA = new OleDbCommand(lcadenaA, miconexion._conexion);
                            int xxx = 0;
                            //xxx = lsqlA.ExecuteNonQuery();
                
                        }
                            
                    }
                    lNumeroMov += 100;
                    
            }
            return "";
        }


        private void mMandarBitacora1(RegMovto x, decimal aExistencia, Boolean Exitoso)
        {
            string lcadena="";
            decimal lresta = aExistencia - x.cUnidades ;
            string ldia = DateTime.Today.Day.ToString() + "/" + DateTime.Today.Month.ToString() + "/" + DateTime.Today.Year.ToString();

            

            if (lvar.Count == 0)
            {
                lcadena = "FECHA DESCARGA,CODIGO,DESCRIPCION,ALMACEN DE DESCARGA,EXISTENCIA,CANTIDAD ORIGINAL,existencia final,STATUS";
                lvar.Add(lcadena);
            }
            if (Exitoso == true)
            {
                lcadena = ldia + "," + x.cCodigoProducto + "," + x.cNombreProducto.Trim() + "," + x.cCodigoAlmacen + "," + aExistencia.ToString() + "," + x.cUnidades.ToString() + "," + lresta + ",Se cargo con Exito";
            }
            else
            {
                lcadena = ldia + "," + x.cCodigoProducto + "," + x.cNombreProducto.Trim() + "," + x.cCodigoAlmacen + "," + aExistencia.ToString() + "," + x.cUnidades.ToString() + "," + lresta + ",No se cargo con Exito";
            }
            lvar.Add(lcadena);
        }

        private void mMandarBitacora(string aCodigoProducto, string aCodigoAlmacen)
        {
            lvar.Add("Producto " + aCodigoProducto + " No tienes existencias en el almacen " + aCodigoAlmacen);
        }

        private decimal mRegresarExistencia(string aCodigoProducto, string aCodigoAlmacen, string aAnio, string aMes, string aDia, long lIdDocumento)
        {

            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            decimal lidclien=0;
            //OleDbParameter lparametroIdDocumento = new OleDbParameter("@p1", _RegDoctoOrigen.cIdDocto);
            lsql.CommandText = "select m30.cidalmacen, m30.centrada01-m30.csalidas01 as ini, m30.centrada02-m30.csalidas02 as enero, m30.centrada03-m30.csalidas03 as febrero, m30.centrada04-m30.csalidas04 as marzo, m30.centrada05-m30.csalidas05 as abril, m30.centrada06-m30.csalidas06 as mayo, m30.centrada07-m30.csalidas07 as junio, m30.centrada08-m30.csalidas08 as julio, m30.centrada09-m30.csalidas09 as agosto, m30.centrada10-m30.csalidas10 as septiembre,m30.centrada11-m30.csalidas11 as octubre,m30.centrada12-m30.csalidas12 as noviembre,m30.centrada13-m30.csalidas13 as diciembre " +
            " from mgw10031 m31 join mgw10030 m30 on m30.cidejerc01 = m31.cidejerc01 "+
            " join mgw10005 m5 on m5.cidprodu01 = m30.cidprodu01 " +
            " join mgw10003 m3 on m3.cidalmacen = m30.cidalmacen "+
            " where cejercicio = " + aAnio +
            " and m5.ccodigop01 = '" + aCodigoProducto + "'" +
            " and m3.ccodigoa01 = '" + aCodigoAlmacen + "'";
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lreader.Read();
                switch (aMes)
                {
                    case "01": 
                        lidclien = long.Parse(lreader["ini"].ToString());
                        break;
                    case "02":
                        lidclien = long.Parse(lreader["enero"].ToString());
                        break;
                    case "03":
                        lidclien = long.Parse(lreader["febrero"].ToString());
                        break;
                    case "04":
                        lidclien = long.Parse(lreader["marzo"].ToString());
                        break;
                    case "05":
                        lidclien = long.Parse(lreader["abril"].ToString());
                        break;
                    case "06":
                        lidclien = long.Parse(lreader["mayo"].ToString());
                        break;
                    case "07":
                        lidclien = long.Parse(lreader["junio"].ToString());
                        break;
                    case "08":
                        lidclien = long.Parse(lreader["julio"].ToString());
                        break;
                    case "09":
                        lidclien = long.Parse(lreader["agosto"].ToString());
                        break;
                    case "10":
                        lidclien = long.Parse(lreader["septiembre"].ToString());
                        break;

                }
                lreader.Close();
                lsql.CommandText = "select sum(m10.cunidades) as uni, cafectae01 from mgw10010 m10 join mgw10005 m5 on m10.cidprodu01 = m5.cidprodu01 " +
                " join mgw10003 m3 on m3.cidalmacen = m10.cidalmacen " +
                " where m5.ccodigop01 = '" + aCodigoProducto + "'" +
                " and m3.ccodigoa01 = '" + aCodigoAlmacen + "'" +
                " and cafectae01 < 3 " +
                " and cafectad01 = 1 " +
                " and substr(dtos(m10.cfecha),5,2) = '" + aMes + "'" + 
                " and substr(dtos(m10.cfecha),1,4) = '" + aAnio + "'" +
                " group by cafectae01" +
                " order by m10.cafectae01"; 
                
                lreader = lsql.ExecuteReader();
                
                if (lreader.HasRows)
                {
                    // puede haber 2 registros uno para entradas y uno para salidas
                    lreader.Read();
                    string lafectaentrada = lreader["cafectae01"].ToString();
                    if (lafectaentrada == "1")
                    {
                        lidclien += decimal.Parse(lreader["uni"].ToString());
                    }
                    else
                        lidclien -= decimal.Parse(lreader["uni"].ToString());

                    if (lreader.Read() == false)
                    {
                        //lreader.Read();
                        if (lreader.HasRows)
                            try
                            {
                                lidclien -= decimal.Parse(lreader["uni"].ToString());
                            }
                            catch (Exception eeee)
                            { }
                    }

                }
                



  
            }
            else
                lidclien = 0;
            lreader.Close();
            lsql.CommandText = "select sum(m10.cunidades) as uni from mgw10010 m10 join mgw10005 m5 on m10.cidprodu01 = m5.cidprodu01 " +
                " join mgw10003 m3 on m3.cidalmacen = m10.cidalmacen " +
                " where m5.ccodigop01 = '" + aCodigoProducto + "'" +
                " and m3.ccodigoa01 = '" + aCodigoAlmacen + "'" +
                " and m10.cidprodu01 = m5.cidprodu01 and m10.cidalmacen = m3.cidalmacen " +
                " and m10.ciddocum01 = " + lIdDocumento; 
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lreader.Read();
                try
                {
                    lidclien -= decimal.Parse(lreader["uni"].ToString());
                }
                catch (Exception aaaaa)
                { }
            }
            lreader.Close();

            return lidclien;
        }

        protected virtual void  mLeerDireccion()
        {


            //_RegDoctoOrigen._RegDireccion; 
        }

        private Boolean mGrabarInterfaz(string aFolio, int aTipo)
        {
            OleDbConnection lconexion = new OleDbConnection();
            string amensaje = "";
            lconexion = miconexion.mAbrirRutaGlobal(out amensaje );
            bool lrespuesta = false;

            //lconexion = miconexion.mAbrirConexionDestino();
            try
            {
                OleDbCommand lsql = new OleDbCommand("insert into interfaz values ('" + aFolio + "'," + aTipo + ")", lconexion);
                lsql.ExecuteNonQuery();
            }
            catch (Exception eeee)
            { }

            miconexion.mCerrarConexionGlobal();






            return lrespuesta;
        }

        private bool mDarAltaCliente()
        {

            OleDbCommand  lsql = new OleDbCommand();
            OleDbDataReader   lreader;
            long lidclien;
            //OleDbParameter lparametroIdDocumento = new OleDbParameter("@p1", _RegDoctoOrigen.cIdDocto);
            lsql.CommandText = "select max(cidclien01) + 1 as cidclien01 from mgw10002";
            lsql.Connection = miconexion._conexion  ;
            lreader = lsql.ExecuteReader();
            _RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                lreader.Read();
                lidclien = long.Parse (lreader["cidclien01"].ToString ());
            }
            else
                lidclien = 1;
            lreader.Close();


            //OleDbConnection lconexion = new OleDbConnection();
           // lconexion = miconexion.mAbrirConexionDestino ();
            bool lrespuesta = false;
            string lfecha = _RegDoctoOrigen.cFecha.ToString();
            DateTime ldate = DateTime.Parse (lfecha);
            lfecha = ldate.ToString("dd/MM/yyyy");

            //lconexion = miconexion.mAbrirConexionDestino();
            string lcadena = "insert into mgw10002 (cidclien01, ccodigoc01,crazonso01,cfechaalta,crfc,cidmoneda, clistapr01, ctipocli01,cestatus) values (" +
                lidclien +
                ",'" + _RegDoctoOrigen.cCodigoCliente + "','" + _RegDoctoOrigen.cRazonSocial  + "'," +
                "ctod('" + lfecha + "'),'" +
                _RegDoctoOrigen.cRFC + "'" + 
                ",1,1,1,1)";
            OleDbCommand lsql1 = new OleDbCommand(lcadena, miconexion._conexion );
            try
            {
                lsql.CommandText = "SET NULL OFF";
                lsql.ExecuteNonQuery();

                lsql1.ExecuteNonQuery();
                lrespuesta = true;
            }
            catch (Exception eee)
            {
                lrespuesta = true;
            }

            //.mCerrarConexionDestino ();



            return lrespuesta;

        }



        protected virtual  bool mActualizaDocumento(long liddocum, int aopcion, double afolionuevo)
        {
            //miconexion.mAbrirConexionDestino();
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            long cidfoldig;
            long cidconce;
            long ciddocum01 = 0;
            string cserie = "";
            double ctotal=0;
            bool lrespuesta = false;
            
            //OleDbParameter lparametroIdDocumento = new OleDbParameter("@p1", _RegDoctoOrigen.cIdDocto);
            string lcadena = "update mgw10008 set cescfd = 1 where ciddocum01 = " + liddocum ;
            
            OleDbCommand lsql1 = new OleDbCommand(lcadena, miconexion._conexion);
            try
            {

                lsql1.ExecuteNonQuery();
                string lfecha = _RegDoctoOrigen.cFecha.ToString();
                DateTime ldate = DateTime.Parse(lfecha);
                lfecha = ldate.ToString("MM/dd/yyyy");


                long ctipocfd = 0;
                //string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
                //"ConceptoDocumento"
                    string lCodigoConcepto = _RegDoctoOrigen.cCodigoConcepto;
                lsql.CommandText = "select cidconce01, ciddocum01,cseriepo01, cverfacele from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'" ;
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    cidconce = long.Parse(lreader["cidconce01"].ToString());
                    ciddocum01 = long.Parse(lreader["ciddocum01"].ToString());
                    cserie = lreader["cseriepo01"].ToString();
                    ctipocfd = long.Parse (lreader["cverfacele"].ToString());

                }
                else
                    cidconce = 1;
                lreader.Close();

                lsql.CommandText = "select ctotal from mgw10008 where ciddocum01 = "+ liddocum ;
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    ctotal = double.Parse(lreader["ctotal"].ToString());

                }
                else
                    cidconce = 1;
                lreader.Close();


                double x = double.Parse(afolionuevo.ToString().Trim());

                lsql.CommandText = "select max(cidfoldig) + 1 as cidclien01 from mgw10045";
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    try
                    {
                        cidfoldig = long.Parse(lreader["cidclien01"].ToString());
                    }
                    catch (Exception ww)
                    {
                        cidfoldig = 1;
                    }
                }
                else
                    cidfoldig = 1;
                lreader.Close();

                lcadena = "insert into mgw10045 (cidfoldig,ciddoctode,cidcptodoc,ciddocto,cserie,cfolio,cestado,centregado, cfechaemi,cestrad,ctotal) " +
                                 " values (" + liddocum  + "," + ciddocum01 + "," + cidconce + "," + liddocum + ",'" + cserie.Trim()  + "'," + x + ",1, 0, ctod('" + lfecha + "'),3," + ctotal + ")";
                //lcadena = "insert into mgw10045 (cidfoldig,ciddoctode,cidcptodoc,ciddocto,cserie,cfolio,cestado, cfechaemi,cestrad) " +
                //                 "values (8,4,3001,11,'B',444,1,ctod('" + lfecha + "'),3)";
                OleDbCommand lsql2 = new OleDbCommand(lcadena, miconexion._conexion);
                lsql1.CommandText = "SET NULL OFF";
                lsql1.ExecuteNonQuery();

                lsql2.ExecuteNonQuery();
                lrespuesta = true;
            }
            catch (Exception eee)
            {
                lrespuesta = true;
            }
            finally
            {
                //miconexion.mCerrarConexionDestino();
            }
            //.mCerrarConexionDestino ();



            return lrespuesta;

        }
        private bool mActualizaDocumento2(long liddocum, long adestino, double afolio)
        {
            //if (adestino > 0)
            //    miconexion.mAbrirConexionOrigen(1);
            //else
                miconexion.mAbrirConexionDestino();
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            long cidfoldig;
            long cidconce;
            long ciddocum01 = 0;
            string cserie = "";
            double ctotal = 0;
            bool lrespuesta = false;

            int cescfd = 0;
            string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            lsql.CommandText = "select cescfd from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                lreader.Read();
                cescfd = int.Parse(lreader["cescfd"].ToString());
            }

            lreader.Close();
            if (cescfd == 0)
                return true;


            //OleDbParameter lparametroIdDocumento = new OleDbParameter("@p1", _RegDoctoOrigen.cIdDocto);
            string lcadena = "update mgw10008 set cescfd = 1 where ciddocum01 = " + liddocum;

            OleDbCommand lsql1 = new OleDbCommand(lcadena, miconexion._conexion);
            try
            {

                lsql1.ExecuteNonQuery();
                string lfecha = _RegDoctoOrigen.cFecha.ToString();
                DateTime ldate = DateTime.Parse(lfecha);
                lfecha = ldate.ToString("MM/dd/yyyy");




                //string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
                lsql.CommandText = "select cidconce01, ciddocum01,cseriepo01,cescfd from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    cidconce = long.Parse(lreader["cidconce01"].ToString());
                    ciddocum01 = long.Parse(lreader["ciddocum01"].ToString());
                    cserie = lreader["cseriepo01"].ToString();

                }
                else
                    cidconce = 1;
                lreader.Close();
                cserie = GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim();

                lsql.CommandText = "select ctotal from mgw10008 where ciddocum01 = " + liddocum;
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    ctotal = double.Parse(lreader["ctotal"].ToString());

                }
                else
                    cidconce = 1;
                lreader.Close();


                //double x = double.Parse(GetSettingValueFromAppConfigForDLL("FolioFactura").ToString().Trim());
                double x = afolio;

                lsql.CommandText = "select min(cidfoldig) as cidfoldig, min(cfolio) as cfolio, min(cserie) as cserie from mgw10045 where ciddocto = 0 and cidcptodoc = " + cidconce + " group by cidcptodoc";
                lsql.CommandText = "select top 1 cidfoldig, cfolio, cserie from mgw10045 where ciddocto = 0 and cidcptodoc = " + cidconce+ " order by cidfoldig asc";

                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                //string cserie;
                if (lreader.HasRows)
                {
                    lreader.Read();
                    cidfoldig = long.Parse(lreader["cidfoldig"].ToString());
                    x = double.Parse(lreader["cfolio"].ToString());
                    cserie = lreader["cserie"].ToString();
                }
                else
                {
                    cidfoldig = 1;
                    x = 1;

                }

                //return false;
                lreader.Close();

                //lcadena = "insert into mgw10045 (cidfoldig,ciddoctode,cidcptodoc,ciddocto,cserie,cfolio,cestado,centregado, cfechaemi,cestrad,ctotal) " +
                //                 " values (" + liddocum + "," + ciddocum01 + "," + cidconce + "," + liddocum + ",'" + cserie.Trim() + "'," + x + ",1, 0, ctod('" + lfecha + "'),3," + ctotal + ")";

                try
                {
                    lcadena = "update mgw10045 set ciddocto=" + liddocum + ",cestado=1,cfechaemi=ctod('" + lfecha + "')" +
                    " where cidfoldig = " + cidfoldig + " and ciddocto = 0 ";

                    //lcadena = "update mgw10045 set ciddocto=" + liddocum + ",cestado=1,cfechaemi=ctod('" + lfecha + "')" +
                    //" where cidfoldig  in (select min(cidfoldig) from mgw10045 where ciddocto = 0 and cidcptodoc = " + cidconce + " group by cidcptodoc)";


                    OleDbCommand lsql2 = new OleDbCommand(lcadena, miconexion._conexion);
                    lsql1.CommandText = "SET NULL OFF";
                    lsql1.ExecuteNonQuery();

                    long lcuantos = lsql2.ExecuteNonQuery();
                    if (lcuantos == 0)
                    {

                        lsql.CommandText = "select min(cidfoldig) as cidfoldig, min(cfolio) as cfolio, min(cserie) as cserie from mgw10045 where ciddocto = 0 and cidcptodoc = " + cidconce + " group by cidcptodoc";
                        lsql.Connection = miconexion._conexion;
                        lreader = lsql.ExecuteReader();
                        _RegDoctoOrigen._RegMovtos.Clear();
                        //string cserie;
                        if (lreader.HasRows)
                        {
                            lreader.Read();
                            cidfoldig = long.Parse(lreader["cidfoldig"].ToString());
                            x = double.Parse(lreader["cfolio"].ToString());
                            cserie = lreader["cserie"].ToString();
                        }
                        else
                        {
                            cidfoldig = 1;
                            x = 1;

                        }


                        lcadena = "update mgw10045 set ciddocto=" + liddocum + ",cestado=1,cfechaemi=ctod('" + lfecha + "')" +
                        " where cidfoldig = " + cidfoldig;
                        lsql2.ExecuteNonQuery();
                    }


                    lcadena = "update mgw10008 set cfolio=" + x + ",cseriedo01='" + cserie + "'" +
                    " where ciddocum01 = " + liddocum;
                    lsql2.CommandText = lcadena;
                    lsql2.ExecuteNonQuery();
                    lrespuesta = true;
                }
                catch (Exception eeeee)
                {
                    OleDbCommand lsql3 = new OleDbCommand(lcadena, miconexion._conexion);
                    lcadena = "delete from  mgw10008 " +
                    " where ciddocum01 = " + liddocum;
                    lsql3.CommandText = lcadena;
                    lsql3.ExecuteNonQuery();

                    lcadena = "delete from  mgw10010 " +
                    " where ciddocum01 = " + liddocum;
                    lsql3.CommandText = lcadena;
                    lsql3.ExecuteNonQuery();
                    lrespuesta = false;
                }









            }
            catch (Exception eee)
            {
                lrespuesta = true;
            }
            finally
            {
                if (adestino > 0)
                    miconexion.mCerrarConexionOrigen(1);
                else
                    miconexion.mCerrarConexionDestino();

            }
            //.mCerrarConexionDestino ();



            return lrespuesta;

        }

        public virtual string mBuscarDoctosArchivo(string aNombreArchivo)
        {
            
            return "";
        }




        #region ISujeto Members

        public void Registrar(IObservador obs)
        {
            //throw new NotImplementedException();
            lista.Clear();
            lista.Add(obs);
            Notificar(0);
        }

        public void Notificar()
        {
            //throw new NotImplementedException();
            foreach (IObservador x in lista)
                x.Actualizar(0);
        }

        public void Notificar(double lavance)
        {
            //throw new NotImplementedException();
            foreach (IObservador x in lista)
                x.Actualizar(lavance);
        }
        public void Notificar(string error)
        {
            //throw new NotImplementedException();
            foreach (IObservador x in lista)
                x.Actualizar(double.Parse(error));
        }

        #endregion


        public void Registrar()
        {
            throw new NotImplementedException();
        }

        public void mTraerInformacionPrimerReporte(ref DataSet aDS, DateTime aFechaInicial, DateTime aFechaFinal)
        {

            OleDbConnection lconexion = new OleDbConnection();
            miconexion.mAbrirConexionDestino();
            lconexion = miconexion._conexion;
            DateTime lfechai = aFechaInicial;
            string sfechai = lfechai.Year.ToString() + lfechai.Month.ToString().PadLeft(2, '0').ToString() + lfechai.Day.ToString().PadLeft(2, '0').ToString();
            DateTime lfechaf = aFechaFinal;
            string sfechaf = lfechaf.Year.ToString() + lfechaf.Month.ToString().PadLeft(2, '0').ToString() + lfechaf.Day.ToString().PadLeft(2, '0').ToString();


            string lquery = "select nc.cserie, nc.cfolio, m8.cfecha, m8.ciddocum02, m2.ccodigoc01, m2.crazonso01,m5.ccodigop01,m5.cnombrep01, nc.descuento, nc.descaplica, nc.cimpuesto, nc.cimpuesto2, m8.ctotal,'Estado' as cestado from mgw10005 m5 join ncprod nc " +
" on m5.cidprodu01 = nc.cidprodu01 " +
" join mgw10008 m8 on m8.cseriedo01 = nc.cserie and m8.cfolio = nc.cfolio and m8.ciddocum02 = 7 " +
" join mgw10002 m2 on m2.cidclien01 = m8.cidclien01 " +
           " and dtos(m8.cfecha) >= '" + sfechai + "' " +
                " and dtos(m8.cfecha) <= '"   + sfechaf +
                "' order by nc.cfolio, m8.cfecha ";


            DataSet ds = new DataSet();
            DataTable dt;


            DataTable dt1 = ds.Tables.Add("set1");
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(lquery, lconexion))
            {
                adapter.Fill(dt1);
            }

            aDS = ds;

        }

        private double mRegresaFolioComercial(string concepto)
        //        public string mGrabarDoctosComercial(List<RegDocto> Doctos)
        {
            double aFolio=0;
            string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            miconexion.mAbrirConexionComercial(true);
            SqlCommand m = new SqlCommand();

            //return 1011;
            m.CommandText = "select m6.ctipofolio, m6.cnofolio, m7.cnofolio from admconceptos m6 join admdocumentosmodelo m7 " +
            " on m6.ciddocumentode = m7.ciddocumentode " +
             " where ccodigoconcepto = '" + concepto + "'";
            m.Connection = miconexion._conexion1;

            SqlDataReader rd;
            rd = m.ExecuteReader();

            if (rd.HasRows)
            {
                rd.Read();
                try
                {
                    aFolio = double.Parse(rd[0].ToString());
                    if (aFolio == 3)
                        aFolio = double.Parse(rd[1].ToString());
                    else
                        aFolio = double.Parse(rd[2].ToString());
                }
                catch (Exception ee)
                {
                    //                    lreader.Close();
                }
                rd.Close();
            }
            
            

            miconexion.mCerrarConexionOrigenComercial();
            return aFolio;
            //return "";

        }


        private string mEjecutaQuery(string query)
        //        public string mGrabarDoctosComercial(List<RegDocto> Doctos)
        {
            
            string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            miconexion.mAbrirConexionComercial(true);
            SqlCommand m = new SqlCommand();

            //return 1011;
            m.CommandText = query;
            m.Connection = miconexion._conexion1;
            int lresult = m.ExecuteNonQuery();
            miconexion.mCerrarConexionOrigenComercial();
            //return aFolio;
            return "";

        }

        
        private double mRegresaFolioComercialuno(string concepto)
        {
            string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            miconexion.mAbrirConexionComercial(true);
            
            string aSerie = "";
            double aFolio = 0;

            StringBuilder sMensaje1 = new StringBuilder(512);
            int lResultado2 = fSetNombrePAQ("CONTPAQ I Comercial");
            if (lResultado2 != 0)
            {
                fErrorComercial(lResultado2, sMensaje1, 512);
                // MessageBox.Show("Error: " + sMensaje);
            }
            //fSiguienteFolioComercial(concepto, ref  aSerie, ref  aFolio);
            fAbreEmpresa(rutadestino);
            fSiguienteFolioComercial(concepto, ref  aSerie, ref  aFolio);
            fCierraEmpresa();

            miconexion.mCerrarConexionOrigenComercial();
            fTerminaSDK();
            return aFolio;
            //return "";

        }

        private double mGrabarCapasComercial(RegDocto areg, double folio)
        {
            miconexion.mAbrirConexionComercial(true);
            SqlCommand m = new SqlCommand();

            //return 1011;
            m.CommandText = "select m10.cidmovimiento from admmovimientos m10 join admdocumentos m8 " +
            " on m8.ciddocumento = m10.ciddocumento " +
             " where m8.cfolio = '" + folio + "'";
            m.Connection = miconexion._conexion1;

            List<long> ints = new List<long>();

            SqlDataReader rd;
            rd = m.ExecuteReader();

            if (rd.HasRows)
            {
                rd.Read();
                try
                {
                    ints.Add(long.Parse(rd[0].ToString()));
                }
                catch (Exception ee)
                {
                    //                    lreader.Close();
                }
                rd.Close();
            }
            
            

            miconexion.mCerrarConexionOrigenComercial();
            
            string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            miconexion.mAbrirConexionComercial(true);

            string aSerie = "";
            double aFolio = 0;

            StringBuilder sMensaje1 = new StringBuilder(512);
            int lResultado2 = fSetNombrePAQ("CONTPAQ I Comercial");
            if (lResultado2 != 0)
            {
                fErrorComercial(lResultado2, sMensaje1, 512);
                // MessageBox.Show("Error: " + sMensaje);
            }
            //fSiguienteFolioComercial(concepto, ref  aSerie, ref  aFolio);
            fAbreEmpresa(rutadestino);

            int i = 0;
            foreach (RegMovto movto in areg._RegMovtos)
            {
                movto.cIdMovto = ints[i++];
            } 

            foreach (RegMovto movto in areg._RegMovtos)
            {
                string lfecha = movto._RegCapa.FechaFabricacion.ToShortDateString();
                fAltaMovimientoSeriesCapas_ParamComercial(movto.cIdMovto.ToString().Trim(), movto.cUnidades.ToString().Trim(), "1", "", movto._RegCapa.Pedimento, movto._RegCapa.NoAduana.ToString().Trim(), lfecha , "", "", "");
            } 
            fCierraEmpresa();

            miconexion.mCerrarConexionOrigenComercial();
            fTerminaSDK();
            return aFolio;
            //return "";

        }


        
        /*
        public string mGrabarDoctosComercial44(List<RegDocto> Doctos,string usu, string pwd)
        {
            mAbrirSDK();

            double folio = 0;
            string serie ="";
            string kTokenSeparadorSACICOM = "¬";
            string lstrREsultado = "";
            int lResultado = 0;
            folio = mRegresaFolioComercial(Doctos[0].cCodigoConcepto);
            //folio = mRegresaFolioComercial(Doctos[0].cCodigoConcepto);
            folio = folio + 1;
            //folio = 1010;

            string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");

            
            string lruta2 = Directory.GetCurrentDirectory();

            CONTPAQiComercial.Comercial comComercialMain = new CONTPAQiComercial.Comercial();
            CONTPAQiComercial.TTInterfazTabla gTablas = new CONTPAQiComercial.TTInterfazTabla();
            string sURLSACI = "http://127.0.0.1:9080/saci/adminpaq";

            string szRegKeyAdminPAQ2001 = lruta2;
            string lRuta = lruta2 + @"\ContPAQiComercial.exe";

            string gGuidCom = "";
            string lConfig = "";
            int lError1 = gTablas.InicializarComunicacion(sURLSACI, "CONTPAQ I Comercial", 10, "", out gGuidCom, out lConfig);

            //return gGuidCom;
            //string kTokenSeparadorSACICOM = "¬";
            lstrREsultado = "";
            lResultado = 0;

            comComercialMain.ProcesarUnaFuncion("SetUrl", sURLSACI + kTokenSeparadorSACICOM + gGuidCom, "", out lstrREsultado, out lResultado);
            if (lstrREsultado != null)
                return lstrREsultado;
            else
                lstrREsultado = "";
            if (lstrREsultado != "")
                return lstrREsultado;


            
            string sUsuario = usu;
            string sPassword = pwd;
            int aIdUsuario = 0;
            string aNombreUsuario = "";
            int aPerfilUsuario = 0;
            string aListaEstados = "";
            string aListaPersmisos = "";
            string aListaDescripciones = "";
            int lResultado1 = 0;
            try
            {
                comComercialMain.seguridadValidaUsuario(1, sUsuario, sPassword, out aIdUsuario, out aNombreUsuario, out aPerfilUsuario, out aListaPersmisos, out aListaEstados, out aListaDescripciones, out lResultado1);
            }
            catch (Exception eee)
            {
                return "Error seguridad cheque datos de usuario";
            }
            if (lResultado1 != 0)
                return "Error seguridad" + lResultado1.ToString();

            string funcion = "Inserta_Documento";

            //string kTokenSeparadorSACICOM = "¬";
            //string aParametros1 = "1¬1¬1¬01/01/2014¬Cotización¬¬5¬UNO¬¬01/01/2014¬01/01/2014¬1¬1¬P1¬1¬1¬100¬0¬0¬0¬0";
            string aParametros = "";
            CultureInfo culture = new CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentCulture = culture;

            string lfecha =  Doctos[0].cFecha.ToShortDateString();
            string lNombreConcepto = "Compra";
            lNombreConcepto = Doctos[0]._cNombreConcepto;

            string primeraparte = "";
            string segundaparte = "";
            aParametros = "2" + kTokenSeparadorSACICOM;   //Ventas
            aParametros += Doctos[0]._RegMovtos.Count.ToString() + kTokenSeparadorSACICOM; // numero movimientos.
            aParametros += Doctos[0].cCodigoConcepto + kTokenSeparadorSACICOM; // id concepto
            aParametros += lfecha + kTokenSeparadorSACICOM; //fehca documento
            aParametros += lNombreConcepto + kTokenSeparadorSACICOM; //nombre concepto
            aParametros += "" + kTokenSeparadorSACICOM; //serie
            primeraparte = aParametros;
            aParametros += folio.ToString() +kTokenSeparadorSACICOM; //folio
            segundaparte = "";

            mValidaClienteProveedor(Doctos[0]);
            segundaparte += Doctos[0].cCodigoCliente + kTokenSeparadorSACICOM; //Codigo Cliente
            //segundaparte += "(Ninguno)" + kTokenSeparadorSACICOM; //Codigo Agente
            segundaparte += "" + kTokenSeparadorSACICOM; //Codigo Agente
            segundaparte += lfecha + kTokenSeparadorSACICOM; //Fecha vencimiento
            segundaparte += lfecha + kTokenSeparadorSACICOM; //Fecha entrega
            segundaparte += "1" + kTokenSeparadorSACICOM; //Moneda
            segundaparte += "1" + kTokenSeparadorSACICOM; //tipo de cambios
            
            foreach (RegMovto movto in Doctos[0]._RegMovtos)
            {
                mValidaProducto(movto);

                segundaparte += movto.cCodigoProducto + kTokenSeparadorSACICOM; //producto
                segundaparte += "1" + kTokenSeparadorSACICOM; //almacen
                segundaparte += movto.cUnidades + kTokenSeparadorSACICOM; //cantidad
                segundaparte += movto.cPrecio + kTokenSeparadorSACICOM; //PRECIO
                segundaparte += "0" + kTokenSeparadorSACICOM; //descuento 1
                segundaparte += "0" + kTokenSeparadorSACICOM; //descuento 2
                segundaparte += movto.cImpuesto + kTokenSeparadorSACICOM; //impuestos 1
                segundaparte += "0" + kTokenSeparadorSACICOM; //impuesto 2
                //updates += " update admmovimientos set referencia = '" + movto.cReferencia + "' where " 
            }
            

            // factura 1¬1¬4¬01/01/2014¬Factura al Contado¬¬55¬UNO¬¬01/01/2014¬01/01/2014¬1¬1¬P1¬1¬1¬100¬0¬0¬0¬0
            // 2¬1¬17¬01/01/2014¬Orden de Compra¬¬55¬P1¬¬01/01/2014¬01/01/2014¬1¬1¬P1¬1¬1¬100¬0¬0¬0¬0

            // tipo 1 ventas 2 compras 3 inventarios
            // numero de movimentos
            // id concepto
            // fecha 
            // Nombre de concepto
            //Serie
            //Folio
            // Codigo cliente/prov
            // Fecha de vencimiento
            // Fecha de entrega
            // Moneda
            // Tipo de cambioo
            // codigo producto
            // codigo almacen
            // cantidad
            // precio
            // descuento1 
            // descuento2
            //impuesto1
            // impuesto2

            //string lRutaEmpresa = @"C:\Compac\Empresas\adtest";
            string lResultado3 = "";
            int lError = 0;
            lResultado3 = "";
            aParametros = primeraparte + folio.ToString() + kTokenSeparadorSACICOM + segundaparte;


            /*
             ? aParametros
1¬1¬4¬12/02/2016¬Factura Crédito¬¬12¬UNO¬¬12/02/2016¬12/02/2016¬1¬1¬UNO¬1¬10¬10¬0¬0¬0¬0
? lrutaempresa
C:\Compac\Empresas\adTEST2
? Funcion
Inserta_Documento
             
            comComercialMain.ProcesarUnaFuncion(funcion, aParametros, @rutadestino, out lResultado3, out lError);
            int xxxx=0;
            if (lResultado3 != "")
            {
                while (lResultado3 == "El documento ya existe.")
                {
                    folio++;
                    aParametros = primeraparte + folio.ToString() + kTokenSeparadorSACICOM + segundaparte;
                    lError = 0;
                    lResultado3 = "";
                    comComercialMain.ProcesarUnaFuncion(funcion, aParametros, rutadestino, out lResultado3, out lError);
                }
                
                comComercialMain.sistemaFinalizar(out xxxx);
                gTablas.TerminarComunicacion();
                return lResultado3;
            }
            else
            {
               comComercialMain.sistemaFinalizar(out xxxx);
                gTablas.TerminarComunicacion();
                return "";
            }
            if (lResultado3 == "")
            {
                SqlCommand xcommand = new SqlCommand();
                string sql = "";
                int lmovto = 1;
                foreach (RegMovto movto in Doctos[0]._RegMovtos)
                {
                    string update = "update admMovimientos set cReferencia = '" + movto.cReferencia + "' from " +
                    " admMovimientos movims " +
                    " join " +
                    " ( " +
                    " select m8.ciddocumento,m10.cidmovimiento from admDocumentos m8 " +
                    " join admMovimientos m10 on m8.ciddocumento = m10.ciddocumento " +
                    " and cnumeromovimiento = " + lmovto.ToString() +
                    " where cfolio = " + folio.ToString() +
                    " ) as movto on movto.ciddocumento = movims.ciddocumento and movto.cidmovimiento = movims.cidmovimiento ";
                    
                    string lres = mEjecutaQuery(update);
                    lmovto++;

                }
            }
            comComercialMain.sistemaFinalizar(out xxxx);
            gTablas.TerminarComunicacion();

            mGrabarCapasComercial(Doctos[0],folio);
            return lResultado3;

            
            string updates = "";
            int lnummovto = 1;
            foreach (RegMovto movto in Doctos[0]._RegMovtos)
            {
                updates += " update admmovimientos set referencia = '" + movto.cReferencia + "' where ";
            }
            
            int lError2 = 0;
            mCerrarSDK();
            return "";

            //1¬1¬4¬02/16/2016¬Factura Crédito¬¬3¬UNO¬¬02/16/2016¬02/16/2016¬1¬1¬UNO¬1¬1¬0¬0¬0¬0¬0 bien
            //1¬1¬4¬2/16/2016¬Factura Crédito¬¬3¬UNO¬(Ninguno)¬2/16/2016¬2/16/2016¬1¬1¬UNO¬1¬20¬0¬0¬0¬0¬0¬

        }
*/

        public void mValidaClienteProveedor(RegDocto adocto)
        {
            StringBuilder aMensaje = new StringBuilder(512);
            int busca = fBuscaCteProvComercial(adocto.cCodigoCliente);
            if (busca != 0)
            {
                fInsertaCteProvComercial();
                busca = fSetDatoCteProvComercial("ccodigocliente", adocto._RegCliente.Codigo);
                if (busca != 0)
                {
                    fErrorComercial(busca, aMensaje, 512);
                    //MessageBox.show("Error: " + aMensaje);
                }
                busca = fSetDatoCteProvComercial("crazonsocial", adocto._RegCliente.RazonSocial);
                busca = fSetDatoCteProvComercial("cRFC", adocto.cRFC);
                busca = fSetDatoCteProvComercial("cTipoCliente", "2");
                busca = fSetDatoCteProvComercial("CLISTAPRECIOCLIENTE", "1");
                busca = fSetDatoCteProvComercial("CESTATUS", "1");
                busca = fSetDatoCteProvComercial("CIDMONEDA", "2");
                busca = fSetDatoCteProvComercial("CFECHAALTA", "10/31/2016");
                busca = fSetDatoCteProvComercial("CBANVENTACREDITO", "1");
                if (busca != 0)
                {
                    fErrorComercial(busca, aMensaje, 512);
                    //MessageBox.show("Error: " + aMensaje);
                }
                
                busca = fGuardaCteProvComercial();
                if (busca != 0)
                {
                    fErrorComercial(busca, aMensaje, 512);
                    //MessageBox.show("Error: " + aMensaje);
                }
            } 
        }


        public bool mValidaProducto(RegMovto amovto, ref string lidunidad, int ConCapas = 1, int sat33 = 0)
        {
            //string lidunidad="";

            SqlCommand m = new SqlCommand();

            //return 1011;
            m.CommandText = "SELECT CIDUNIDAD FROM admUnidadesMedidaPeso where CCLAVEINT ='" + amovto._RegProducto.CodigoMedidaPesoSAT + "'";
            m.Connection = miconexion._conexion1;

            SqlDataReader rd;
            rd = m.ExecuteReader();

            if (rd.HasRows)
            {
                rd.Read();
                lidunidad = rd[0].ToString();
            }
            rd.Close();
            
            StringBuilder aMensaje = new StringBuilder(512);
            int busca = fBuscaProductoComercial(amovto.cCodigoProducto.Trim());
            if (busca != 0)
            {
                fInsertaProductoComercial();
                busca = fSetDatoProductoComercial("ccodigoproducto", amovto.cCodigoProducto.Trim());
                if (busca != 0)
                {
                    fErrorComercial(busca, aMensaje, 512);
                    //MessageBox.show("Error: " + aMensaje);
                } 
                busca = fSetDatoProductoComercial("CNOMBREPRODUCTO", amovto._RegProducto.Nombre);
                if (ConCapas == 1)
                {
                    busca = fSetDatoProductoComercial("CCONTROLEXISTENCIA", "9");
                    busca = fSetDatoProductoComercial("CMETODOCOSTEO", "7");
                }
                else
                {
                    busca = fSetDatoProductoComercial("CCONTROLEXISTENCIA", "1");
                    busca = fSetDatoProductoComercial("CMETODOCOSTEO", "1");
                }
                busca = fSetDatoProductoComercial("CSTATUSPRODUCTO", "1");
                busca = fSetDatoProductoComercial("CIDUNIDADBASE", lidunidad);
                if (sat33 != 0)
                {
                    busca = fSetDatoProductoComercial("CCLAVESAT", amovto._RegProducto.noIdentificacion);
                }
                int lret = fGuardaProductoComercial();
                if (lret != 0)
                {
                    fErrorComercial(lret, aMensaje, 512);
                    //MessageBox.show("Error: " + aMensaje);
                } 
            }
            
            return true;
        }

        public void mAbrirSDK()
        {
            string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            miconexion.mAbrirConexionComercial(true);

            string aSerie = "";
            double aFolio = 0;

            StringBuilder sMensaje1 = new StringBuilder(512);
            int lResultado2 = fSetNombrePAQ("CONTPAQ I Comercial");
            if (lResultado2 != 0)
            {
                fErrorComercial(lResultado2, sMensaje1, 512);
                // MessageBox.Show("Error: " + sMensaje);
            }
            //fSiguienteFolioComercial(concepto, ref  aSerie, ref  aFolio);
            fAbreEmpresa(rutadestino);
        }

        public void mCerrarSDK()
        {
            miconexion.mCerrarConexionOrigenComercial();
            fCierraEmpresa();
        }

        public void liberarrecursos()
        {
            int x;
            //gTablas.Cerrar(0);
            //If Not (comComercialMain Is Nothing) Then

            //End If
            // comComercialMain.empresaCerrar(out x);
            // gTablas = null;
            //  comComercialMain = null;

        }

        private int mGrabaEncabezadoComercial(RegDocto doc, ref int aIdDocumento)
        {
            int lret2=0;
            int lerrordocto = 0;
            StringBuilder sMensaje1 = new StringBuilder(512);
            string aCodigoConcepto = "";
                string aSerie = "";
                double aFolio = 0;
                if (doc.cFolio == 0)
                {
                    try
                    {
                        fSiguienteFolioComercial(doc.cCodigoConcepto, ref  aSerie, ref  aFolio);
                    }
                    catch (Exception ii)
                    {
                    }
                }
                else
                    aFolio = doc.cFolio;



                if (aFolio == 0)
                    aFolio = 1;

                

                fInsertarDocumentoComercial();

                lret2 = fSetDatoDocumentoComercial("cCodigoConcepto", doc.cCodigoConcepto);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    //continue;
                }


                lret2 = fSetDatoDocumentoComercial("cCodigoCliente", doc.cCodigoCliente);

                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    mValidaClienteProveedor(doc);
                    lret2 = fSetDatoDocumentoComercial("cCodigoCliente", doc.cCodigoCliente);
                    if (lret2 != 0)
                    {
                        fErrorComercial(lret2, sMensaje1, 512);
                        fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                        //continue;
                    }
                }
                else
                {
                    int busca = fBuscaCteProvComercial(doc.cCodigoCliente);
                    if (busca == 0)
                    {
                        StringBuilder aValorRFC = new StringBuilder(20);
                        lret2 = fLeeDatoDocumentoComercial("CRFC", aValorRFC, 20);
                        doc.cRFC = aValorRFC.ToString();
                    }
                }
                lret2 = fSetDatoDocumentoComercial("cRazonSocial", doc._RegCliente.RazonSocial);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    //continue;
                }
                lret2 = fSetDatoDocumentoComercial("cRFC", doc.cRFC);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    //continue;
                }

                



                //lret2 = fSetDatoDocumentoComercial("cIdMoneda", "2");
                //if (lret2 != 0)
                //    fErrorComercial(lret2, sMensaje1, 512);

                //lret2 = fSetDatoDocumentoComercial("cTipoCambio", doc.cTipoCambio.ToString().Trim());
                //if (lret2 != 0)
                //    fErrorComercial(lret2, sMensaje1, 512);

                DateTime lFechaVencimiento = DateTime.Today;
                string lfechavenc = String.Format("{0:MM/dd/yyyy}", DateTime.Today);
                lfechavenc = String.Format("{0:MM/dd/yyyy}", doc.cFecha);
                lret2 = fSetDatoDocumentoComercial("cFecha", lfechavenc);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    return 0;
                }

                lret2 = fSetDatoDocumentoComercial("cFolio", aFolio.ToString());
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    return 0;
                }
                lret2 = fSetDatoDocumentoComercial("cFechaVencimiento", lfechavenc);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    return 0;
                }

                lret2 = fSetDatoDocumentoComercial("cCodigoCliente", doc.cCodigoCliente);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    return 0;
                }

                RegCliente lc = new RegCliente();
                
                lc = mBuscarClienteComercial(doc.cCodigoCliente);

                lret2 = fSetDatoDocumentoComercial("CMETODOPAG", lc.MetodoPago);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    return 0;
                }

                lret2 = fGuardaDocumentoComercial();
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    return 0;
                }

                StringBuilder aValor = new StringBuilder(12);
                lret2 = fLeeDatoDocumentoComercial("CIDDOCUMENTO", aValor, 12);
                int liddocumento = int.Parse(aValor.ToString());
                

                doc.cIdDocto = liddocumento;
                //if (incluyedireccion == 1)
                //    lret2 = mGrabaDireccionComercial(doc);
                lret2 = mgrabamoneda(liddocumento, doc.cMoneda, doc.cTipoCambio);

                return 1;
        }


        private int mGrabarMovimientosComercial(RegDocto doc, int indicedoc, ref decimal ltotaunidadesdocto)
        {
            int lret2 = 0;
            StringBuilder sMensaje1 = new StringBuilder(512);
            int lerrordocto = 0;
            //decimal ltotaunidadesdocto = 0;
            int lerrormovto = 0;
            int indicemov = 0;

            if (doc.cFolio == 7142)
                doc.cFolio = 7142;

            foreach (RegMovto movto in doc._RegMovtos)
            {
                if (lerrormovto != 0)
                    continue;
                fInsertarMovimientoComercial();
                string lidunidad = "";
                mValidaProducto(movto, ref lidunidad, 0, 1 );
                lret2 = fSetDatoMovimientoComercial("cIdDocumento", doc.cIdDocto.ToString().Trim());
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El producto " + movto.cCodigoProducto + " " + sMensaje1.ToString(), ref lerrormovto);
                    continue;
                }
                lret2 = fSetDatoMovimientoComercial("cCodigoProducto", movto.cCodigoProducto.Trim());
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El producto " + movto.cCodigoProducto + " " + sMensaje1.ToString(), ref lerrormovto);
                    continue;
                }
                lret2 = fSetDatoMovimientoComercial("cCodigoAlmacen", movto.cCodigoAlmacen);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El producto " + movto.cCodigoProducto + " " + sMensaje1.ToString(), ref lerrormovto);
                    continue;
                }
                //int lRet3 = fSetDatoMovimientoComercial("cUnidadesCapturadas", movto.cUnidades.ToString().Trim());

                lret2 = fSetDatoMovimientoComercial("CIDUNIDAD", lidunidad);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El producto " + movto.cCodigoProducto + " " + sMensaje1.ToString(), ref lerrormovto);
                    continue;
                }


                int lRet = 0;
                if (movto._RegCapa.Pedimento != "")
                {

                    lret2 = fGuardaMovimientoComercial();
                    if (lret2 != 0)
                    {
                        fErrorComercial(lret2, sMensaje1, 512);
                        // MessageBox.Show("Error: " + sMensaje);
                        lret2 = fGuardaMovimientoComercial();
                        if (lret2 != 0)
                        {
                            fProcesaError("El producto " + movto.cCodigoProducto + " " + sMensaje1.ToString(), ref lerrormovto);
                            continue;
                        }
                    }
                    StringBuilder aValor = new StringBuilder(12);
                    aValor.Length = 0;
                    lret2 = fLeeDatoMovimientoComercial("CIDMOVIMIENTO", aValor, 12);
                    int lidmovimiento = int.Parse(aValor.ToString());

                    doc._RegMovtos[indicemov].cIdMovto = lidmovimiento;
                    _RegDoctos[indicedoc]._RegMovtos[indicemov++].cIdMovto = lidmovimiento;

                    sMensaje1.Length = 0;
                    string lfechaped = movto._RegCapa.FechaFabricacion.ToShortDateString();
                    lRet = fBuscarIdMovimientoComercial((int)movto.cIdMovto);
                    StringBuilder sIdproducto = new StringBuilder(12);
                    sIdproducto.Length = 0;
                    fLeeDatoMovimientoComercial("cidproducto", sIdproducto, 12);
                    movto._RegProducto.Id = int.Parse(sIdproducto.ToString());


                    ltotaunidadesdocto += movto.cUnidades;

                    lRet = fAltaMovimientoSeriesCapas_ParamComercial(movto.cIdMovto.ToString().Trim(), movto.cUnidades.ToString().Trim(), movto._RegCapa.tc.ToString().Trim(), "", movto._RegCapa.Pedimento, movto._RegCapa.NoAduana.ToString(), lfechaped, "", "", "");
                    if (lRet != 0)
                    {
                        //lRet = fAltaMovimientoSeriesCapas_ParamComercial("42", movto.cUnidades.ToString().Trim(), "1", "", movto._RegCapa.Pedimento, movto._RegCapa.NoAduana.ToString(), lfechaped, "", "", "");
                        int lultimacapa = mRegresaUltimaCapa();
                        mGrabaCapasinSDK(movto, lultimacapa);
                        fErrorComercial(lRet, sMensaje1, 512);
                    }
                    else
                    {
                        mGrabarTCCapa(movto.cIdMovto, movto._RegCapa.tc);
                    }
                    lRet = fBuscarIdMovimientoComercial((int)movto.cIdMovto);
                    lRet = fEditarMovimientoComercial();
                }
                else
                {
                    lRet = fSetDatoMovimientoComercial("cUnidadesCapturadas", movto.cUnidades.ToString());
                    if (lRet != 0)
                    {
                        fErrorComercial(lRet, sMensaje1, 512);
                        // MessageBox.Show("Error: " + sMensaje);
                    }
                }


                //lRet = fEditarMovimientoComercial();
                //string cantidad = movto.cUnidades.ToString().Trim();
                //decimal dcantidad = decimal.Parse(cantidad);
                // cantidad = dcantidad.ToString("F");
                //cantidad = "10.00";
                //lRet = fSetDatoMovimientoComercial("cUnidades", cantidad);




                lRet = fSetDatoMovimientoComercial("cPrecioCapturado", movto.cPrecio.ToString().Trim());
                if (lRet != 0)
                {
                    fErrorComercial(lRet, sMensaje1, 512);
                    // MessageBox.Show("Error: " + sMensaje);
                }
                lRet = fSetDatoMovimientoComercial("cImpuesto1", movto.cImpuesto.ToString().Trim());
                if (lRet != 0)
                {
                    fErrorComercial(lRet, sMensaje1, 512);
                    // MessageBox.Show("Error: " + sMensaje);
                }
                lRet = fGuardaMovimientoComercial();
                if (lRet != 0)
                {
                    fErrorComercial(lRet, sMensaje1, 512);
                    lret2 = fGuardaMovimientoComercial();
                    if (lret2 != 0)
                    {
                        fErrorComercial(lRet, sMensaje1, 512);
                        fProcesaError("El producto " + movto.cCodigoProducto + " " + sMensaje1.ToString(), ref lerrormovto);
                        return 0;
                    }

                }

            }
            return 1;
        }

        public string mGrabarDoctosComercial(int incluyetimbrado = 1)
        {
            StringBuilder sMensaje1 = new StringBuilder(512);
            
            string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            string aNombreEmpresa = "0000000";
            string aDirectorioEmpresa = "0000000000";
            int aIdEmpresa = 0;


            if (sdkcomercial == false)
            {

                miconexion.mAbrirConexionComercial(true);    
                int lResultado = fSetNombrePAQ("CONTPAQ I Comercial");
                if (lResultado != 0)
                {
                    fErrorComercial(lResultado, sMensaje1, 512);
                }

                if (incluyetimbrado == 1)
                {
                    int lresp10 = fInicializaLicenseInfoComercial(0);
                    if (lresp10 != 0)
                    {
                        fErrorComercial(lresp10, sMensaje1, 512);
                    }
                }
                
                sdkcomercial = true;
            }

            fAbreEmpresa(rutadestino);


            
            


            int indicedoc = 0;
            int lret2;
            int lcuantos = _RegDoctos.Count;
            int ltotales = lcuantos;
            int lindice = 1;
            int liddocumento = 0;
            decimal ltotalunidadesdocto = 0;

            foreach (RegDocto doc in _RegDoctos)
            {

                int lRetorno = mGrabaEncabezadoComercial(doc, ref liddocumento);
                if (lRetorno == 1)
                {
                    lRetorno = mGrabarMovimientosComercial(doc, indicedoc, ref ltotalunidadesdocto);
                    indicedoc++;

                }
                
                

                if (lRetorno == 1)
                {
                    mGrabarUnidadesDocto(doc.cIdDocto, ltotalunidadesdocto);
                    Notificar((double)(lindice++ * 100) / lcuantos);

                    if (incluyetimbrado == 1)
                    { 
                        
                        string lpass = "";
                        lpass = GetSettingValueFromAppConfigForDLL("Pass").ToString().Trim();


                        int lresp20 = fEmitirDocumentoComercial(doc.cCodigoConcepto, doc.cSerie, doc.cFolio, lpass, "");
                        if (lresp20 != 0)
                        {
                            fErrorComercial(lresp20, sMensaje1, 512);
                            // MessageBox.Show("Error: " + sMensaje);

                            
                        }
                        else
                        {
                            lresp20 = fEntregEnDiscoXMLComercial(doc.cCodigoConcepto, doc.cSerie, doc.cFolio, 1, @"C:\Compac\Empresas\Reportes\Formatos Digitales\reportes_Servidor\COMERCIAL\Factura.rdl");
                            if (lresp20 != 0)
                            {
                                fErrorComercial(lresp20, sMensaje1, 512);
                                // MessageBox.Show("Error: " + sMensaje);


                            }
                        }
                    }


                    //lexitosos++;
                }
                /*else
                    fBorraDocumentoComercial();*/

            }












            return "";
            
        }

        public void mCerrarSdkComercial()
        {
            try
            {
                fCierraEmpresa();
                miconexion.mCerrarConexionOrigenComercial();
                fTerminaSDK();
            }
            catch (Exception eeeeee)
            { 
            }
            
        }



        public string mGrabarDoctosComercial(List<RegDocto> Doctos,  ref int lexitosos, ref int ltotales,int incluyedireccion = 1)
        {

            string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            miconexion.mAbrirConexionComercial(true);

            
            StringBuilder sMensaje1 = new StringBuilder(512);
            int lResultado = fSetNombrePAQ("CONTPAQ I Comercial");
            if (lResultado != 0)
            {
                fErrorComercial(lResultado, sMensaje1, 512);
            }



            string aNombreEmpresa = "0000000";
            string aDirectorioEmpresa = "0000000000";

            //StringBuilder aNombreEmpresa = new StringBuilder(30);
            //StringBuilder aDirectorioEmpresa = new StringBuilder(30);

            int aIdEmpresa = 0;
            //lResultado = fPosPrimerEmpresa(ref aIdEmpresa, ref aNombreEmpresa, ref aDirectorioEmpresa);

            //string lDirectorioEmpresa = @"C:\Compac\Empresas\adBIOS2";
            //string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            fAbreEmpresa(rutadestino);
            int indicedoc = 0;
            int lret2;
            int lcuantos = Doctos.Count;
            ltotales = lcuantos;
            int lindice = 1;
            
            foreach (RegDocto doc in Doctos)
            {
                int lerrordocto = 0;
                string aCodigoConcepto = "";
                string aSerie = "";
                double aFolio = 0;
                try
                {
                    fSiguienteFolioComercial(doc.cCodigoConcepto, ref  aSerie, ref  aFolio);
                }
                catch (Exception ii)
                {
                    if (aFolio != 0)
                        aFolio = aFolio;
                    else
                        aFolio = 0;
                }
                if (aFolio == 0)
                    aFolio = 1;

                fInsertarDocumentoComercial();
                lret2 = fSetDatoDocumentoComercial("cCodigoCliente", doc.cCodigoCliente);

                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    mValidaClienteProveedor(doc);
                    lret2 = fSetDatoDocumentoComercial("cCodigoCliente", doc.cCodigoCliente);
                    if (lret2 != 0)
                    {
                        fErrorComercial(lret2, sMensaje1, 512);
                        fProcesaError("El documento con cliente " + doc.cCodigoCliente + " "  + sMensaje1.ToString(), ref lerrordocto);
                        continue;
                    }
                }
                lret2 = fSetDatoDocumentoComercial("cRazonSocial", doc._RegCliente.RazonSocial);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    continue;
                }
                lret2 = fSetDatoDocumentoComercial("cRFC", doc.cRazonSocial);
                if (lret2 != 0)
                    {
                        fErrorComercial(lret2, sMensaje1, 512);
                        fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                        continue;
                    }

                lret2 = fSetDatoDocumentoComercial("cCodigoConcepto", doc.cCodigoConcepto);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    continue;
                }



                //lret2 = fSetDatoDocumentoComercial("cIdMoneda", "2");
                //if (lret2 != 0)
                //    fErrorComercial(lret2, sMensaje1, 512);

                //lret2 = fSetDatoDocumentoComercial("cTipoCambio", doc.cTipoCambio.ToString().Trim());
                //if (lret2 != 0)
                //    fErrorComercial(lret2, sMensaje1, 512);

                DateTime lFechaVencimiento = DateTime.Today;
                string lfechavenc = String.Format("{0:MM/dd/yyyy}", DateTime.Today);
                lfechavenc = String.Format("{0:MM/dd/yyyy}", doc.cFecha);
                lret2 = fSetDatoDocumentoComercial("cFecha", lfechavenc);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    continue;
                }

                lret2 = fSetDatoDocumentoComercial("cFolio", aFolio.ToString());
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    continue;
                }
                lret2 = fSetDatoDocumentoComercial("cFechaVencimiento", lfechavenc);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    continue;
                }

                lret2 = fSetDatoDocumentoComercial("cCodigoCliente", doc.cCodigoCliente);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    continue;
                }
                lret2 = fGuardaDocumentoComercial();
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError("El documento con cliente " + doc.cCodigoCliente + " " + sMensaje1.ToString(), ref lerrordocto);
                    continue;
                }

                StringBuilder aValor = new StringBuilder(12);
                lret2 = fLeeDatoDocumentoComercial("CIDDOCUMENTO", aValor, 12);
                int liddocumento = int.Parse(aValor.ToString());
                int indicemov = 0;

                doc.cIdDocto = liddocumento;
                if (incluyedireccion == 1)
                    lret2 = mGrabaDireccionComercial(doc);
                lret2 = mgrabamoneda(liddocumento, doc.cMoneda, doc.cTipoCambio);


                decimal ltotaunidadesdocto = 0;
                int lerrormovto = 0;
                foreach (RegMovto movto in doc._RegMovtos)
                {
                    if (lerrormovto != 0)
                        continue;
                    fInsertarMovimientoComercial();
                    string lidunidad = "";
                    mValidaProducto(movto,ref lidunidad);
                    lret2 = fSetDatoMovimientoComercial("cIdDocumento", liddocumento.ToString().Trim());
                    if (lret2 != 0)
                    {
                        fErrorComercial(lret2, sMensaje1, 512);
                        fProcesaError("El producto " + movto.cCodigoProducto + " " + sMensaje1.ToString(), ref lerrormovto);
                        continue;
                    }
                    lret2 = fSetDatoMovimientoComercial("cCodigoProducto", movto.cCodigoProducto);
                    if (lret2 != 0)
                    {
                        fErrorComercial(lret2, sMensaje1, 512);
                        fProcesaError("El producto " + movto.cCodigoProducto + " " + sMensaje1.ToString(), ref lerrormovto);
                        continue;
                    }
                    lret2 = fSetDatoMovimientoComercial("cCodigoAlmacen", movto.cCodigoAlmacen);
                    if (lret2 != 0)
                    {
                        fErrorComercial(lret2, sMensaje1, 512);
                        fProcesaError("El producto " + movto.cCodigoProducto + " " + sMensaje1.ToString(), ref lerrormovto);
                        continue;
                    }
                    //int lRet3 = fSetDatoMovimientoComercial("cUnidadesCapturadas", movto.cUnidades.ToString().Trim());

                    lret2 = fSetDatoMovimientoComercial("CIDUNIDAD", "1");
                    if (lret2 != 0)
                    {
                        fErrorComercial(lret2, sMensaje1, 512);
                        fProcesaError("El producto " + movto.cCodigoProducto + " " + sMensaje1.ToString(), ref lerrormovto);
                        continue;
                    }

                    lret2 = fGuardaMovimientoComercial();
                    if (lret2 != 0)
                    {
                         fErrorComercial(lret2, sMensaje1, 512);
                        // MessageBox.Show("Error: " + sMensaje);
                        lret2 = fGuardaMovimientoComercial();
                        if (lret2 != 0)
                        {
                            fProcesaError("El producto " + movto.cCodigoProducto + " " + sMensaje1.ToString(), ref lerrormovto);
                            continue;
                        }
                    }
                    aValor.Length = 0;
                    lret2 = fLeeDatoMovimientoComercial("CIDMOVIMIENTO", aValor, 12);

                    int lidmovimiento = int.Parse(aValor.ToString());
                    
                    doc._RegMovtos[indicemov].cIdMovto = lidmovimiento;
                    Doctos[indicedoc]._RegMovtos[indicemov++].cIdMovto = lidmovimiento;

                    sMensaje1.Length = 0;
                    string lfechaped = movto._RegCapa.FechaFabricacion.ToShortDateString();
                    int lRet = fBuscarIdMovimientoComercial((int)movto.cIdMovto);
                    StringBuilder sIdproducto = new StringBuilder(12) ;
                    sIdproducto.Length = 0;
                    fLeeDatoMovimientoComercial("cidproducto", sIdproducto, 12);
                    movto._RegProducto.Id = int.Parse(sIdproducto.ToString());


                    ltotaunidadesdocto += movto.cUnidades;
                    if (movto._RegCapa.Pedimento != "")
                    {
                        lRet = fAltaMovimientoSeriesCapas_ParamComercial(movto.cIdMovto.ToString().Trim(), movto.cUnidades.ToString().Trim(), movto._RegCapa.tc.ToString().Trim(), "", movto._RegCapa.Pedimento, movto._RegCapa.NoAduana.ToString(), lfechaped, "", "", "");
                        if (lRet != 0)
                        {
                            //lRet = fAltaMovimientoSeriesCapas_ParamComercial("42", movto.cUnidades.ToString().Trim(), "1", "", movto._RegCapa.Pedimento, movto._RegCapa.NoAduana.ToString(), lfechaped, "", "", "");
                            int lultimacapa = mRegresaUltimaCapa();
                            mGrabaCapasinSDK(movto, lultimacapa);
                            fErrorComercial(lRet, sMensaje1, 512);
                        }
                        else
                        {
                            mGrabarTCCapa(movto.cIdMovto, movto._RegCapa.tc);
                        }
                        lRet = fBuscarIdMovimientoComercial((int)movto.cIdMovto);
                        lRet = fEditarMovimientoComercial();
                    }

                    //lRet = fEditarMovimientoComercial();
                    //string cantidad = movto.cUnidades.ToString().Trim();
                    //decimal dcantidad = decimal.Parse(cantidad);
                   // cantidad = dcantidad.ToString("F");
                    //cantidad = "10.00";
                    //lRet = fSetDatoMovimientoComercial("cUnidades", cantidad);

                    //lRet = fSetDatoMovimientoComercial("cUnidadesCapturadas", cantidad);
                    if (lRet != 0)
                    {
                        fErrorComercial(lRet, sMensaje1, 512);
                        // MessageBox.Show("Error: " + sMensaje);
                    }


                    lRet = fSetDatoMovimientoComercial("cPrecioCapturado", movto.cPrecio.ToString().Trim());
                    if (lRet != 0)
                    {
                        fErrorComercial(lRet, sMensaje1, 512);
                        // MessageBox.Show("Error: " + sMensaje);
                    }
                    lRet = fSetDatoMovimientoComercial("cImpuesto1", movto.cImpuesto.ToString().Trim());
                    if (lRet != 0)
                    {
                        fErrorComercial(lRet, sMensaje1, 512);
                        // MessageBox.Show("Error: " + sMensaje);
                    }
                    lRet = fGuardaMovimientoComercial();
                    if (lRet != 0)
                    {
                        fErrorComercial(lRet, sMensaje1, 512);
                        fProcesaError("El producto" + movto.cCodigoProducto + sMensaje1.ToString(), ref lerrormovto);
                        continue;
                    }

                }
                if (lerrormovto == 0)
                {
                    mGrabarUnidadesDocto(doc.cIdDocto, ltotaunidadesdocto);
                    Notificar((double)(lindice++ * 100) / lcuantos);
                    lexitosos++;
                }
                else
                    fBorraDocumentoComercial();
                
            }
            
            
           

                

            
            




            
            fCierraEmpresa();

            miconexion.mCerrarConexionOrigenComercial();
            fTerminaSDK();
            return "";
        }
        private void fProcesaError(string error, ref int lerrormovto)
        {
            Notificar(error);
            lerrormovto = 1;
            
        }

        private void mDeshacer(long id)
        {
            
        }

        private int mGrabarUnidadesDocto(long aIdDocto, decimal ltotalunidaes)
        {
            SqlCommand lsql = new SqlCommand();
            lsql.CommandText = "  update admDocumentos set CTOTALUNIDADES = " + ltotalunidaes.ToString().Trim() + " , CUNIDADESPENDIENTES = " + ltotalunidaes.ToString().Trim() +
" WHERE CIDDOCUMENTO = " + aIdDocto.ToString().Trim();
            lsql.Connection = miconexion._conexion1;
            int lret = lsql.ExecuteNonQuery();
            return lret;

        }

        private int mGrabarTCCapa(long aIdMovto, decimal atc)
        {
            SqlCommand lsql = new SqlCommand();
            lsql.CommandText = "  update admCapasProducto set CTIPOCAMBIO = " + atc.ToString().Trim() +
" from admCapasProducto c join admMovimientosCapas mc " +
" on c.CIDCAPA = mc.CIDCAPA " +
" and mc.CIDMOVIMIENTO = " + aIdMovto.ToString().Trim();
                lsql.Connection = miconexion._conexion1;
                int lret = lsql.ExecuteNonQuery();
                return lret;

        }

        private int mgrabamoneda(int aiddocumento, string aMoneda, decimal aTC)
        {
            SqlCommand lsql = new SqlCommand();
            int lmoneda = 0;
            if ( aMoneda != "Peso Mexicano")
            {
                lsql.CommandText = "update admDocumentos set cidmoneda = 2 ,ctipocambio = " + aTC.ToString().Trim() + " where ciddocumento = " + aiddocumento.ToString().Trim();
                lsql.Connection = miconexion._conexion1;
                int lret = lsql.ExecuteNonQuery();
            }
            return 1;
        }

        private int mGrabaDireccionComercial(RegDocto doc)
        {
            fInsertaDireccionComercial();
            int lret = fSetDatoDireccionComercial("CIDCATALOGO",doc.cIdDocto.ToString().Trim());
            lret = fSetDatoDireccionComercial("CTIPOCATALOGO","3");
            lret = fSetDatoDireccionComercial("CTIPODIRECCION","1");
            lret = fSetDatoDireccionComercial("CNOMBRECALLE",doc._RegDireccion.cNombreCalle);
            lret = fSetDatoDireccionComercial("CNUMEROEXTERIOR", doc._RegDireccion.cNumeroExterior);
            lret = fSetDatoDireccionComercial("CCOLONIA", doc._RegDireccion.cColonia);
            lret = fSetDatoDireccionComercial("CCODIGOPOSTAL",doc._RegDireccion.cCodigoPostal);
            //lret = fSetDatoDireccionComercial("CTELEFONO1", );
            lret = fSetDatoDireccionComercial("CEMAIL", doc._RegDireccion.cEmail);
            lret = fSetDatoDireccionComercial("CPAIS","Mexico");
            lret = fSetDatoDireccionComercial("CCIUDAD", doc._RegDireccion.cCiudad);
            lret = fSetDatoDireccionComercial("CMUNICIPIO", doc._RegDireccion.cCiudad);



            lret = fGuardaDireccionComercial();
            return 1;
        }


        /*int lultimacapa = mRegresaUltimaCapa();
                        mGrabaCapasinSDK(movto);*/


        public void mGrabaCapasinSDK(RegMovto movto, int lultimacapa)
        {
            string lfechaped = movto._RegCapa.FechaFabricacion.Year.ToString() + movto._RegCapa.FechaFabricacion.Month.ToString().PadLeft(2, '0') + movto._RegCapa.FechaFabricacion.Day.ToString().PadLeft(2, '0');
            string lfechahoy = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0');
            string lfechatime = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss:000");
            string laduana = "";
            switch (movto._RegCapa.NoAduana)
                {
                    case "010": laduana = "ACAPULCO, GRO."; break;
                    case "020": laduana = "AGUA PRIETA, SON."; break;
                    case "050": laduana = "SUBTENIENTE LOPEZ, Q. ROO."; break;
                    case "060": laduana = "CD. DEL CARMEN, CAMP."; break;
                    case "070": laduana = "CD. JUAREZ, CHIH."; break;
                    case "080": laduana = "COATZACOALCOS, VER."; break;
                    case "110": laduana = "ENSENADA, B.C."; break;
                    case "120": laduana = "GUAYMAS, SON."; break;
                    case "140": laduana = "LA PAZ, B.C.S."; break;
                    case "160": laduana = "MANZANILLO, COL."; break;
                    case "170": laduana = "MATAMOROS,TAMPS."; break;
                    case "180": laduana = "MAZATLAN, SIN."; break;
                    case "190": laduana = "MEXICALI, B.C."; break;
                    case "200": laduana = "MEXICO"; break;
                    case "220": laduana = "NACO, SON."; break;
                    case "230": laduana = "NOGALES, SON."; break;
                    case "240": laduana = "NUEVO LAREDO, TAMPS."; break;
                    case "250": laduana = "OJINAGA, CHIH."; break;
                    case "260": laduana = "PUERTO PALOMAS, CHIH."; break;
                    case "270": laduana = "PIEDRAS NEGRAS, COAH."; break;
                    case "280": laduana = "PROGRESO, YUC."; break;
                    case "300": laduana = "CD. REYNOSA, TAMPS."; break;
                    case "310": laduana = "SALINA CRUZ, OAX."; break;
                    case "330": laduana = "SAN LUIS RIO COLORADO, SON."; break;
                    case "340": laduana = "CD. MIGUEL ALEMAN, TAMPS."; break;
                    case "370": laduana = "CD. HIDALGO, CHIS."; break;
                    case "380": laduana = "TAMPICO, TAMPS."; break;
                    case "390": laduana = "TECATE, B.C."; break;
                    case "400": laduana = "TIJUANA, B.C."; break;
                    case "420": laduana = "TUXPAN, VER."; break;
                    case "430": laduana = "VERACRUZ, VER."; break;
                    case "440": laduana = "CD. ACUNA, COAH."; break;
                    case "460": laduana = "TORREON, COAH."; break;
                    case "470": laduana = "AEROPUERTO INTERNAL. CD. DE MEXICO, D.F."; break;
                    case "480": laduana = "GUADALAJARA, JAL."; break;
                    case "500": laduana = "SONOYTA, SON."; break;
                    case "510": laduana = "LAZARO CARDENAS, MICH."; break;
                    case "520": laduana = "MONTERREY, N.L."; break;
                    case "530": laduana = "CANCUN, Q. ROO."; break;
                    case "640": laduana = "QUERETARO, QRO."; break;
                    case "650": laduana = "TOLUCA, MEX."; break;
                    case "670": laduana = "CHIHUAHUA, CHIH."; break;
                    case "730": laduana = "AGUASCALIENTES, AGS."; break;
                    case "750": laduana = "PUEBLA, PUE."; break;
                    case "800": laduana = "COLOMBIA, N.L."; break;
                    case "810": laduana = "ALTAMIRA, TAMPS."; break;
                    case "820": laduana = "CD. CAMARGO, TAMPS."; break;
                    case "830": laduana = "DOS BOCAS"; break;
                    case "840": laduana = "GUANAJUATO, GTO"; break;
                }


            SqlCommand lsql = new SqlCommand();
            lsql.CommandText = "insert into admCapasProducto (CIDCAPA,CIDALMACEN,CIDPRODUCTO,CFECHA,CTIPOCAPA,CNUMEROLOTE,CFECHACADUCIDAD,CFECHAFABRICACION, " +
            " CPEDIMENTO,CADUANA,CFECHAPEDIMENTO,CTIPOCAMBIO,CEXISTENCIA,CCOSTO,CTIMESTAMP,CNUMADUANA) "+ 
            " values ("+
            lultimacapa + "," + movto._RegCapa._Almacen.Id.ToString() + "," + movto._RegProducto.Id.ToString().Trim() + ",'" + lfechahoy + "'," + "2,''," + "'18991230'," + "'18991230'," + "'" + movto._RegCapa.Pedimento + "'," 
            + "'" + laduana + "','"+ lfechaped + "'," +movto._RegCapa.tc.ToString().Trim() + "," +  movto.cUnidades + "," + movto.cPrecio + ",'" + lfechatime + "'," + movto._RegCapa.NoAduana + ")";
            lsql.Connection = miconexion._conexion1;
            int lret = lsql.ExecuteNonQuery();
            

            lsql.CommandText = "insert into admMovimientosCapas (CIDMOVIMIENTO,CIDCAPA,CFECHA,CUNIDADES,CTIPOCAPA,CIDUNIDAD) " +
            " values (" +
            movto.cIdMovto + "," + lultimacapa + ",'" + lfechahoy + "'," + movto.cUnidades + ",2,1)";
            //lsql.Connection = miconexion._conexion1;
             lret = lsql.ExecuteNonQuery();

             lsql.CommandText = "update admMovimientos set cunidades = " + movto.cUnidades.ToString().Trim() + ",cunidadescapturadas = " + movto.cUnidades.ToString().Trim() + ",cunidadespendientes = " + movto.cUnidades.ToString().Trim() + " where cidmovimiento = " + movto.cIdMovto;
             lsql.Connection = miconexion._conexion1;
             lret = lsql.ExecuteNonQuery();


        }

        public int mRegresaUltimaCapa()
        {
            int lregresa = 1;
            SqlCommand lsql = new SqlCommand();
            SqlDataReader lreader;
            //miconexion.mAbrirConexionDestino();
            lsql.CommandText = "select max(cidcapa) +1 from admCapasProducto";
            lsql.Connection = miconexion._conexion1;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                string x;
                lreader.Read();
                try
                {
                    x = lreader[0].ToString();
                }
                catch (Exception ee)
                {
                    x = "1";
                }
                lreader.Close();
                //miconexion.mCerrarConexionDestino();
                lregresa = int.Parse(x);
            }
            return lregresa;
        }

        public RegProducto mBuscarClasificacion(string codigo, int anumClasif, int tipocatalogo)
        {
            OleDbConnection lconexion = new OleDbConnection();
            //OleDbDataReader lreader;
            RegProducto lprod = new RegProducto();
            string lcadena="";
            string sLimite = "";

            /*
                switch (tipocatalogo)
                {
                    case 4:
                        switch (anumClasif)
                        {
                            case 1:
                                sLimite = "25";
                                sLimite = "1";
                                break;
                            case 2:
                                sLimite = "26";
                                sLimite = "2";
                                break;
                            case 3:
                                sLimite = "27";
                                sLimite = "3";
                                break;
                            case 4:
                                sLimite = "28";
                                break;
                            case 5:
                                sLimite = "29";
                                break;
                            case 6:
                                sLimite = "30";
                                break;

                        }
                        break;
                }*/
            sLimite = anumClasif.ToString();

            lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {
                lcadena = "select cidvalor01,cvalorcl01 from mgw10020 where ccodigov01 = '" + codigo + "' and cidclasi01 =" + sLimite;
                
                OleDbCommand lsql = new OleDbCommand(lcadena, lconexion);
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                       
                    lprod.Codigo = lreader[1].ToString();
                    lprod.Nombre = lreader[1].ToString();
                   lprod.Id = long.Parse(lreader[0].ToString());
                      //  lprods.Add(lprod);

                    }
                }
                lreader.Close();
                miconexion.mCerrarConexionDestino();
            }
            return lprod;

        }

        public void mGrabarComplemento(List<RegProducto> lista, string aValor1, string aValor2, string aValor3, string aValor4)
        {
            OleDbConnection lconexion = new OleDbConnection();
            //OleDbDataReader lreader;
            List<RegProducto> lprods = new List<RegProducto>();
            string lcadena = "";

            lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {
                foreach (RegProducto lprod in lista)
                {


                    //lcadena = "select cidvalor01,cvalorcl01 from mgw10020 where ccodigov01 = '" + codigo + "' and cidclasi01 =" + anumClasif.ToString();
                    lcadena = "UPDATE mgw10046 set valor = '" + aValor1 + "' where idaddenda = 336 and tipocat = 2 and idcat = " + lprod.Id.ToString() + " and numCampo = 1";
                    OleDbCommand lsql = new OleDbCommand(lcadena, lconexion);
                    int lcuantos = lsql.ExecuteNonQuery();

                    if (lcuantos == 0)
                    {
                        lcadena = "insert into mgw10046 values (336,2," + lprod.Id.ToString() + ",1,'" + aValor1 + "')";
                        lsql.CommandText = lcadena;
                        lsql.ExecuteNonQuery();
                    }
                    lcadena = "UPDATE mgw10046 set valor = '" + aValor2 + "' where idaddenda = 336 and tipocat = 2 and idcat = " + lprod.Id.ToString() + " and numCampo = 2";
                    lsql.CommandText = lcadena;
                    lcuantos = lsql.ExecuteNonQuery();

                    if (lcuantos == 0)
                    { 
                        lcadena = "insert into mgw10046 values (336,2," + lprod.Id.ToString() + ",2,'" + aValor2 + "')";
                        lsql.CommandText = lcadena;
                        lsql.ExecuteNonQuery();
                    }
                    lcadena = "UPDATE mgw10046 set valor = '" + aValor3 + "' where idaddenda = 336 and tipocat = 2 and idcat = " + lprod.Id.ToString() + " and numCampo = 3";
                    lsql.CommandText = lcadena;
                    lcuantos = lsql.ExecuteNonQuery();

                    if (lcuantos == 0)
                    {
                        lcadena = "insert into mgw10046 values (336,2," + lprod.Id.ToString() + ",3,'" + aValor3 + "')";
                        lsql.CommandText = lcadena;
                        lsql.ExecuteNonQuery();
                    }
                    lcadena = "UPDATE mgw10046 set valor = '" + aValor4 + "' where idaddenda = 336 and tipocat = 2 and idcat = " + lprod.Id.ToString() + " and numCampo = 4";
                    lsql.CommandText = lcadena;
                    lcuantos = lsql.ExecuteNonQuery();

                    if (lcuantos == 0)
                    {
                        lcadena = "insert into mgw10046 values (336,2," + lprod.Id.ToString() + ",4,'" + aValor4 + "')";
                        lsql.CommandText = lcadena;
                        lsql.ExecuteNonQuery();
                    }
                }
                    

                
                miconexion.mCerrarConexionDestino();
            }
            
        }

        public List<RegProducto> mMostrarProductos(string aCodigo1, string aCodigo2,string aCodigo3,string aCodigo4,string aCodigo5,string aCodigo6)
        {
            OleDbConnection lconexion = new OleDbConnection();
            //OleDbDataReader lreader;
            List< RegProducto> lprods = new List<RegProducto>();
            string lcadena="";



            lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {
                //lcadena = "select cidvalor01,cvalorcl01 from mgw10020 where ccodigov01 = '" + codigo + "' and cidclasi01 =" + anumClasif.ToString();
                lcadena = "select m5.* from mgw10005 m5 ";
               
                if (aCodigo1 !="")
                    lcadena += "join mgw10020 m20a " +
                    " on  m5.cidvalor01 = m20a.cidvalor01 " +
                    " and m20a.ccodigov01 = '" + aCodigo1 + "'";

                if (aCodigo2 !="")
                    lcadena += " join mgw10020 m20b  " +
                    " on  m5.cidvalor02 = m20b.cidvalor01  " +
                    " and m20b.ccodigov01 = '" + aCodigo2 + "'";
                if (aCodigo3 != "")
                    lcadena += " join mgw10020 m20c  " +
                    " on  m5.cidvalor03 = m20c.cidvalor01  " +
                    " and m20c.ccodigov01 = '" + aCodigo3 + "'";
                if (aCodigo4 != "")
                    lcadena += " join mgw10020 m20d  " +
                    " on  m5.cidvalor04 = m20d.cidvalor01  " +
                    " and m20d.ccodigov01 = '" + aCodigo4 + "'";
                if (aCodigo5 != "")
                    lcadena += " join mgw10020 m20e  " +
                    " on  m5.cidvalor05 = m20e.cidvalor01  " +
                    " and m20e.ccodigov01 = '" + aCodigo5 + "'";
                if (aCodigo6 != "")
                    lcadena += " join mgw10020 m20f  " +
                    " on  m5.cidvalor06 = m20f.cidvalor01  " +
                    " and m20f.ccodigov01 = '" + aCodigo6 + "'";

                OleDbCommand lsql = new OleDbCommand(lcadena, lconexion);
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegProducto lprod = new RegProducto();
                    lprod.Codigo = lreader[1].ToString();
                    lprod.Nombre = lreader[2].ToString();
                   lprod.Id = long.Parse(lreader[0].ToString());
                        lprods.Add(lprod);

                    }
                }
                lreader.Close();
                miconexion.mCerrarConexionDestino();
            }
            return lprods;

        }


        public RegProducto mBuscarClasificacion1(string codigo, string anumClasif)
        {
            OleDbConnection lconexion = new OleDbConnection();
            //OleDbDataReader lreader;
            RegProducto lprod = new RegProducto();
            string lcadena;



            lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {
                lcadena = "select cidvalor01,cvalorcl01 from mgw10020 where ccodigov01 = '" + codigo + "' and cidclasi01 =" + anumClasif.ToString();


                OleDbCommand lsql = new OleDbCommand(lcadena, lconexion);
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    //lprod.Codigo = lreader[1].ToString();
                    lprod.Nombre = lreader[1].ToString();
                    lprod.Id = long.Parse(lreader[0].ToString());
                }
                lreader.Close();
                miconexion.mCerrarConexionDestino();
            }
            return lprod;




        }

        public RegProducto mBuscarProducto(string codigo)
        {
            OleDbConnection lconexion = new OleDbConnection();
            //OleDbDataReader lreader;
            RegProducto lprod = new RegProducto();
            string lcadena;
            
                
            
                lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {
                lcadena = "select cidprodu01,ccodigop01,cnombrep01 from mgw10005 where ccodigop01 = '" + codigo + "'";
             

                OleDbCommand lsql = new OleDbCommand(lcadena, lconexion);
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    lprod.Codigo = lreader[1].ToString();
                    lprod.Nombre = lreader[2].ToString();
                    lprod.Id = long.Parse(lreader[0].ToString());
                }
                lreader.Close();
                miconexion.mCerrarConexionDestino();
            }
            return lprod;

            


        }

        public string mGrabarmGrabarComplemento(string codigoini, string codigofin)
        {
            OleDbConnection lconexion = new OleDbConnection();
            //OleDbDataReader lreader;
            List<RegProducto> lprod = new List<RegProducto>();
            string lcadena;



            lconexion = miconexion.mAbrirConexionDestino();

            lcadena = "select cidprodu01 from mgw10005 where ccodigop01 >= '" + codigoini + "' and ccodigop01 <= '" + codigofin + "'";


            OleDbCommand lsql = new OleDbCommand(lcadena, lconexion);
            OleDbDataReader lreader;
            //long lIdDocumento = 0;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                while (lreader.Read())
                {
                    long Id = long.Parse(lreader[0].ToString());


                }

            }
            return "";
        }


        public RegProducto mBuscarProductoComercial(string codigo)
        {
            

            //OleDbConnection lconexion = new OleDbConnection();
            miconexion.mAbrirConexionComercial(false);
            //lconexion = miconexion._conexion;
            RegProducto lprod = new RegProducto();

            string lquery = "select ccodigoproducto from mgw10005 where ccodigop01 = '" + codigo + "'";

            SqlCommand lsql = new SqlCommand ();
            SqlDataReader lreader;

               // miconexion.mAbrirConexionDestino();

            lsql.CommandText = "select cidproducto,ccodigoproducto,cnombreproducto,CIMPORTEEXTRA1, cprecio1 from admProductos where ccodigoproducto = '" + codigo + "'";
            lsql.Connection = miconexion._conexion1;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            string lregresa = "";
            string x = "";

            if (lreader.HasRows)
            {
                lreader.Read();
                try
                {
                    lprod.Id = long.Parse(lreader[0].ToString());
                    lprod.Codigo = lreader[1].ToString();
                    lprod.Nombre = lreader[2].ToString();
                    lprod.ImporteExtra1 = decimal.Parse(lreader[3].ToString());
                    lprod.Precio = double.Parse(lreader[4].ToString());

                }
                catch (Exception ee)
                {
//                    lreader.Close();
                }
                lreader.Close();
            }
            miconexion.mCerrarConexionOrigenComercial();
            return lprod;


        }


        public RegCliente mBuscarClienteComercial(string codigo)
        {
            Boolean lcerrar = false;
            if (miconexion._conexion1 == null)
            {
                miconexion.mAbrirConexionComercial(false);
                lcerrar = true;
            }
            //lconexion = miconexion._conexion;
            RegCliente lcliente = new RegCliente();

            string lquery = "";

            SqlCommand lsql = new SqlCommand();
            SqlDataReader lreader;

            // miconexion.mAbrirConexionDestino();

            lsql.CommandText = "select cidclienteproveedor,ccodigocliente,crazonsocial, cmetodopag from admClientes where ccodigocliente = '" + codigo + "'";
            lsql.Connection = miconexion._conexion1;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            string lregresa = "";
            string x = "";

            if (lreader.HasRows)
            {
                lreader.Read();
                try
                {
                    lcliente.Id = long.Parse(lreader[0].ToString());
                    lcliente.Codigo = lreader[1].ToString();
                    lcliente.RazonSocial= lreader[2].ToString();
                    lcliente.MetodoPago = lreader[3].ToString();

                }
                catch (Exception ee)
                {
                    //                    lreader.Close();
                }
                lreader.Close();
            }
            if (lcerrar == true)
            {
                miconexion.mCerrarConexionOrigenComercial();
            }
            return lcliente;


        }

        public string mBuscarDescripcionProducto(string descripcion)
        {

            //OleDbConnection lconexion = new OleDbConnection();
            //miconexion.mAbrirConexionDestino();
            //lconexion = miconexion._conexion;


            string lquery = "select ccodigop10 from mgw10005 where cnombrep01 = '" + descripcion + "'";

            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;

                miconexion.mAbrirConexionDestino();

            lsql.CommandText = "select ccodigop10 from mgw10005 where cnombrep01 = '" + descripcion + "'";
            lsql.Connection = miconexion._conexion;
            
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            string  lregresa = "";
            string x= "" ;
                
            if (lreader.HasRows)
            {
                lreader.Read();
                try
                {
                    x = lreader[0].ToString();
                }
                catch (Exception ee)
                {
                    lreader.Close();
                    lsql.CommandText = "select max(cidprodu01) from mgw10005 ";
                    lreader = lsql.ExecuteReader();
                    if (lreader.HasRows)
                    {
                        lreader.Read();
                        try
                        {
                            x = lreader[0].ToString();
                        }
                        catch (Exception eee)
                        {
                            x = "1";
                        }
                    }
                }
                lreader.Close();
                miconexion.mCerrarConexionDestino();

                lregresa = x;
            }
            return lregresa;


        }


    }
    
}
