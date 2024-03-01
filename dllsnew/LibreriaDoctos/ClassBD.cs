using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Configuration;
using System.IO;
//using BarradeProgreso;
using System.Data.SqlClient;
using Interfaces;
using System.Collections;
using System.Data;
using System.Globalization;
using System.Net;
using System.Linq;
using System.Xml;
using System.Data.Odbc;
using Newtonsoft.Json;
using System.Windows.Forms;


//testing
namespace LibreriaDoctos
{
    public class ClassBD : ISujeto
    {


        public List<string> lvar = new List<string>();
        protected decimal lsubtotal;
        protected decimal limpuestos;
        protected string aRutaExe;
        public string productos;
        public string almacenes;
        public bool sdkcomercial = false;
        public RegDocto primerdocto = new RegDocto();

        public string cadenaconexion = "";

        List<IObservador> lista = new List<IObservador>();



        [DllImport("MGW_SDK.DLL")] static extern int fLeeDatoMovimiento(string aCampo, StringBuilder aMensaje, int aLen);


        [DllImport("MGW_SDK.DLL")] static extern int fInsertaCteProv();
        [DllImport("MGW_SDK.DLL")] static extern int fEditaCteProv();
        [DllImport("MGW_SDK.DLL")] static extern int fGuardaCteProv();
        [DllImport("MGW_SDK.DLL")] static extern int fSetDatoCteProv(string aCampo, string aValor);

        [DllImport("MGW_SDK.DLL")] static extern int fInsertaProducto();
        [DllImport("MGW_SDK.DLL")] static extern int fGuardaProducto();
        [DllImport("MGW_SDK.DLL")] static extern int fSetDatoProducto(string aCampo, string aValor);

        [DllImport("MGW_SDK.DLL")] static extern int fInsertaAlmacen();
        [DllImport("MGW_SDK.DLL")] static extern int fGuardaAlmacen();
        [DllImport("MGW_SDK.DLL")] static extern int fSetDatoAlmacen(string aCampo, string aValor);

        [DllImport("MGW_SDK.DLL")] static extern int fInsertarDocumento();
        [DllImport("MGW_SDK.DLL")] static extern int fGuardaDocumento();


        [DllImport("MGW_SDK.DLL")] static extern int fSetDatoDocumento(string aCampo, string aValor);
        [DllImport("MGW_SDK.DLL")] static extern int fInsertarMovimiento();
        [DllImport("MGW_SDK.DLL")] static extern int fGuardaMovimiento();
        [DllImport("MGW_SDK.DLL")] static extern int fSetDatoMovimiento(string aCampo, string aValor);
        [DllImport("MGW_SDK.DLL")] static extern int fInsertaDireccion();

        [DllImport("MGW_SDK.DLL")] static extern int fGuardaDireccion();
        [DllImport("MGW_SDK.DLL")] static extern int fBorraDocumento();
        //[DllImport("MGW_SDK.DLL")]        static extern int fBorraMovimiento();


        [DllImport("MGW_SDK.DLL")] static extern int fSetDatoDireccion(string aCampo, string aValor);
        [DllImport("MGW_SDK.DLL")] static extern int fAfectaDocto_Param(string aConcepto, string aSerie, double aFolio, Boolean aAfecta);
        [DllImport("MGW_SDK.DLL")] static extern int fError(int aNumErrror, StringBuilder aError, long aLen);

        [DllImport("MGW_SDK.DLL")]
        static extern int fSiguienteFolio(string lCodigoConcepto, ref string lSerie, ref double lFolio);

        [DllImport("MGW_SDK.DLL")]
        static extern int fInicializaLicenseInfo(int aSistema);

        //Private Declare Function fEmitirDocumento Lib "MGW_SDK.DLL" (ByVal aCodigoConcepto As String, ByVal aNumSerie As String, ByVal aFolio As Double, ByVal aPassword As String, ByVal aArchivo As String) As Long
        [DllImport("MGW_SDK.DLL")]
        static extern int fEmitirDocumento(string aCodigoConcepto, string aNumSerie, double aFolio, string aPassword, string aArchivo);

        [DllImport("MGW_SDK.DLL")]
        static extern int fEntregEnDiscoXML(string aCodigoConcepto, string aNumSerie, double aFolio, int aFormato, ref string aFormatoamigo);

        //(aCodConcepto, aSerie, aFolio, aFormato, aFormatoAmig)
        //lError = fEntregEnDiscoXML (“4”, “B1”, 45, 1, “C:\Compacw\Empresas\Reportes\AdminPAQ\Plantilla_Factura_cfdi_1.html”)

        [DllImport("MGW_SDK.DLL")]
        static extern long fSaldarDocumento_Param(string lCodConcepto_Pagar, string lSerie_Pagar, double lFolio_Pagar,
string lCodConcepto_Pago, string lSerie_Pago, double lFolio_Pago, double lImporte, int lIdMoneda, string lFecha);


        [DllImport("MGW_SDK.DLL")]
        static extern long fRegresaExistencia(string lCodigoProducto, string lCodigoAlmacen, string lAnio, string lMes, string lDia, ref double lExistencia);



        [DllImport("MGW_SDK.DLL")]
        static extern int fLeeDatoProducto(string aCampo, StringBuilder aMensaje, int aLen);


        // Need this DllImport statement to reset the floating point register below
        [DllImport("msvcr71.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int _controlfp(int n, int mask);

        [DllImport("KERNEL32.DLL")]
        static extern int SetCurrentDirectory(string pPtrDirActual);
        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fSetNombrePAQ(string aSistema);
        //[DllImport("MGWSERVICIOS.DLL")]
        //static extern int fError(int aNumError, StringBuilder aMensaje, int aLen);
        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fTerminaSDK();
        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fPosPrimerEmpresa(ref int aIdEmpresa, ref string aNombreEmpresa, ref string aDirectorioEmpresa);

        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fAbreEmpresa(string aDirectorioEmpresa);


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fRegresaExistencia")]
        static extern long fRegresaExistenciaComercial(string lCodigoProducto, string lCodigoAlmacen, string lAnio, string lMes, string lDia, ref double lExistencia);


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fCancelaCambiosMovimiento")]
        static extern long fCancelaCambiosMovimientoComercial();



        [DllImport("MGWSERVICIOS.DLL")]
        static extern void fCierraEmpresa();

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fSiguienteFolio")]
        static extern int fSiguienteFolioComercial(string aCodigoConcepto, ref string aSerie, ref double aFolio);

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
        static extern int fAltaMovimientoSeriesCapas_ParamComercial(string aIdMovimiento, string aUnidades, string aTipoCambio, string aSeries,
 string aPedimento, string aAgencia, string aFechaPedimento, string aNumeroLote, string aFechaFabricacion, string aFechaCaducidad);


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


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fLeeDatoProducto")]
        static extern int fLeeDatoProductoComercial(string aCampo, StringBuilder aMensaje, int aLen);


        //public static extern void fError(int NumeroError, StringBuilder Mensaje, int Longitud);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fLeeDatoMovimiento")]
        static extern int fLeeDatoMovimientoComercial(string aCampo, StringBuilder aMensaje, int aLen);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fModificaCostoEntrada")]
        static extern int fModificaCostoEntradaComercial(string aIdMovimiento, string aCostoEntrada);


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fBuscarIdMovimiento")]
        static extern int fBuscarIdMovimientoComercial(int aIdMovimeinto);


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fBuscaAlmacen")]
        static extern int fBuscaAlmacenComercial(string aCodigoAlmacen);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fBorraDocumento")]
        static extern int fBorraDocumentoComercial();


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fBorraMovimiento")]
        static extern int fBorraMovimientoComercial();


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fInicializaLicenseInfo")]
        static extern int fInicializaLicenseInfoComercial(int aSistema);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fEmitirDocumento")]
        static extern int fEmitirDocumentoComercial(string aCodigoConcepto, string aNumSerie, double aFolio, string aPassword, string aArchivo);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fLeeDatoCteProv")]
        static extern int fLeeDatoCteProvComercial(string aCampo, StringBuilder aMensaje, int aLen);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fEntregEnDiscoXML")]
        static extern int fEntregEnDiscoXMLComercial(string aCodigoConcepto, string aNumSerie, double aFolio, int aFormato, string aFormatoAmig);


        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fAfectaDocto_Param")]

        static extern int fAfectaDocto_ParamComercial(string aConcepto, string aSerie, double aFolio, Boolean aAfecta);

        [DllImport("MGWSERVICIOS.DLL", EntryPoint = "fAfectaDocto")]
        private static extern int fAfectaDocto_Comercial(documento doc, bool aAfecta);


        public struct documento
        {
            public string aConcepto;
            public string aSerie;
            public double aFolio;
            public Boolean aAfecta;
        };

        //fEntregEnDiscoXML (aCodConcepto, aSerie, aFolio, aFormato, aFormatoAmig)


        //protected ClassConexion miconexion;
        public ClassConexion miconexion = new ClassConexion();
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


        protected OleDbConnection _con;

        public string Cadenaconexion;
        public string cserver;
        public string cbd;
        public string cusr;
        public string cpwd;


        public string mLlenarinfoAutorizaciones(int liddocumento, string concepto, string doctode)
        {
            RegDocto lDocto = new RegDocto();

            SqlConnection lconexion = new SqlConnection();
            lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {

                SqlCommand lsql = new SqlCommand("select  FORMAT(m8.cfecha,'dd-MM-yyyy') as cfecha,  m2.crazonsocial crazonso01, m2.cidclienteproveedor cidclien01, m8.ctotal,m2.ccodigocliente, " +
                    "m10.cpreciocapturado as cprecio, m10.cimpuesto1," +
                    "m2.crfc as rfc, m10.cunidades as unidades, m10.cneto, m5.ccodigoproducto, m3.ccodigoalmacen, m8.ciddocumento, m10.cidmovimiento, m8.cTotalUnidades " +
                    "from admdocumentos m8 " +
" join admMovimientos m10 on m10.ciddocumento = m8.ciddocumento " +
" join admAlmacenes m3 on m3.cidalmacen = m10.cidalmacen " +
" join admProductos m5 on m5.cidproducto = m10.cidproducto" +
" join admclientes m2 on m2.cidclienteproveedor = m8.cidclienteproveedor " +
" WHERE m8.ciddocumento = " + liddocumento.ToString()
//" AND m8.ccancelado = 0"
, lconexion);

                Boolean noseguir = false;
                SqlDataReader dr;
                dr = lsql.ExecuteReader();
                if (dr.HasRows)
                {

                    long lfolioleido = 0;
                    string cserie = "";
                    //dr.Read();
                    long lFoliox;
                    string aSerie = "";
                    double aFolio = 0;
                    string lclienteanterior = "";
                    while (noseguir == false)
                    {

                        if (!dr.Read())
                            break;

                        string lcliente = dr["ccodigocliente"].ToString();
                        if (lcliente == "")
                            break;


                        //lFoliox = fSiguienteFolioComercial("2", ref aSerie, ref aFolio);



                        if (lcliente != lclienteanterior && lclienteanterior != "")
                        {
                            _RegDoctos.Add(lDocto);
                            lDocto = new RegDocto();
                        }


                        if (lcliente != lclienteanterior)
                        {
                            lDocto.cSerie = aSerie;
                            lDocto.cCodigoCliente = dr["ccodigoCliente"].ToString();
                            //lcliente = lDocto.cCodigoCliente;
                            lDocto.cCodigoConcepto = concepto;
                            lDocto.cReferencia = doctode;


                            lDocto.cFolio = -1;
                            //--lFoliox++;
                            try
                            {
                                //lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoP");
                                lDocto.cFecha = DateTime.Parse(dr["cfecha"].ToString());

                            }
                            catch (Exception eeeeee)
                            {
                                //lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");
                                lDocto.cFecha = DateTime.Parse(dr["Fecha"].ToString());
                            }
                            //lfolioleido = lFoliox;
                            lDocto.cMoneda = "Peso Mexicano";
                            lDocto.cTipoCambio = 1;



                            lDocto._RegCliente.RazonSocial = dr["crazonso01"].ToString();
                            lDocto._RegCliente.RazonSocial = dr["crazonso01"].ToString();

                            lDocto.cIdDocto = long.Parse(dr["ciddocumento"].ToString());
                            lDocto.cTotalUnidades = long.Parse(dr["cTotalUnidades"].ToString());
                        }

                        RegMovto regmov = new RegMovto();


                        try
                        {
                            regmov.cCodigoProducto = dr["ccodigoproducto"].ToString();

                            regmov.cUnidades = decimal.Parse(dr["unidades"].ToString());
                        }
                        catch (Exception eeeeee)
                        {
                        }
                        regmov.cCodigoAlmacen = dr["ccodigoalmacen"].ToString();


                        try
                        {
                            regmov.cPrecio = decimal.Parse(dr["cprecio"].ToString());
                            regmov.cImpuesto = decimal.Parse(dr["cImpuesto1"].ToString());

                            regmov.cIdMovto = long.Parse(dr["cidmovimiento"].ToString());
                        }

                        catch (Exception yyyyyyy)
                        {
                            regmov.cPrecio = decimal.Parse(dr["Neto"].ToString());
                            regmov.cImpuesto = decimal.Parse(dr["cImpuesto1"].ToString());
                            regmov.cTotal = decimal.Parse(dr["cTotal"].ToString());

                        }

                        lclienteanterior = dr["ccodigoCliente"].ToString();
                        lDocto._RegMovtos.Add(regmov);

                        //dr.Read();

                    }
                }

                dr.Close();
                if (lDocto.cCodigoCliente != "")

                    _RegDoctos.Add(lDocto);

            }
            return "";
        }


        public string mLLenarInfoPedidosFacturas(string archivo)
        {
            //string archivo1 = @archivo;
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @archivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");

            // OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @archivo + ";Extended Properties=" + Convert.ToChar(34).ToString() + @"Excel 8.0" + Convert.ToChar(34).ToString() + ";");

            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            try
            {
                conn.Open();

                cmd.Connection = conn;
                cmd.CommandText = "SELECT * FROM [Hoja1$]";

                cmd.ExecuteNonQuery();
            }
            catch (Exception eeeee)
            {
                return eeeee.Message;
            }

            System.Data.OleDb.OleDbDataReader dr;
            dr = cmd.ExecuteReader();
            Boolean noseguir = false;
            _RegDoctos.Clear();
            List<RegDocto> doctos = new List<RegDocto>();
            RegDocto lDocto = new RegDocto();
            if (dr.HasRows)
            {
                long lfolioleido = 0;
                string cserie = "";
                //dr.Read();
                long lFoliox;
                while (noseguir == false)
                {

                    dr.Read();

                    string lcliente = dr["Cliente ID"].ToString();
                    if (lcliente == "")
                        break;

                    try
                    {
                        lFoliox = long.Parse(dr["Folio dispensador"].ToString());
                    }
                    catch (Exception eee)
                    {
                        lFoliox = long.Parse(dr["Folio"].ToString());
                    }


                    if (lFoliox != lfolioleido)
                    {
                        if (lDocto.cCodigoCliente != "")
                        {
                            _RegDoctos.Add(lDocto);
                            lDocto = new RegDocto();
                        }


                        //lDocto.cSerie = cserie;
                        lDocto.cCodigoCliente = dr["Cliente ID"].ToString();
                        //lcliente = lDocto.cCodigoCliente;
                        lDocto.cCodigoConcepto = "2";


                        lDocto.cFolio = lFoliox;
                        //--lFoliox++;
                        try
                        {
                            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoP");
                            lDocto.cFecha = DateTime.Parse(dr["Fecha Ticket"].ToString());

                        }
                        catch (Exception eeeeee)
                        {
                            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");
                            lDocto.cFecha = DateTime.Parse(dr["Fecha"].ToString());
                        }
                        try
                        {
                            lDocto.cReferencia = dr["Referencia "].ToString();
                        }
                        catch (Exception iiii)
                        { }
                        lfolioleido = lFoliox;
                        lDocto.cMoneda = "Peso Mexicano";
                        lDocto.cTipoCambio = 1;
                    }

                    RegMovto regmov = new RegMovto();
                    //                    regmov.cCodigoProducto = dr["Producto"].ToString();


                    try
                    {
                        regmov.cCodigoProducto = @"001";

                        regmov.cUnidades = decimal.Parse(dr["Litros"].ToString());
                    }
                    catch (Exception eeeeee)
                    {
                        regmov.cUnidades = 1;
                        regmov.cCodigoProducto = @"(Ninguno)                     ";
                    }
                    regmov.cCodigoAlmacen = "1";

                    try
                    {
                        regmov.cPrecio = decimal.Parse(dr["Precio x Litro"].ToString());
                    }
                    catch (Exception yyyyyyy)
                    {
                        regmov.cPrecio = decimal.Parse(dr["Subtotal"].ToString());
                        regmov.cImpuesto = decimal.Parse(dr["Importe IVA"].ToString());
                        regmov.cTotal = decimal.Parse(dr["Total"].ToString());

                    }
                    //regmov.cObservaciones = dr["DESCRIPCION"].ToString();
                    lDocto._RegMovtos.Add(regmov);

                    //dr.Read();

                }


                if (lDocto.cCodigoCliente != "")

                    _RegDoctos.Add(lDocto);

            }
            return "";

        }


        public string mLLenarInfoAdrianaTraspaso(string archivo)
        {

            string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");

            StringBuilder sMensaje1 = new StringBuilder(512);
            miconexion.mAbrirConexionComercial(true);
            int lResultado = fSetNombrePAQ("CONTPAQ I Comercial");
            if (lResultado != 0)
            {
                fErrorComercial(lResultado, sMensaje1, 512);
            }
            int zzzzz = fAbreEmpresa(rutadestino);

            if (zzzzz != 0)
            {
                fErrorComercial(zzzzz, sMensaje1, 512);
            }




            //string archivo1 = @archivo;
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @archivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");

            // OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @archivo + ";Extended Properties=" + Convert.ToChar(34).ToString() + @"Excel 8.0" + Convert.ToChar(34).ToString() + ";");

            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            try
            {
                conn.Open();

                cmd.Connection = conn;
                cmd.CommandText = "SELECT * FROM [Hoja1$]";

                cmd.ExecuteNonQuery();
            }
            catch (Exception eeeee)
            {
                return eeeee.Message;
            }

            System.Data.OleDb.OleDbDataReader dr;
            dr = cmd.ExecuteReader();
            Boolean noseguir = false;
            _RegDoctos.Clear();
            List<RegDocto> doctos = new List<RegDocto>();
            RegDocto lDoctos = new RegDocto();
            RegDocto lDoctoe = new RegDocto();

            lDoctos.cCodigoConcepto = "35";
            lDoctos.cSerie = "";
            lDoctos.cFecha = DateTime.Now;
            lDoctos.cMoneda = "Peso Mexicano";
            lDoctos.cTipoCambio = 1;
            //lDoctos.cCodigoCliente = "Traspaso";


            lDoctoe.cCodigoConcepto = "34";
            lDoctoe.cSerie = "";
            lDoctoe.cFecha = DateTime.Now;
            lDoctoe.cMoneda = "Peso Mexicano";
            lDoctoe.cTipoCambio = 1;

            while (dr.HasRows && noseguir == false)
            {

                while (noseguir == false)
                {
                    dr.Read();

                    RegMovto regmov = new RegMovto();
                    try
                    {
                        regmov.cError = "";
                        regmov.cCodigoProducto = dr["Codigo Producto"].ToString();
                        int encontrado = fBuscaProductoComercial(regmov.cCodigoProducto);

                        regmov.cError = "";
                        if (encontrado != 0)
                        {
                            regmov.cError = " Producto con Codigo " + regmov.cCodigoProducto + " no existe";
                        }
                        else
                        {
                            StringBuilder aValor1 = new StringBuilder(12);
                            fLeeDatoProductoComercial("cstatusproducto", aValor1, 5);
                            if (aValor1.ToString() == "0")
                                regmov.cError = " Producto con Codigo " + regmov.cCodigoProducto + " esta inactivo";

                        }


                        regmov.cUnidades = decimal.Parse(dr["Unidades"].ToString());
                        regmov.cCodigoAlmacen = dr["Codigo Almacen Origen"].ToString();
                        int encontrado1 = fBuscaAlmacenComercial(regmov.cCodigoAlmacen);

                        //regmov.cError = "";
                        if (encontrado1 != 0)
                        {
                            regmov.cError += " Almacen Origen con Codigo " + regmov.cCodigoAlmacen + " no existe";
                        }

                        if (regmov.cError == "")
                        {
                            double existencia = 0;
                            fRegresaExistenciaComercial(regmov.cCodigoProducto, regmov.cCodigoAlmacen, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(), ref existencia);

                            regmov.cError = "";
                            if (double.Parse(regmov.cUnidades.ToString()) > existencia)
                            {
                                regmov.cError = "No hay existencias";
                            }
                        }
                        lDoctos._RegMovtos.Add(regmov);
                        RegMovto regmovt = new RegMovto();
                        regmovt.cCodigoAlmacen = dr["Codigo Almacen Destino"].ToString();
                        int encontrado2 = fBuscaAlmacenComercial(regmovt.cCodigoAlmacen);

                        if (encontrado2 != 0)
                        {
                            regmov.cError += " Almacen Destino con Codigo " + regmovt.cCodigoAlmacen + " no existe";
                        }
                        regmovt.cCodigoProducto = dr["Codigo Producto"].ToString();
                        regmovt.cUnidades = decimal.Parse(dr["Unidades"].ToString());
                        regmovt.cError = regmov.cError;



                        lDoctoe._RegMovtos.Add(regmovt);



                    }
                    catch (Exception eee)
                    {
                        noseguir = true;
                        break;
                    }



                    //dr.Read();
                }
            }



            //      if (lDocto.cCodigoCliente != "")

            _RegDoctos.Add(lDoctos);
            _RegDoctos.Add(lDoctoe);

            return "";

        }


        public string mLLenarInfoMontesori(string archivo)
        {
            //string archivo1 = @archivo;
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @archivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");

            // OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @archivo + ";Extended Properties=" + Convert.ToChar(34).ToString() + @"Excel 8.0" + Convert.ToChar(34).ToString() + ";");

            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            try
            {
                conn.Open();

                cmd.Connection = conn;
                cmd.CommandText = "SELECT * FROM [Sheet$]";

                cmd.ExecuteNonQuery();
            }
            catch (Exception eeeee)
            {
                return eeeee.Message;
            }

            System.Data.OleDb.OleDbDataReader dr;
            dr = cmd.ExecuteReader();
            Boolean noseguir = false;
            _RegDoctos.Clear();
            List<RegDocto> doctos = new List<RegDocto>();
            RegDocto lDocto = new RegDocto();
            while (dr.HasRows && noseguir == false)
            {
                long lfolioleido = 0;
                string cserie = "";
                //dr.Read();
                long lFoliox = 0;
                while (noseguir == false)
                {

                    lDocto = new RegDocto();
                    dr.Read();
                    string lcliente = "";
                    try
                    {
                        lcliente = dr["Matricula"].ToString();
                    }
                    catch (Exception eee)
                    {
                        noseguir = true;
                        break;
                    }
                    if (lcliente == "")
                    {
                        lDocto.cCodigoCliente = "";
                        break;
                    }

                    string lReferencia = dr["Referencia"].ToString();

                    string lSerie = lReferencia.Substring(0, lReferencia.IndexOf("-"));
                    string lFolio = lReferencia.Substring(lReferencia.IndexOf("-") + 1);



                    try
                    {
                        lFoliox = long.Parse(lFolio);
                    }
                    catch (Exception eee)
                    {
                        // lFoliox = long.Parse(dr["Folio"].ToString());
                    }







                    //lDocto.cSerie = cserie;
                    lDocto.cCodigoCliente = dr["Matricula"].ToString();
                    //lcliente = lDocto.cCodigoCliente;
                    lDocto.cCodigoConcepto = "2";


                    lDocto.cFolio = lFoliox;
                    lDocto.cSerie = lSerie;
                    //--lFoliox++;
                    try
                    {
                        lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");
                        lDocto.cFecha = DateTime.Parse(dr["Fecha"].ToString());

                    }
                    catch (Exception eeeeee)
                    {
                        lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");
                        lDocto.cFecha = DateTime.Parse(dr["Fecha"].ToString());
                    }
                    try
                    {
                        lDocto.cReferencia = lReferencia;
                    }
                    catch (Exception iiii)
                    { }
                    lfolioleido = lFoliox;
                    lDocto.cMoneda = "Peso Mexicano";
                    lDocto.cTipoCambio = 1;
                    lDocto.cTextoExtra1 = dr["UUID"].ToString();
                    lDocto.cRFC = dr["RFC"].ToString();
                    lDocto.cFormaPago = dr["Forma de Pago"].ToString();
                    lDocto.cObservaciones = dr["Concepto"].ToString();

                    lDocto.cRazonSocial = dr["Razón Social"].ToString();

                    lDocto.cMoneda = "Peso Mexicano";
                    lDocto.cTipoCambio = 1;
                    RegMovto regmov = new RegMovto();
                    //                    regmov.cCodigoProducto = dr["Producto"].ToString();


                    try
                    {
                        regmov.cCodigoProducto = dr["Matricula"].ToString();

                        regmov.cUnidades = 1;
                    }
                    catch (Exception eeeeee)
                    {
                        regmov.cUnidades = 1;
                        regmov.cCodigoProducto = dr["Matricula"].ToString();
                    }
                    regmov.cCodigoAlmacen = "1";

                    try
                    {
                        regmov.cPrecio = decimal.Parse(dr["SubTotal"].ToString());
                        regmov.cImpuesto = 0;
                        regmov.cTotal = decimal.Parse(dr["SubTotal"].ToString());
                    }
                    catch (Exception yyyyyyy)
                    {

                    }
                    //regmov.cObservaciones = dr["DESCRIPCION"].ToString();

                    lDocto._RegMovtos.Add(regmov);
                    _RegDoctos.Add(lDocto);

                    //dr.Read();
                }
            }



            if (lDocto.cCodigoCliente != "")

                _RegDoctos.Add(lDocto);


            return "";

        }




        public string mLLenarInfoAddendas(string archivo)
        {
            //string archivo1 = @archivo;
            //OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @archivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");
            //OleDbConnection conn = new OleDbConnection();
            string conn1 = "";
            /*if (archivo.CompareTo(".xlsx") != 0)
                conn1 = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + archivo + ";Extended Properties='Excel 8.0;HDT=Yes;IMEX=1';"; //for below excel 2007  
            else*/
            conn1 = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + archivo + ";Extended Properties='Excel 12.0 Xml;HDR=YES';"; //for above excel 2007  

            OleDbConnection conn = new OleDbConnection(conn1);


            // OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @archivo + ";Extended Properties=" + Convert.ToChar(34).ToString() + @"Excel 8.0" + Convert.ToChar(34).ToString() + ";");

            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            try
            {
                conn.Open();

                cmd.Connection = conn;
                cmd.CommandText = "SELECT * FROM [CONTPAQDoctos$]";

                cmd.ExecuteNonQuery();
            }
            catch (Exception eeeee)
            {
                return eeeee.Message;
            }

            System.Data.OleDb.OleDbDataReader dr;
            dr = cmd.ExecuteReader();
            Boolean noseguir = false;
            _RegDoctos.Clear();
            List<RegDocto> doctos = new List<RegDocto>();
            RegDocto lDocto = new RegDocto();
            List<string> addenda = new List<string>();

            //Tipo Fecha   Serie Folio   Forma de Pago Uso CFDI Método de pago  Codigo de Cliente Codigo de Producto  Almacen Cantidad de ECD Precio Unitario NETO IVA Total Movimiento    Folio Unico de Factura FUF Fecha de la Factura Fecha Limite de Pago Cuenta de Orden del PM  Nombre del Banco Sucursal del Banco  Numero de Cuenta del Proveedor Numero de Cuenta CLABE del Proveedor Referencia del Banco    Contacto del Proveedor Num Linea Folio Unico Concepto    Cantidad Unidad  Precio Unitario Importe Linea   Importe Orig    Importe Modif   Monto Ajuste    IVA Total   Monto Letra


            cAddendaDocumento addendiux = new cAddendaDocumento();
            if (dr.HasRows)
            {

                long lfolioleido = 0;
                string cserie = "";
                //dr.Read();
                long lFoliox;

                string ltipo = "";

                while (noseguir == false)
                {

                    dr.Read();



                    try
                    {
                        ltipo = dr["Tipo"].ToString();
                        lFoliox = long.Parse(dr["Folio"].ToString());
                    }
                    catch (Exception eee)
                    {
                        break;
                    }


                    if (lFoliox != lfolioleido)
                    {
                        if (lDocto.cCodigoCliente != "")
                        {
                            foreach (string x in addenda)
                            {
                                lDocto._Addendas.Add(x);
                            }
                            lDocto.addendiux = addendiux;
                            _RegDoctos.Add(lDocto);
                            lDocto = new RegDocto();
                            addenda = new List<string>();
                            addendiux = new cAddendaDocumento();
                        }

                        lDocto.cTextoExtra1 = ltipo;
                        lDocto.cCodigoCliente = dr["Codigo de Cliente"].ToString();

                        if (ltipo == "F")
                            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");

                        if (ltipo == "NC")
                            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoD");

                        if (ltipo == "ND")
                            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoP");

                        //Forma de Pago	Uso CFDI	Método de pago

                        lDocto.cFormaPago = dr["Forma de Pago"].ToString(); ;
                        lDocto.cUsoCFDI = dr["Uso CFDI"].ToString();
                        lDocto.cMetodoPago = dr["Método de pago"].ToString();

                        lDocto.cFolio = lFoliox;
                        lDocto.cFecha = DateTime.Parse(dr["Fecha"].ToString());

                        lfolioleido = lFoliox;
                        lDocto.cMoneda = "Peso Mexicano";
                        lDocto.cTipoCambio = 1;
                        lDocto.cUUID = dr["UUID"].ToString();
                        lDocto.cTipoRelacion = dr["tipo de relacion"].ToString();



                    }

                    RegMovto regmov = new RegMovto();
                    try
                    {
                        regmov.cCodigoProducto = dr["Codigo de Producto"].ToString();

                        regmov.cImpuesto = decimal.Parse(dr["IVA"].ToString());
                        regmov.cPorcent01 = (100 * decimal.Parse(dr["IVA"].ToString())) / decimal.Parse(dr["NETO"].ToString());
                        regmov.cPorcent01 = Math.Round(regmov.cPorcent01, 0);

                        regmov.cUnidades = decimal.Parse(dr["Cantidad de ECD"].ToString());
                    }
                    catch (Exception eeeeee)
                    {
                        regmov.cUnidades = 1;
                        regmov.cCodigoProducto = @"(Ninguno)                     ";
                    }
                    regmov.cCodigoAlmacen = dr["Almacen"].ToString();

                    try
                    {
                        regmov.cPrecio = decimal.Parse(dr["Precio Unitario"].ToString());
                    }
                    catch (Exception yyyyyyy)
                    {
                        //   regmov.cPrecio = decimal.Parse(dr["Precio"].ToString());
                        // regmov.cImpuesto = decimal.Parse(dr["IVA"].ToString());
                        // regmov.cTotal = decimal.Parse(dr["Total"].ToString());

                    }

                    //regmov.cReferencia = ltipo;
                    lDocto._RegMovtos.Add(regmov);

                    if (addenda.Count == 0)
                    {

                        string lfecha = dr["Fecha de la Factura"].ToString();
                        addenda.Add(dr["Folio Unico de Factura FUF"].ToString());
                        addenda.Add(dr["Fecha de la Factura"].ToString());
                        addenda.Add(dr["Fecha Limite de Pago"].ToString());
                        addenda.Add(dr["Cuenta de Orden del PM"].ToString());
                        addenda.Add(dr["Nombre del Banco"].ToString());
                        addenda.Add(dr["Sucursal del Banco"].ToString());
                        addenda.Add(dr["Numero de Cuenta del Proveedor"].ToString());
                        addenda.Add(dr["Numero de Cuenta CLABE del Proveedor"].ToString());
                        addenda.Add(dr["Referencia del Banco"].ToString());
                        addenda.Add(dr["Contacto del Proveedor"].ToString());
                        addenda.Add(dr["Num Linea"].ToString());
                        addenda.Add(dr["Folio Unico"].ToString());
                        addenda.Add(dr["Concepto"].ToString());
                        addenda.Add(dr["Cantidad"].ToString());
                        addenda.Add(dr["Unidad"].ToString());
                        addenda.Add(dr["Precio Unitario"].ToString());
                        addenda.Add(dr["Importe Linea"].ToString());



                        addenda.Add(dr["Importe Orig"].ToString());
                        addenda.Add(dr["Importe Modif"].ToString());
                        addenda.Add(dr["Monto Ajuste"].ToString());

                        /*  addenda.Add(dr["Concepto"].ToString());
                          addenda.Add(dr["Cantidad"].ToString());
                          addenda.Add(dr["Unidad"].ToString());
                          addenda.Add(dr["Precio Unitario"].ToString());
                          */
                        //IVA Total   Monto Letra

                        addenda.Add(dr["IVA"].ToString());
                        addenda.Add(dr["Total"].ToString());
                        addenda.Add(dr["Monto Letra"].ToString());


                        addendiux.FolioUnicodeFacturaFUF = dr["Folio Unico de Factura FUF"].ToString();
                        addendiux.FechadelaFactura = dr["Fecha de la Factura"].ToString();
                        addendiux.FechaLimitedePago = dr["Fecha Limite de Pago"].ToString();
                        addendiux.CuentadeOrdendelPM = dr["Cuenta de Orden del PM"].ToString();
                        addendiux.NombredelBanco = dr["Nombre del Banco"].ToString();
                        addendiux.SucursaldelBanco = dr["Sucursal del Banco"].ToString();
                        addendiux.NumerodeCuentadelProveedor = dr["Numero de Cuenta del Proveedor"].ToString();
                        addendiux.NumerodeCuentaCLABEdelProveedor = dr["Numero de Cuenta CLABE del Proveedor"].ToString();
                        addendiux.ReferenciadelBanco = dr["Referencia del Banco"].ToString();
                        addendiux.ContactodelProveedor = dr["Contacto del Proveedor"].ToString();

                    }
                    //dr.Read();
                    cAddendaMovimiento movi = new cAddendaMovimiento();

                    movi.NumLinea = dr["Num Linea"].ToString();
                    movi.FolioUnico = dr["Folio Unico"].ToString();
                    movi.Concepto = dr["Concepto"].ToString();
                    movi.Cantidad = dr["Cantidad"].ToString();
                    movi.Unidad = dr["Unidad"].ToString();
                    movi.PrecioUnitario = dr["Precio Unitario"].ToString();
                    movi.ImporteLinea = dr["Importe Linea"].ToString();
                    movi.ImporteOrig = dr["Importe Orig"].ToString();
                    movi.ImporteModif = dr["Importe Modif"].ToString();
                    movi.MontoAjuste = dr["Monto Ajuste"].ToString();
                    movi.IVA = dr["IVA"].ToString();
                    movi.Total = dr["Total"].ToString();
                    movi.MontoLetra = dr["Monto Letra"].ToString();

                    addendiux.lista.Add(movi);
                }


                if (lDocto.cCodigoCliente != "")
                {
                    foreach (string x in addenda)
                    {
                        lDocto._Addendas.Add(x);

                    }
                    lDocto.addendiux = addendiux;
                    _RegDoctos.Add(lDocto);

                }

            }
            return "";

        }

        private int mGrabaEncabezadoComercial(RegDocto doc, int incluyedireccion, ref int aIdDocumento, ref long aFolio1, ref string aSerie, int conComercioExterior = 0, int grabacliente = 1)
        {
            int lret2 = 0;
            int lerrordocto = 0;
            StringBuilder sMensaje1 = new StringBuilder(512);
            string aCodigoConcepto = "";
            string ltextoextra1cliente = "";

            if (conComercioExterior == 1)
            {
                SqlCommand lsql = new SqlCommand();
                lsql.CommandText = "select ctextoextra1 from admClientes where ccodigocliente = '" + doc.cCodigoCliente + "'";
                lsql.Connection = miconexion._conexion1;

                SqlDataReader l;
                l = lsql.ExecuteReader();
                if (l.HasRows)
                {
                    l.Read();
                    ltextoextra1cliente = l["ctextoextra1"].ToString().Trim();
                }
                l.Close();
            }



            double aFolio = 0;
            if (doc.cFolio == 0)
            {
                try
                {
                    // int z = fSiguienteFolioComercial(doc.cCodigoConcepto, ref  aSerie, ref  aFolio);
                    aFolio1 = long.Parse(aFolio.ToString());
                    aFolio = -1;
                }
                catch (Exception ii)
                {
                }
            }
            else
            {
                if (doc.cFolio == -1)
                {
                    int z = fSiguienteFolioComercial(doc.cCodigoConcepto, ref aSerie, ref aFolio);
                    if (aSerie == null)
                        aSerie = "";
                    doc.cSerie = aSerie;

                }
                else
                    aFolio = doc.cFolio;
            }




            if (aFolio == 0)
            {
                aFolio = 1;
                aFolio1 = long.Parse(aFolio.ToString());
            }



            fInsertarDocumentoComercial();

            lret2 = fSetDatoDocumentoComercial("cCodigoConcepto", doc.cCodigoConcepto);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                //fProcesaError(doc, doc.cIdDocto, "El documento con cliente " + doc.cCodigoCliente.Trim() + " y folio " + doc.cFolio.ToString() + " presenta el sig. problema " + sMensaje1.ToString(), ref lret2);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }


            lret2 = fSetDatoDocumentoComercial("cCodigoCliente", doc.cCodigoCliente);

            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                mValidaClienteProveedor(doc, grabacliente);
                lret2 = fSetDatoDocumentoComercial("cCodigoCliente", doc.cCodigoCliente);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    return 0;
                }
            }
            else
            {
                int busca = fBuscaCteProvComercial(doc.cCodigoCliente);
                if (busca == 0)
                {
                    StringBuilder aValorRFC = new StringBuilder(20);
                    StringBuilder aValorRazonSocial = new StringBuilder(250);
                    lret2 = fLeeDatoDocumentoComercial("CRFC", aValorRFC, 20);
                    doc.cRFC = aValorRFC.ToString();
                    lret2 = fLeeDatoDocumentoComercial("CRAZONSOCIAL", aValorRazonSocial, 250);
                    doc._RegCliente.RazonSocial = aValorRazonSocial.ToString();
                }
            }
            lret2 = fSetDatoDocumentoComercial("cRazonSocial", doc._RegCliente.RazonSocial);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }
            lret2 = fSetDatoDocumentoComercial("cRFC", doc.cRFC);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }





            //lret2 = fSetDatoDocumentoComercial("cIdMoneda", "2");
            //if (lret2 != 0)
            //    fErrorComercial(lret2, sMensaje1, 512);

            //lret2 = fSetDatoDocumentoComercial("cTipoCambio", doc.cTipoCambio.ToString().Trim());
            //if (lret2 != 0)
            //    fErrorComercial(lret2, sMensaje1, 512);

            //DateTime lFechaVencimiento = DateTime.Today;
            string lfechavenc = String.Format("{0:MM/dd/yyyy}", DateTime.Today);
            lfechavenc = String.Format("{0:MM/dd/yyyy}", doc.cFecha);
            lret2 = fSetDatoDocumentoComercial("cFecha", lfechavenc);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }

            if (aFolio != -1)
            {
                lret2 = fSetDatoDocumentoComercial("cFolio", aFolio.ToString());
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    return 0;
                }
            }

            lret2 = fSetDatoDocumentoComercial("cSerieDocumento", doc.cSerie);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }

            lret2 = fSetDatoDocumentoComercial("cFechaVencimiento", lfechavenc);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }

            lret2 = fSetDatoDocumentoComercial("cCodigoCliente", doc.cCodigoCliente);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }

            RegCliente lc = new RegCliente();

            lc = mBuscarClienteComercial(doc.cCodigoCliente);



            lret2 = fSetDatoDocumentoComercial("CMETODOPAG", doc.cFormaPago);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }

            lret2 = fSetDatoDocumentoComercial("CREFERENCIA", doc.cReferencia);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }


            lret2 = fSetDatoDocumentoComercial("COBSERVACIONES", doc.cObservaciones);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }

            if (doc.cMetodoPago == "PPD")
            {
                lret2 = fSetDatoDocumentoComercial("CCANTPARCI", "2");
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    return 0;
                }
            }

            if (doc.cUsoCFDI != "")
            {
                lret2 = fSetDatoDocumentoComercial("CCODCONCBA", doc.cUsoCFDI);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    return 0;
                }
            }


            lret2 = fGuardaDocumentoComercial();
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }

            StringBuilder aValor = new StringBuilder(12);
            lret2 = fLeeDatoDocumentoComercial("CIDDOCUMENTO", aValor, 12);
            int liddocumento = int.Parse(aValor.ToString());


            lret2 = fLeeDatoDocumentoComercial("CFOLIO", aValor, 12);
            long llfolio = Convert.ToInt32(decimal.Parse(aValor.ToString()));

            lret2 = fLeeDatoDocumentoComercial("CSERIEDOCUMENTO", aValor, 12);
            string lSerie = aValor.ToString();
            aSerie = lSerie;

            doc.cFolio = llfolio;
            doc.cSerie = lSerie;


            long liddocumento1 = doc.cIdDocto;
            doc.cIdDocto = liddocumento;
            /*if (incluyedireccion == 1)
                lret2 = mGrabaDireccionComercial(doc);*/
            lret2 = mgrabamoneda(liddocumento, doc.cMoneda, doc.cTipoCambio);

            if (conComercioExterior == 1)
                lret2 = mgrabacomercioexterior(liddocumento, ltextoextra1cliente, doc.cTipoCambio);


            if (doc.cTextoExtra3 != "")
            {
                SqlCommand m = new SqlCommand();

                m.CommandText = "update admDocumentos set cTextoExtra3 = '" + doc.cTextoExtra3 + "' where ciddocumento = " + doc.cTextoExtra3;
                m.Connection = miconexion._conexion1;
                m.ExecuteNonQuery();

            }


            return 1;
        }




        public void mValidaClienteProveedor(RegDocto adocto, int grabacliente = 1)
        {
            StringBuilder aMensaje = new StringBuilder(512);
            int busca = fBuscaCteProvComercial(adocto.cCodigoCliente);
            if (busca != 0)
            {
                if (grabacliente == 1)
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
                //     else
                //   {
                // aMensaje.Append("Cliente " + adocto._RegCliente.Codigo + " no Existe");
                // fErrorComercial(busca, aMensaje, 512);
                //   }


            }

        }


        public long mGrabarDoctosComercialEncabezado()
        {
            string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            string aNombreEmpresa = "0000000";
            string aDirectorioEmpresa = "0000000000";
            int aIdEmpresa = 0;
            StringBuilder sMensaje1 = new StringBuilder(512);


            miconexion.mAbrirConexionComercial(true);
            int lResultado = fSetNombrePAQ("CONTPAQ I Comercial");
            if (lResultado != 0)
            {
                fErrorComercial(lResultado, sMensaje1, 512);
            }

            /*            if (incluyetimbrado == 1)
                        {
                            int lresp10 = fInicializaLicenseInfoComercial(0);
                            if (lresp10 != 0)
                            {
                                fErrorComercial(lresp10, sMensaje1, 512);
                            }
                        }*/

            int zzzzz = fAbreEmpresa(rutadestino);

            if (zzzzz != 0)
            {
                fErrorComercial(zzzzz, sMensaje1, 512);
            }

            RegDocto doc = _RegDoctoOrigen;
            int liddocumento = 0;
            long aFolio = _RegDoctoOrigen.cFolio;
            string aSerie = _RegDoctoOrigen.cSerie;
            liddocumento = 0;
            int lRetorno = mGrabaEncabezadoComercial(doc, 0, ref liddocumento, ref aFolio, ref aSerie, 0, 0);

            return doc.cIdDocto;




        }

        public long mGrabarSeriesPedidoComercial(long lidmovim, DataGridView migrid)
        {
            decimal ltotalunidadesdocto = 0;
            //int lRetorno = mGrabarMovimientosComercial(doc, 0, ref ltotalunidadesdocto, 0, 0);

            StringBuilder aValor = new StringBuilder(12);
            aValor.Length = 0;
            //long lidmovimiento = fLeeDatoMovimientoComercial("CIDMOVIMIENTO", aValor, 12);
            long lidmovimiento = lidmovim;

            SqlCommand lsql = new SqlCommand();
            lsql.Connection = miconexion._conexion1;

            if (lidmovimiento > 0)
            {
                lsql.CommandText = "delete admMovmientosSeriePedido where cidmovimiento = " + lidmovimiento.ToString();
                int lret1 = lsql.ExecuteNonQuery();


                foreach (DataGridViewRow x in migrid.Rows)
                {
                    if (x.Cells[2].Value != null)
                    {
                        string lFecha = DateTime.Now.Day.ToString() + "/" + DateTime.Now.Month + "/" + DateTime.Now.Year.ToString();

                        lFecha = DateTime.Now.Year.ToString() + "/" + DateTime.Now.Month + "/" + DateTime.Now.Day.ToString();

                        lsql.CommandText = "insert into admMovmientosSeriePedido values (" + lidmovimiento.ToString() + "," + x.Cells[2].Value + ",'" + lFecha + "')";
                        lret1 = lsql.ExecuteNonQuery();
                    }
                }
            }

            return lidmovimiento;
        }

        public long mGrabarMovimientoComercial(RegDocto doc, DataGridView migrid)
        {
            decimal ltotalunidadesdocto = 0;
            int lRetorno = mGrabarMovimientosComercial(doc, 0, ref ltotalunidadesdocto, 0, 0);

            StringBuilder aValor = new StringBuilder(12);
            aValor.Length = 0;
            long lidmovimiento = fLeeDatoMovimientoComercial("CIDMOVIMIENTO", aValor, 12);

            if (doc._RegMovtos[0].cIdMovto > 0)
            {
                SqlCommand lsql = new SqlCommand();
                lsql.Connection = miconexion._conexion1;

                foreach (DataGridViewRow x in migrid.Rows)
                {
                    string lFecha = DateTime.Now.Day.ToString() + "/" + DateTime.Now.Month + "/" + DateTime.Now.Year.ToString();

                    lFecha = DateTime.Now.Year.ToString() + "/" + DateTime.Now.Month + "/" + DateTime.Now.Day.ToString();

                    lsql.CommandText = "insert into admMovmientosSeriePedido values (" + doc._RegMovtos[0].cIdMovto.ToString() + "," + x.Cells[2].Value + ",'" + lFecha + "')";
                    int lret1 = lsql.ExecuteNonQuery();
                }
            }

            return lidmovimiento;
        }


        public string mGrabarDoctosComercial(int incluyetimbrado, ref long lultimoFolio, int incluyedireccion, int concomercioexterior, int grabarcliente = 1, int traspaso = 0)
        {
            StringBuilder sMensaje1 = new StringBuilder(512);

            string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            string aNombreEmpresa = "0000000";
            string aDirectorioEmpresa = "0000000000";
            int aIdEmpresa = 0;


            if (traspaso == 0)
            {
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

                //C:\\Compac\\Empresas\\aduno

                int zzzzz = fAbreEmpresa(rutadestino);

                if (zzzzz != 0)
                {
                    fErrorComercial(zzzzz, sMensaje1, 512);
                }
            }




            int indicedoc = 0;
            int lret2;
            int lcuantos = _RegDoctos.Count;
            int ltotales = lcuantos;
            int lindice = 1;
            int liddocumento = 0;
            decimal ltotalunidadesdocto = 0;
            long lfoliopedido = 0;
            long lidsalida = 0;
            long lidentrada = 0;

            double aFoliot = 0;

            foreach (RegDocto doc in _RegDoctos)
            {

                long aFolio = 0;
                aFoliot = 0;

                string aSerie = "";

                if (traspaso == 1)
                {

                    fSiguienteFolioComercial("36", ref aSerie, ref aFoliot);
                    doc.cFolio = long.Parse(aFoliot.ToString());
                    aFolio = 0;

                }
                long liddoc = doc.cIdDocto;
                double ltotalunidades = doc.cTotalUnidades;
                ltotalunidadesdocto = 0;
                if (lfoliopedido != 0)
                    doc.cFolio = lfoliopedido;
                int lRetorno = mGrabaEncabezadoComercial(doc, incluyedireccion, ref liddocumento, ref aFolio, ref aSerie, concomercioexterior, grabarcliente);
                if (lRetorno == 1)
                {
                    if (doc.cFolio == 0)
                    {
                        doc.cFolio = aFolio;
                        doc.cSerie = aSerie;
                    }
                    if (traspaso == 1)
                    {
                        if (doc.cCodigoConcepto == "35")
                            lidsalida = doc.cIdDocto;
                        if (doc.cCodigoConcepto == "34")
                            lidentrada = doc.cIdDocto;
                    }

                    lRetorno = mGrabarMovimientosComercial(doc, indicedoc, ref ltotalunidadesdocto, concomercioexterior, traspaso);
                    long lret4 = 0;

                    if (traspaso == 1)
                    {

                        if (ltotalunidadesdocto == 0)
                        {
                            lret4 = fBuscarDocumentoComercial("34", "", doc.cFolio.ToString());
                            if (lret4 == 0)

                                fBorraDocumentoComercial();

                            lret4 = fBuscarDocumentoComercial("35", "", doc.cFolio.ToString());
                            if (lret4 == 0)
                            {
                                fBorraDocumentoComercial();
                                lRetorno = 0;
                                traspaso = 2;
                                break;
                            }
                        }
                    }

                    if (concomercioexterior == 3) // autorizaciones
                    {
                        long liddocumento1 = doc.cIdDocto;
                        doc.cIdDocto = liddoc;
                        doc.cTotalUnidades = ltotalunidades;
                        lret2 = mgrabaorigen(doc, liddocumento1);
                        if (doc.cReferencia == "2")
                            lfoliopedido = doc.cFolio;
                    }
                    indicedoc++;


                }

                Notificar((double)(lindice++ * 100) / lcuantos);

                if (lRetorno == 1)
                {
                    mGrabarUnidadesDocto(doc.cIdDocto, ltotalunidadesdocto);

                    /*  if (traspaso == 10)
                      {
                          mGrabarTraspaso(doc.cIdDocto);


                      }*/

                    if (incluyetimbrado == 1 || incluyetimbrado == 3)
                    {

                        string lpass = "";
                        lpass = GetSettingValueFromAppConfigForDLL("Pass").ToString().Trim();

                        lultimoFolio = doc.cFolio;
                        double dfolio = 0.0;
                        dfolio = Convert.ToDouble(doc.cFolio);

                        int lresp20 = fEmitirDocumentoComercial(doc.cCodigoConcepto, doc.cSerie, dfolio, lpass, "");
                        if (lresp20 != 0)
                        {
                            fErrorComercial(lresp20, sMensaje1, 512);
                            fProcesaError(doc, null, "Doc", sMensaje1.ToString(), 0);


                        }
                        else
                        {
                            //lresp20 = fEntregEnDiscoXMLComercial(doc.cCodigoConcepto, doc.cSerie, doc.cFolio, 0, @"C:\Compac\Empresas\Reportes\Formatos Digitales\reportes_Servidor\COMERCIAL\Factura.rdl");
                            lresp20 = fEntregEnDiscoXMLComercial(doc.cCodigoConcepto, doc.cSerie, doc.cFolio, 0, "");
                            if (lresp20 != 0)
                            {
                                fErrorComercial(lresp20, sMensaje1, 512);
                                fProcesaError(doc, null, "Doc", sMensaje1.ToString(), 0);
                            }
                            if (doc.cNombreArchivo != "")
                            {
                                string archivoorigencompleto = @GetSettingValueFromAppConfigForDLL("RutaOrigen").ToString().Trim() + @"\" + doc.cNombreArchivo;
                                string archivodestinocompleto = @GetSettingValueFromAppConfigForDLL("RutaBien").ToString().Trim() + @"\" + doc.cNombreArchivo;
                                try
                                {
                                    System.IO.File.Move(archivoorigencompleto, archivodestinocompleto);
                                }
                                catch (Exception aaaa)
                                { }


                            }
                        }
                    }
                    if (incluyetimbrado == 2)
                    {
                        // grabar en al tabla de folios digitales
                        string ltexto = "select cidconceptodocumento from admConceptos where ccodigoconcepto = '" + doc.cCodigoConcepto + "'";
                        int lidconc = mRegresaId(ltexto, miconexion._conexion1);
                        mFolioDigital1(doc, lidconc, "0");
                    }


                    //lexitosos++;
                }
                //else
                //  fBorraDocumentoComercial();
                //Notificar((double)(lindice++ * 100) / lcuantos);

            }


            if (traspaso == 1)
            {
                if (lidsalida != 0 && lidentrada != 0)
                {
                    mGrabarTraspaso(lidsalida, lidentrada, aFoliot);
                    lultimoFolio = long.Parse(aFoliot.ToString());
                }



            }







            fCierraEmpresa();

            return "";

        }




        public int mEjecutarComando(string comando, int aClientes, int aporcodigo)
        {
            miconexion.mAbrirConexionOrigen();
            //string lcadena44 = "update mgw10010 set cneto = " + x.cSubtotal + ", cimpuesto1 = " + x.cImpuesto + ", ctotal = " + total + " where ciddocum01 = " + lIdDocumento;





            OleDbDataAdapter lda = new OleDbDataAdapter(comando, miconexion._conexion);
            System.Data.DataSet lds = new System.Data.DataSet();
            lda.Fill(lds);
            miconexion.mCerrarConexionOrigen();

            miconexion.mAbrirConexionDestino();

            OleDbCommand com = new OleDbCommand("select NVL(max(ciddocum01),0) from mgw10008");
            OleDbDataReader ldr;

            //com.CommandText = "select NVL(max(ciddocum01),1) from mgw10008";
            com.Connection = miconexion._conexion;
            ldr = com.ExecuteReader();
            ldr.Read();
            int liddocum = int.Parse(ldr[0].ToString()) + 1;

            OleDbCommand com1 = new OleDbCommand("select NVL(max(cidmovim01),0) from mgw10010");
            OleDbDataReader ldr1;

            //com1.CommandText = "select NVL(max(ciddocum01),1) from mgw10008";
            com1.Connection = miconexion._conexion;
            ldr1 = com1.ExecuteReader();
            ldr1.Read();
            int lidmovim = int.Parse(ldr1[0].ToString()) + 1;

            foreach (DataRow zz in lds.Tables[0].Rows)
            {

                string ltexto = "";
                if (aporcodigo == 1)
                    ltexto = "select cidclien01 from mgw10002 where ccodigoc01 = '" + zz["ccodigoc01"].ToString() + "'";
                else
                    ltexto = "select cidclien01 from mgw10002 where ctextoex01 = '" + zz["ccodigoc01"].ToString() + "'";

                OleDbCommand com2 = new OleDbCommand(ltexto);

                OleDbDataReader ldr2;
                com2.Connection = miconexion._conexion;
                ldr2 = com2.ExecuteReader();
                int lidclien = 0;
                if (ldr2.HasRows == true)
                {
                    ldr2.Read();
                    lidclien = int.Parse(ldr2[0].ToString());
                }
                if (lidclien > 0)
                {
                    StringBuilder x = new StringBuilder();
                    x.AppendLine("insert into mgw10008 values(");
                    //lds.Tables[0].Rows["ciddocum01"].ToString();
                    x.AppendLine(liddocum + ",");
                    if (aClientes == 1)
                    {
                        //x.AppendLine(zz["ciddocum02"].ToString() + ",");
                        x.AppendLine("13,");
                        x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        x.AppendLine("40,");
                    }
                    //x.AppendLine(zz["cidconce01"].ToString() + ",");


                    x.AppendLine("'" + zz["cseriedo01"].ToString() + "',");
                    //x.AppendLine("'ABCD',");

                    x.AppendLine(zz["cfolio"].ToString() + ",");
                    //x.AppendLine(liddocum.ToString() + ",");


                    string fecha = zz["cfecha"].ToString().Substring(0, 10);

                    DateTime dfecha = DateTime.Parse(fecha);
                    fecha = dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');

                    x.AppendLine("ctod('" + fecha + "'),");
                    //x.AppendLine(zz["cidclien01"].ToString() + ",");

                    x.AppendLine(lidclien.ToString() + ",");


                    x.AppendLine("'" + zz["crazonso01"].ToString() + "',");
                    x.AppendLine("'" + zz["crfc"].ToString() + "',");
                    x.AppendLine(zz["cidagente"].ToString() + ",");

                    string fechav = zz["cfechave01"].ToString().Substring(0, 10);

                    DateTime dfechav = DateTime.Parse(fechav);
                    fechav = dfechav.Month.ToString().PadLeft(2, '0') + "/" + dfechav.Day.ToString().PadLeft(2, '0') + "/" + dfechav.Year.ToString().PadLeft(4, '0');
                    x.AppendLine("ctod('" + fechav + "'),");
                    x.AppendLine("ctod('" + zz["cfechapr01"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("ctod('" + zz["cfechaen01"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("ctod('" + zz["cfechaul01"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine(zz["cidmoneda"].ToString() + ",");
                    x.AppendLine(zz["ctipocam01"].ToString() + ",");
                    x.AppendLine("'" + zz["creferen01"].ToString() + "',");
                    //x.AppendLine("'" + zz["cobserva01"].ToString() + "',");
                    x.AppendLine("'" + zz["cfolio"].ToString() + "',");

                    x.AppendLine(zz["cnatural01"].ToString() + ",");
                    x.AppendLine(zz["ciddocum03"].ToString() + ",");
                    x.AppendLine(zz["cplantilla"].ToString() + ",");
                    x.AppendLine(zz["cusaclie01"].ToString() + ",");
                    x.AppendLine(zz["cusaprov01"].ToString() + ",");
                    x.AppendLine(zz["cafectado"].ToString() + ",");
                    //x.AppendLine(zz["cimpreso"].ToString() + ",");
                    x.AppendLine("0,");

                    x.AppendLine(zz["ccancelado"].ToString() + ",");
                    x.AppendLine(zz["cdevuelto"].ToString() + ",");
                    x.AppendLine(zz["cidprepo01"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["cidprepo02"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["cestadoc01"].ToString() + ",");
                    //x.AppendLine(zz["cneto"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    //x.AppendLine(zz["cimpuesto1"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cimpuesto2"].ToString() + ",");
                    x.AppendLine(zz["cimpuesto3"].ToString() + ",");
                    x.AppendLine(zz["cretenci01"].ToString() + ",");
                    x.AppendLine(zz["cretenci02"].ToString() + ",");
                    x.AppendLine(zz["cdescuen01"].ToString() + ",");
                    x.AppendLine(zz["cdescuen02"].ToString() + ",");
                    x.AppendLine(zz["cdescuen03"].ToString() + ",");
                    x.AppendLine(zz["cgasto1"].ToString() + ",");
                    x.AppendLine(zz["cgasto2"].ToString() + ",");
                    x.AppendLine(zz["cgasto3"].ToString() + ",");
                    //x.AppendLine(zz["ctotal"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    //x.AppendLine(zz["ctotalun01"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cdescuen04"].ToString() + ",");
                    x.AppendLine(zz["cporcent01"].ToString() + ",");
                    x.AppendLine(zz["cporcent02"].ToString() + ",");
                    x.AppendLine(zz["cporcent03"].ToString() + ",");
                    x.AppendLine(zz["cporcent04"].ToString() + ",");
                    x.AppendLine(zz["cporcent05"].ToString() + ",");
                    x.AppendLine(zz["cporcent06"].ToString() + ",");
                    x.AppendLine("'" + zz["ctextoex01"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoex02"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoex03"].ToString() + "',");
                    //x.AppendLine("'" + zz["cfechaex01"].ToString() + "',");
                    x.AppendLine("ctod('" + zz["cfechaex01"].ToString().Substring(0, 10) + "'),");


                    //x.AppendLine(zz["cporcent06"].ToString() + ",");
                    x.AppendLine(zz["cimporte01"].ToString() + ",");

                    x.AppendLine(zz["cimporte02"].ToString() + ",");
                    x.AppendLine(zz["cimporte03"].ToString() + ",");
                    x.AppendLine(zz["cimporte04"].ToString() + ",");
                    x.AppendLine("'" + zz["cdestina01"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumerog01"].ToString() + "',");
                    x.AppendLine("'" + zz["cmensaje01"].ToString() + "',");
                    x.AppendLine("'" + zz["ccuentam01"].ToString() + "',");
                    x.AppendLine(zz["cnumeroc01"].ToString() + ",");

                    x.AppendLine(zz["cpeso"].ToString() + ",");
                    x.AppendLine(zz["cbanobse01"].ToString() + ",");
                    x.AppendLine(zz["cbandato01"].ToString() + ",");
                    x.AppendLine(zz["cbancond01"].ToString() + ",");
                    x.AppendLine(zz["cbangastos"].ToString() + ",");
                    x.AppendLine(zz["cunidade01"].ToString() + ",");
                    x.AppendLine("ctod('" + zz["ctimestamp"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine(zz["cimpcheq01"].ToString() + ",");
                    x.AppendLine(zz["csistorig"].ToString() + ",");
                    x.AppendLine(zz["cidmonedca"].ToString() + ",");
                    x.AppendLine(zz["ctipocamca"].ToString() + ",");
                    //x.AppendLine(zz["cescfd"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["ctienecfd"].ToString() + ",");
                    x.AppendLine("'" + zz["clugarexpe"].ToString() + "',");
                    x.AppendLine("'" + zz["cmetodopag"].ToString() + "',");
                    x.AppendLine(zz["cnumparcia"].ToString() + ",");
                    x.AppendLine(zz["ccantparci"].ToString() + ",");
                    x.AppendLine("'" + zz["ccondipago"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumctapag"].ToString() + "')");

                    comando = x.ToString();
                    OleDbCommand lsql3 = new OleDbCommand(comando, miconexion._conexion);
                    lsql3.ExecuteNonQuery();


                    x = new StringBuilder();
                    x.AppendLine("insert into mgw10010 values(");
                    //lds.Tables[0].Rows["ciddocum01"].ToString();
                    x.AppendLine(lidmovim.ToString() + ",");
                    x.AppendLine(liddocum.ToString() + ",");
                    x.AppendLine("1,");
                    //x.AppendLine("13,");

                    if (aClientes == 1)
                    {
                        //x.AppendLine(zz["ciddocum02"].ToString() + ",");
                        x.AppendLine("13,");
                        //  x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        //x.AppendLine("40,");
                    }

                    x.AppendLine("0,"); //producto
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //unidades2
                    x.AppendLine("0,"); //cidunidad
                    x.AppendLine("0,"); //cidunida01
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    //x.AppendLine(zz["cneto"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");

                    //x.AppendLine(zz["cimpuesto1"].ToString() + ",");
                    x.AppendLine("0,"); //impuesto1
                    x.AppendLine(zz["cporcent01"].ToString() + ",");
                    x.AppendLine("0,"); //impuesto2
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //impuesto3
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //retencion1
                    x.AppendLine("0,");
                    x.AppendLine("0,");//retencion2
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 1
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 2
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 3
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 4
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 5
                    x.AppendLine("0,");
                    //x.AppendLine(zz["ctotal"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine("'',");
                    x.AppendLine("'',"); //observaciones
                    x.AppendLine("3,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    //x.AppendLine("3,");

                    x.AppendLine("ctod('" + zz["cfecha"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("0,0,0,0,0,0,0,");
                    x.AppendLine("0,0,'','','',");
                    x.AppendLine("ctod('" + fecha + "'),");
                    x.AppendLine("0,0,0,0,");
                    x.AppendLine("'',0,'',0,0,0)");

                    comando = x.ToString();
                    OleDbCommand lsql4 = new OleDbCommand(comando, miconexion._conexion);
                    lsql4.ExecuteNonQuery();

                    liddocum++;
                    lidmovim++;
                }






                /*
                 INSERT INTO [DESTINO]...[mgw10010]
               ([cidmovim01]
               ,[ciddocum01]
               ,[cnumerom01]
               ,[ciddocum02]
               ,[cidprodu01]
               ,[cidalmacen]
               ,[cunidades]
               ,[cunidade01]
               ,[cunidade02]
               ,[cidunidad]
               ,[cidunida01]
               ,[cprecio]
               ,[cprecioc01]
               ,[ccostoca01]
               ,[ccostoes01]
               ,[cneto]
               ,[cimpuesto1]
               ,[cporcent01]
               ,[cimpuesto2]
               ,[cporcent02]
               ,[cimpuesto3]
               ,[cporcent03]
               ,[cretenci01]
               ,[cporcent04]
               ,[cretenci02]
               ,[cporcent05]
               ,[cdescuen01]
               ,[cporcent06]
               ,[cdescuen02]
               ,[cporcent07]
               ,[cdescuen03]
               ,[cporcent08]
               ,[cdescuen04]
               ,[cporcent09]
               ,[cdescuen05]
               ,[cporcent10]
               ,[ctotal]
               ,[cporcent11]
               ,[creferen01]
               ,[cobserva01]
               ,[cafectae01]
               ,[cafectad01]
               ,[cafectad02]
               ,[cfecha]
               ,[cmovtooc01]
               ,[cidmovto01]
               ,[cidmovto02]
               ,[cunidade03]
               ,[cunidade04]
               ,[cunidade05]
               ,[cunidade06]
               ,[ctipotra01]
               ,[cidvalor01]
               ,[ctextoex01]
               ,[ctextoex02]
               ,[ctextoex03]
               ,[cfechaex01]
               ,[cimporte01]
               ,[cimporte02]
               ,[cimporte03]
               ,[cimporte04]
               ,[ctimestamp]
               ,[cgtomovto]
               ,[cscmovto]
               ,[ccomventa]
               ,[cidmovdest]
               ,[cnumconsol])

                 */


            }
            miconexion.mCerrarConexionDestino();

            return 0;
        }

        public int mEjecutarComando2(string comando, int aClientes, int aporcodigo, string empresa)
        {
            miconexion.mAbrirConexionOrigen();



            //mValidaSQLConexion(txtServer.Text, txtBD.Text, txtUser.Text, txtPass.Text);


            OleDbDataAdapter lda = new OleDbDataAdapter(comando, miconexion._conexion);
            System.Data.DataSet lds = new System.Data.DataSet();
            lda.Fill(lds);
            miconexion.mCerrarConexionOrigen();



            string sempresa = empresa.Substring(empresa.LastIndexOf("\\") + 1);

            string Cadenaconexion1 = "data source =" + cserver + ";initial catalog = " + sempresa + ";user id = " + cusr + "; password = " + cpwd + ";";
            SqlConnection _con = new SqlConnection();

            _con.ConnectionString = Cadenaconexion1;

            //miconexion.mAbrirConexionDestino();

            _con.Open();
            SqlCommand com = new SqlCommand("select ISNULL(max(ciddocumento),0) from admDocumentos");
            SqlDataReader ldr;

            //com.CommandText = "select NVL(max(ciddocum01),1) from mgw10008";
            com.Connection = _con;
            ldr = com.ExecuteReader();
            ldr.Read();
            int liddocum = int.Parse(ldr[0].ToString()) + 1;

            SqlCommand com1 = new SqlCommand("select ISNULL(max(ciddocumento),0) from admMovimientos");
            SqlDataReader ldr1;

            //com1.CommandText = "select NVL(max(ciddocum01),1) from mgw10008";
            com1.Connection = _con;
            ldr.Close();
            ldr1 = com1.ExecuteReader();
            ldr1.Read();
            int lidmovim = int.Parse(ldr1[0].ToString()) + 1;
            ldr1.Close();
            foreach (DataRow zz in lds.Tables[0].Rows)
            {

                string ltexto = "";
                if (aporcodigo == 1)
                    ltexto = "select cidclienteproveedor from admClientes where ccodigocliente = '" + zz["ccodigoc01"].ToString() + "'";
                else
                    ltexto = "select cidclienteproveedor from admClientes where ctextoextra1 = '" + zz["ccodigoc01"].ToString() + "'";

                SqlCommand com2 = new SqlCommand(ltexto);

                SqlDataReader ldr2;
                com2.Connection = _con;
                ldr2 = com2.ExecuteReader();
                int lidclien = 0;
                if (ldr2.HasRows == true)
                {
                    ldr2.Read();
                    lidclien = int.Parse(ldr2[0].ToString());
                }
                ldr2.Close();
                if (lidclien > 0)
                {
                    StringBuilder x = new StringBuilder();
                    x.AppendLine("insert into admDocumentos values(");
                    //lds.Tables[0].Rows["ciddocum01"].ToString();
                    x.AppendLine(liddocum + ",");
                    if (aClientes == 1)
                    {
                        //x.AppendLine(zz["ciddocum02"].ToString() + ",");
                        x.AppendLine("13,");
                        x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        x.AppendLine("40,");
                    }
                    //x.AppendLine(zz["cidconce01"].ToString() + ",");


                    x.AppendLine("'" + zz["cseriedo01"].ToString() + "',");
                    //x.AppendLine("'ABCD',");

                    x.AppendLine(zz["cfolio"].ToString() + ",");
                    //x.AppendLine(liddocum.ToString() + ",");


                    string fecha = zz["cfecha"].ToString().Substring(0, 10);
                    int espacio = fecha.IndexOf(" ");
                    if (espacio > -1)
                        fecha = fecha.Substring(0, espacio);
                    DateTime dfecha = DateTime.Parse(fecha);
                    fecha = dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                    string sFecha = dfecha.Year.ToString().PadLeft(4, '0') + dfecha.Month.ToString().PadLeft(2, '0') + dfecha.Day.ToString().PadLeft(2, '0');
                    x.AppendLine("'" + sFecha + "',");
                    //x.AppendLine(zz["cidclien01"].ToString() + ",");

                    x.AppendLine(lidclien.ToString() + ",");


                    x.AppendLine("'" + zz["crazonso01"].ToString() + "',");
                    x.AppendLine("'" + zz["crfc"].ToString() + "',");
                    x.AppendLine(zz["cidagente"].ToString() + ",");

                    string fechav = zz["cfechave01"].ToString().Substring(0, 10);
                    int espacio1 = fechav.IndexOf(" ");
                    if (espacio1 > -1)
                        fechav = fechav.Substring(0, espacio1);

                    DateTime dfechav = DateTime.Parse(fechav);
                    fechav = dfechav.Month.ToString().PadLeft(2, '0') + "/" + dfechav.Day.ToString().PadLeft(2, '0') + "/" + dfechav.Year.ToString().PadLeft(4, '0');
                    string sFechav = dfechav.Year.ToString().PadLeft(4, '0') + dfechav.Month.ToString().PadLeft(2, '0') + dfechav.Day.ToString().PadLeft(2, '0');
                    x.AppendLine("'" + sFechav + "',");
                    //x.AppendLine("ctod('" + zz["cfechapr01"].ToString().Substring(0, 10) + "'),");
                    //x.AppendLine("ctod('" + zz["cfechaen01"].ToString().Substring(0, 10) + "'),");
                    //x.AppendLine("ctod('" + zz["cfechaul01"].ToString().Substring(0, 10) + "'),");

                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("'" + sFechav + "',");

                    x.AppendLine(zz["cidmoneda"].ToString() + ",");
                    x.AppendLine(zz["ctipocam01"].ToString() + ",");
                    x.AppendLine("'" + zz["creferen01"].ToString() + "',");
                    //x.AppendLine("'" + zz["cobserva01"].ToString() + "',");
                    x.AppendLine("'" + zz["cfolio"].ToString() + "',");

                    x.AppendLine(zz["cnatural01"].ToString() + ",");
                    x.AppendLine(zz["ciddocum03"].ToString() + ",");
                    x.AppendLine(zz["cplantilla"].ToString() + ",");
                    x.AppendLine(zz["cusaclie01"].ToString() + ",");
                    x.AppendLine(zz["cusaprov01"].ToString() + ",");
                    x.AppendLine(zz["cafectado"].ToString() + ",");
                    //x.AppendLine(zz["cimpreso"].ToString() + ",");
                    x.AppendLine("0,");

                    x.AppendLine(zz["ccancelado"].ToString() + ",");
                    x.AppendLine(zz["cdevuelto"].ToString() + ",");
                    x.AppendLine(zz["cidprepo01"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["cidprepo02"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["cestadoc01"].ToString() + ",");
                    //x.AppendLine(zz["cneto"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    //x.AppendLine(zz["cimpuesto1"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cimpuesto2"].ToString() + ",");
                    x.AppendLine(zz["cimpuesto3"].ToString() + ",");
                    x.AppendLine(zz["cretenci01"].ToString() + ",");
                    x.AppendLine(zz["cretenci02"].ToString() + ",");
                    x.AppendLine(zz["cdescuen01"].ToString() + ",");
                    x.AppendLine(zz["cdescuen02"].ToString() + ",");
                    x.AppendLine(zz["cdescuen03"].ToString() + ",");
                    x.AppendLine(zz["cgasto1"].ToString() + ",");
                    x.AppendLine(zz["cgasto2"].ToString() + ",");
                    x.AppendLine(zz["cgasto3"].ToString() + ",");
                    //x.AppendLine(zz["ctotal"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    //x.AppendLine(zz["ctotalun01"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cdescuen04"].ToString() + ",");
                    x.AppendLine(zz["cporcent01"].ToString() + ",");
                    x.AppendLine(zz["cporcent02"].ToString() + ",");
                    x.AppendLine(zz["cporcent03"].ToString() + ",");
                    x.AppendLine(zz["cporcent04"].ToString() + ",");
                    x.AppendLine(zz["cporcent05"].ToString() + ",");
                    x.AppendLine(zz["cporcent06"].ToString() + ",");
                    x.AppendLine("'" + zz["ctextoex01"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoex02"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoex03"].ToString() + "',");
                    //x.AppendLine("'" + zz["cfechaex01"].ToString() + "',");
                    //x.AppendLine("ctod('" + zz["cfechaex01"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("'" + sFechav + "',");


                    //x.AppendLine(zz["cporcent06"].ToString() + ",");
                    x.AppendLine(zz["cimporte01"].ToString() + ",");

                    x.AppendLine(zz["cimporte02"].ToString() + ",");
                    x.AppendLine(zz["cimporte03"].ToString() + ",");
                    x.AppendLine(zz["cimporte04"].ToString() + ",");
                    x.AppendLine("'" + zz["cdestina01"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumerog01"].ToString() + "',");
                    x.AppendLine("'" + zz["cmensaje01"].ToString() + "',");
                    x.AppendLine("'" + zz["ccuentam01"].ToString() + "',");
                    x.AppendLine(zz["cnumeroc01"].ToString() + ",");

                    x.AppendLine(zz["cpeso"].ToString() + ",");
                    x.AppendLine(zz["cbanobse01"].ToString() + ",");
                    x.AppendLine(zz["cbandato01"].ToString() + ",");
                    x.AppendLine(zz["cbancond01"].ToString() + ",");
                    x.AppendLine(zz["cbangastos"].ToString() + ",");
                    x.AppendLine(zz["cunidade01"].ToString() + ",");
                    //x.AppendLine("ctod('" + zz["ctimestamp"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("'" + sFechav + "',");

                    x.AppendLine(zz["cimpcheq01"].ToString() + ",");
                    x.AppendLine(zz["csistorig"].ToString() + ",");
                    x.AppendLine(zz["cidmonedca"].ToString() + ",");
                    x.AppendLine(zz["ctipocamca"].ToString() + ",");
                    //x.AppendLine(zz["cescfd"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["ctienecfd"].ToString() + ",");
                    x.AppendLine("'" + zz["clugarexpe"].ToString() + "',");
                    x.AppendLine("'" + zz["cmetodopag"].ToString() + "',");
                    x.AppendLine(zz["cnumparcia"].ToString() + ",");
                    x.AppendLine(zz["ccantparci"].ToString() + ",");
                    x.AppendLine("'" + zz["ccondipago"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumctapag"].ToString() + "'");
                    x.AppendLine(",NEWID(),'',0)");


                    comando = x.ToString();
                    SqlCommand lsql3 = new SqlCommand(comando, _con);
                    lsql3.ExecuteNonQuery();


                    x = new StringBuilder();
                    x.AppendLine("insert into admmovimientos values(");
                    //lds.Tables[0].Rows["ciddocum01"].ToString();
                    x.AppendLine(lidmovim.ToString() + ",");
                    x.AppendLine(liddocum.ToString() + ",");
                    x.AppendLine("1,");
                    //x.AppendLine("13,");

                    if (aClientes == 1)
                    {
                        //x.AppendLine(zz["ciddocum02"].ToString() + ",");
                        x.AppendLine("13,");
                        //  x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        //x.AppendLine("40,");
                    }

                    x.AppendLine("0,"); //producto
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //unidades2
                    x.AppendLine("0,"); //cidunidad
                    x.AppendLine("0,"); //cidunida01
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    //x.AppendLine(zz["cneto"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");

                    //x.AppendLine(zz["cimpuesto1"].ToString() + ",");
                    x.AppendLine("0,"); //impuesto1
                    x.AppendLine(zz["cporcent01"].ToString() + ",");
                    x.AppendLine("0,"); //impuesto2
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //impuesto3
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //retencion1
                    x.AppendLine("0,");
                    x.AppendLine("0,");//retencion2
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 1
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 2
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 3
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 4
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 5
                    x.AppendLine("0,");
                    //x.AppendLine(zz["ctotal"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine("'',");
                    x.AppendLine("'',"); //observaciones
                    x.AppendLine("3,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    //x.AppendLine("3,");

                    //x.AppendLine("ctod('" + zz["cfecha"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("0,0,0,0,0,0,0,");
                    x.AppendLine("0,0,'','','',");
                    //x.AppendLine("ctod('" + fecha + "'),");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("0,0,0,0,");
                    x.AppendLine("'',0,'',0,0,0)");

                    comando = x.ToString();
                    SqlCommand lsql4 = new SqlCommand(comando, _con);
                    lsql4.ExecuteNonQuery();

                    liddocum++;
                    lidmovim++;
                }






                /*
                 INSERT INTO [DESTINO]...[mgw10010]
               ([cidmovim01]
               ,[ciddocum01]
               ,[cnumerom01]
               ,[ciddocum02]
               ,[cidprodu01]
               ,[cidalmacen]
               ,[cunidades]
               ,[cunidade01]
               ,[cunidade02]
               ,[cidunidad]
               ,[cidunida01]
               ,[cprecio]
               ,[cprecioc01]
               ,[ccostoca01]
               ,[ccostoes01]
               ,[cneto]
               ,[cimpuesto1]
               ,[cporcent01]
               ,[cimpuesto2]
               ,[cporcent02]
               ,[cimpuesto3]
               ,[cporcent03]
               ,[cretenci01]
               ,[cporcent04]
               ,[cretenci02]
               ,[cporcent05]
               ,[cdescuen01]
               ,[cporcent06]
               ,[cdescuen02]
               ,[cporcent07]
               ,[cdescuen03]
               ,[cporcent08]
               ,[cdescuen04]
               ,[cporcent09]
               ,[cdescuen05]
               ,[cporcent10]
               ,[ctotal]
               ,[cporcent11]
               ,[creferen01]
               ,[cobserva01]
               ,[cafectae01]
               ,[cafectad01]
               ,[cafectad02]
               ,[cfecha]
               ,[cmovtooc01]
               ,[cidmovto01]
               ,[cidmovto02]
               ,[cunidade03]
               ,[cunidade04]
               ,[cunidade05]
               ,[cunidade06]
               ,[ctipotra01]
               ,[cidvalor01]
               ,[ctextoex01]
               ,[ctextoex02]
               ,[ctextoex03]
               ,[cfechaex01]
               ,[cimporte01]
               ,[cimporte02]
               ,[cimporte03]
               ,[cimporte04]
               ,[ctimestamp]
               ,[cgtomovto]
               ,[cscmovto]
               ,[ccomventa]
               ,[cidmovdest]
               ,[cnumconsol])

                 */


            }
            miconexion.mCerrarConexionDestino();

            return 0;
        }



        public int mEjecutarComando3(string comando, int aClientes, int aporcodigo, string empresaorigen, string sempresadestino)
        {
            //miconexion.mAbrirConexionOrigen();
            SqlConnection _conOrigen = new SqlConnection();
            SqlConnection _conDestino = new SqlConnection();
            string sempresa = empresaorigen.Substring(empresaorigen.LastIndexOf("\\") + 1);
            string CadenaconexionOrigen = "data source =" + cserver + ";initial catalog = " + sempresa + ";user id = " + cusr + "; password = " + cpwd + ";";
            SqlConnection _con = new SqlConnection();

            _conOrigen.ConnectionString = CadenaconexionOrigen;
            _conOrigen.Open();





            /*OleDbDataAdapter lda = new OleDbDataAdapter(comando, miconexion._conexion);*/
            SqlDataAdapter lda = new SqlDataAdapter(comando, _conOrigen);
            System.Data.DataSet lds = new System.Data.DataSet();
            lda.Fill(lds);
            /*miconexion.mCerrarConexionOrigen();*/



            sempresa = sempresadestino.Substring(sempresadestino.LastIndexOf("\\") + 1);

            string Cadenaconexion1 = "data source =" + cserver + ";initial catalog = " + sempresa + ";user id = " + cusr + "; password = " + cpwd + ";";
            SqlConnection _conDes = new SqlConnection();

            _conDes.ConnectionString = Cadenaconexion1;

            //miconexion.mAbrirConexionDestino();

            _conDes.Open();
            SqlCommand com = new SqlCommand("select ISNULL(max(ciddocumento),0) from admDocumentos");
            SqlDataReader ldr;

            //com.CommandText = "select NVL(max(ciddocum01),1) from mgw10008";
            com.Connection = _conDes;
            ldr = com.ExecuteReader();
            ldr.Read();
            int liddocum = int.Parse(ldr[0].ToString()) + 1;

            SqlCommand com1 = new SqlCommand("select ISNULL(max(ciddocumento),0) from admMovimientos");
            SqlDataReader ldr1;

            //com1.CommandText = "select NVL(max(ciddocum01),1) from mgw10008";
            com1.Connection = _conDes;
            ldr.Close();
            ldr1 = com1.ExecuteReader();
            ldr1.Read();
            int lidmovim = int.Parse(ldr1[0].ToString()) + 1;
            ldr1.Close();
            foreach (DataRow zz in lds.Tables[0].Rows)
            {

                string ltexto = "";
                if (aporcodigo == 1)
                    ltexto = "select cidclienteproveedor from admClientes where ccodigocliente = '" + zz["ccodigocliente"].ToString() + "'";
                else
                    ltexto = "select cidclienteproveedor from admClientes where ctextoextra1 = '" + zz["ccodigoccliente"].ToString() + "'";

                SqlCommand com2 = new SqlCommand(ltexto);

                SqlDataReader ldr2;
                com2.Connection = _conDes;
                ldr2 = com2.ExecuteReader();
                int lidclien = 0;
                if (ldr2.HasRows == true)
                {
                    ldr2.Read();
                    lidclien = int.Parse(ldr2[0].ToString());
                }
                ldr2.Close();
                if (lidclien > 0)
                {
                    StringBuilder x = new StringBuilder();
                    x.AppendLine("insert into admDocumentos values(");
                    //lds.Tables[0].Rows["ciddocum01"].ToString();
                    x.AppendLine(liddocum + ",");
                    if (aClientes == 1)
                    {
                        //x.AppendLine(zz["ciddocum02"].ToString() + ",");
                        x.AppendLine("13,");
                        x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        x.AppendLine("40,");
                    }
                    //x.AppendLine(zz["cidconce01"].ToString() + ",");


                    x.AppendLine("'" + zz["cseriedocumento"].ToString() + "',");
                    //x.AppendLine("'ABCD',");

                    x.AppendLine(zz["cfolio"].ToString() + ",");
                    //x.AppendLine(liddocum.ToString() + ",");


                    string fecha = zz["cfecha"].ToString().Substring(0, 10);
                    int espacio = fecha.IndexOf(" ");
                    if (espacio > -1)
                        fecha = fecha.Substring(0, espacio);
                    DateTime dfecha = DateTime.Parse(fecha);
                    fecha = dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                    string sFecha = dfecha.Year.ToString().PadLeft(4, '0') + dfecha.Month.ToString().PadLeft(2, '0') + dfecha.Day.ToString().PadLeft(2, '0');
                    x.AppendLine("'" + sFecha + "',");
                    //x.AppendLine(zz["cidclien01"].ToString() + ",");

                    x.AppendLine(lidclien.ToString() + ",");


                    x.AppendLine("'" + zz["crazonsocial"].ToString() + "',");
                    x.AppendLine("'" + zz["crfc"].ToString() + "',");
                    x.AppendLine(zz["cidagente"].ToString() + ",");

                    string fechav = zz["cfechavencimiento"].ToString().Substring(0, 10);
                    int espacio1 = fechav.IndexOf(" ");
                    if (espacio1 > -1)
                        fechav = fechav.Substring(0, espacio1);

                    DateTime dfechav = DateTime.Parse(fechav);
                    fechav = dfechav.Month.ToString().PadLeft(2, '0') + "/" + dfechav.Day.ToString().PadLeft(2, '0') + "/" + dfechav.Year.ToString().PadLeft(4, '0');
                    string sFechav = dfechav.Year.ToString().PadLeft(4, '0') + dfechav.Month.ToString().PadLeft(2, '0') + dfechav.Day.ToString().PadLeft(2, '0');
                    x.AppendLine("'" + sFechav + "',");
                    //x.AppendLine("ctod('" + zz["cfechapr01"].ToString().Substring(0, 10) + "'),");
                    //x.AppendLine("ctod('" + zz["cfechaen01"].ToString().Substring(0, 10) + "'),");
                    //x.AppendLine("ctod('" + zz["cfechaul01"].ToString().Substring(0, 10) + "'),");

                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("'" + sFechav + "',");

                    x.AppendLine(zz["cidmoneda"].ToString() + ",");
                    x.AppendLine(zz["ctipocambio"].ToString() + ",");
                    x.AppendLine("'" + zz["creferencia"].ToString() + "',");
                    //x.AppendLine("'" + zz["cobserva01"].ToString() + "',");
                    x.AppendLine("'" + zz["cfolio"].ToString() + "',");

                    x.AppendLine(zz["cnaturaleza"].ToString() + ",");
                    x.AppendLine(zz["ciddocumentoorigen"].ToString() + ",");
                    x.AppendLine(zz["cplantilla"].ToString() + ",");
                    x.AppendLine(zz["cusacliente"].ToString() + ",");
                    x.AppendLine(zz["cusaproveedor"].ToString() + ",");
                    x.AppendLine(zz["cafectado"].ToString() + ",");
                    //x.AppendLine(zz["cimpreso"].ToString() + ",");
                    x.AppendLine("0,");

                    x.AppendLine(zz["ccancelado"].ToString() + ",");
                    x.AppendLine(zz["cdevuelto"].ToString() + ",");
                    x.AppendLine(zz["cidprepoliza"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["cidprepolizacancelacion"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["cestadocontable"].ToString() + ",");
                    //x.AppendLine(zz["cneto"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    //x.AppendLine(zz["cimpuesto1"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cimpuesto2"].ToString() + ",");
                    x.AppendLine(zz["cimpuesto3"].ToString() + ",");
                    x.AppendLine(zz["cretencion1"].ToString() + ",");
                    x.AppendLine(zz["cretencion2"].ToString() + ",");
                    x.AppendLine(zz["cdescuentomov"].ToString() + ",");
                    x.AppendLine(zz["cdescuentodoc1"].ToString() + ",");
                    x.AppendLine(zz["cdescuentodoc2"].ToString() + ",");
                    x.AppendLine(zz["cgasto1"].ToString() + ",");
                    x.AppendLine(zz["cgasto2"].ToString() + ",");
                    x.AppendLine(zz["cgasto3"].ToString() + ",");
                    //x.AppendLine(zz["ctotal"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    //x.AppendLine(zz["ctotalun01"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cdescuentoprontopago"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeimpuesto1"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeimpuesto2"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeimpuesto3"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeretencion1"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeretencion2"].ToString() + ",");
                    x.AppendLine(zz["CPORCENTAJEINTERES"].ToString() + ",");
                    x.AppendLine("'" + zz["ctextoextra1"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoextra2"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoextra3"].ToString() + "',");
                    //x.AppendLine("'" + zz["cfechaex01"].ToString() + "',");
                    //x.AppendLine("ctod('" + zz["cfechaex01"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("'" + sFechav + "',");


                    //x.AppendLine(zz["cporcent06"].ToString() + ",");
                    x.AppendLine(zz["cimporteextra1"].ToString() + ",");

                    x.AppendLine(zz["cimporteextra2"].ToString() + ",");
                    x.AppendLine(zz["cimporteextra3"].ToString() + ",");
                    x.AppendLine(zz["cimporteextra4"].ToString() + ",");
                    x.AppendLine("'" + zz["cdestinatario"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumeroguia"].ToString() + "',");
                    x.AppendLine("'" + zz["cmensajeria"].ToString() + "',");
                    x.AppendLine("'" + zz["ccuentamensajeria"].ToString() + "',");
                    x.AppendLine(zz["cnumerocajas"].ToString() + ",");

                    x.AppendLine(zz["cpeso"].ToString() + ",");
                    x.AppendLine(zz["cbanobservaciones"].ToString() + ",");
                    x.AppendLine(zz["cbandatosenvio"].ToString() + ",");
                    x.AppendLine(zz["cbancondicionescredito"].ToString() + ",");
                    x.AppendLine(zz["cbangastos"].ToString() + ",");
                    x.AppendLine(zz["cunidadespendientes"].ToString() + ",");
                    //x.AppendLine("ctod('" + zz["ctimestamp"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("'" + sFechav + "',");

                    x.AppendLine(zz["cimpcheqpaq"].ToString() + ",");
                    x.AppendLine(zz["csistorig"].ToString() + ",");
                    x.AppendLine(zz["cidmonedca"].ToString() + ",");
                    x.AppendLine(zz["ctipocamca"].ToString() + ",");
                    //x.AppendLine(zz["cescfd"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["ctienecfd"].ToString() + ",");
                    x.AppendLine("'" + zz["clugarexpe"].ToString() + "',");
                    x.AppendLine("'" + zz["cmetodopag"].ToString() + "',");
                    x.AppendLine(zz["cnumparcia"].ToString() + ",");
                    x.AppendLine(zz["ccantparci"].ToString() + ",");
                    x.AppendLine("'" + zz["ccondipago"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumctapag"].ToString() + "'");
                    x.AppendLine(",NEWID(),'',0)");
                    //x.AppendLine(",NEWID(),'')");


                    comando = x.ToString();
                    SqlCommand lsql3 = new SqlCommand(comando, _conDes);
                    lsql3.ExecuteNonQuery();


                    x = new StringBuilder();
                    x.AppendLine("insert into admmovimientos values(");
                    //lds.Tables[0].Rows["ciddocum01"].ToString();
                    x.AppendLine(lidmovim.ToString() + ",");
                    x.AppendLine(liddocum.ToString() + ",");
                    x.AppendLine("1,");
                    //x.AppendLine("13,");

                    if (aClientes == 1)
                    {
                        //x.AppendLine(zz["ciddocum02"].ToString() + ",");
                        x.AppendLine("13,");
                        //  x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        //x.AppendLine("40,");
                    }

                    x.AppendLine("0,"); //producto
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //unidades2
                    x.AppendLine("0,"); //cidunidad
                    x.AppendLine("0,"); //cidunida01
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    //x.AppendLine(zz["cneto"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");

                    //x.AppendLine(zz["cimpuesto1"].ToString() + ",");
                    x.AppendLine("0,"); //impuesto1
                    x.AppendLine(zz["cporcentajeimpuesto1"].ToString() + ",");
                    x.AppendLine("0,"); //impuesto2
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //impuesto3
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //retencion1
                    x.AppendLine("0,");
                    x.AppendLine("0,");//retencion2
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 1
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 2
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 3
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 4
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 5
                    x.AppendLine("0,");
                    //x.AppendLine(zz["ctotal"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine("'',");
                    x.AppendLine("'',"); //observaciones
                    x.AppendLine("3,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    //x.AppendLine("3,");

                    //x.AppendLine("ctod('" + zz["cfecha"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("0,0,0,0,0,0,0,");
                    x.AppendLine("0,0,'','','',");
                    //x.AppendLine("ctod('" + fecha + "'),");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("0,0,0,0,");
                    x.AppendLine("'',0,'',0,0,0)");

                    comando = x.ToString();
                    SqlCommand lsql4 = new SqlCommand(comando, _conDes);
                    lsql4.ExecuteNonQuery();

                    liddocum++;
                    lidmovim++;
                }






                /*
                 INSERT INTO [DESTINO]...[mgw10010]
               ([cidmovim01]
               ,[ciddocum01]
               ,[cnumerom01]
               ,[ciddocum02]
               ,[cidprodu01]
               ,[cidalmacen]
               ,[cunidades]
               ,[cunidade01]
               ,[cunidade02]
               ,[cidunidad]
               ,[cidunida01]
               ,[cprecio]
               ,[cprecioc01]
               ,[ccostoca01]
               ,[ccostoes01]
               ,[cneto]
               ,[cimpuesto1]
               ,[cporcent01]
               ,[cimpuesto2]
               ,[cporcent02]
               ,[cimpuesto3]
               ,[cporcent03]
               ,[cretenci01]
               ,[cporcent04]
               ,[cretenci02]
               ,[cporcent05]
               ,[cdescuen01]
               ,[cporcent06]
               ,[cdescuen02]
               ,[cporcent07]
               ,[cdescuen03]
               ,[cporcent08]
               ,[cdescuen04]
               ,[cporcent09]
               ,[cdescuen05]
               ,[cporcent10]
               ,[ctotal]
               ,[cporcent11]
               ,[creferen01]
               ,[cobserva01]
               ,[cafectae01]
               ,[cafectad01]
               ,[cafectad02]
               ,[cfecha]
               ,[cmovtooc01]
               ,[cidmovto01]
               ,[cidmovto02]
               ,[cunidade03]
               ,[cunidade04]
               ,[cunidade05]
               ,[cunidade06]
               ,[ctipotra01]
               ,[cidvalor01]
               ,[ctextoex01]
               ,[ctextoex02]
               ,[ctextoex03]
               ,[cfechaex01]
               ,[cimporte01]
               ,[cimporte02]
               ,[cimporte03]
               ,[cimporte04]
               ,[ctimestamp]
               ,[cgtomovto]
               ,[cscmovto]
               ,[ccomventa]
               ,[cidmovdest]
               ,[cnumconsol])

                 */


            }
            //miconexion.mCerrarConexionDestino();
            _conDes.Close();

            return 0;
        }

        private int mRegresaId(string sql, SqlConnection _conDes)
        {
            SqlCommand com2 = new SqlCommand(sql);

            SqlDataReader ldr2;
            com2.Connection = _conDes;
            ldr2 = com2.ExecuteReader();
            int lid = 0;
            if (ldr2.HasRows == true)
            {
                ldr2.Read();
                lid = int.Parse(ldr2[0].ToString());
            }
            ldr2.Close();
            return lid;

        }

        public int mProcesarInventarios(string comando, string comandomovtos, int aClientes, int aporcodigo, string empresaorigen, string sempresadestino, string comandofolios = "")
        {
            foliodig = 1;
            //lidmovim = 1;
            //liddocum = 1;
            SqlConnection _conOrigen = new SqlConnection();
            SqlConnection _conDestino = new SqlConnection();
            string sempresa = empresaorigen.Substring(empresaorigen.LastIndexOf("\\") + 1);
            string CadenaconexionOrigen = "data source =" + cserver + ";initial catalog = " + sempresa + ";user id = " + cusr + "; password = " + cpwd + ";";
            SqlConnection _con = new SqlConnection();

            _conOrigen.ConnectionString = CadenaconexionOrigen;
            _conOrigen.Open();


            SqlDataAdapter lda = new SqlDataAdapter(comando, _conOrigen);
            System.Data.DataSet lds = new System.Data.DataSet();
            lda.Fill(lds);

            System.Data.DataSet lds1 = null;

            if (comandomovtos != "")
            {
                SqlDataAdapter lda1 = new SqlDataAdapter(comandomovtos, _conOrigen);
                lds1 = new System.Data.DataSet();
                lda1.Fill(lds1);
            }
            SqlDataAdapter lda2 = new SqlDataAdapter();
            System.Data.DataSet lds2 = new System.Data.DataSet();

            SqlCommand lcom = null;
            if (comandofolios != "")
            {
                //SqlDataAdapter lda1 = new SqlDataAdapter(comandomovtos, _conOrigen);
                lcom = new SqlCommand(comandofolios, _conOrigen);
                lda2.SelectCommand = lcom;

                lda2.Fill(lds2);
            }
            /*miconexion.mCerrarConexionOrigen();*/



            sempresa = sempresadestino.Substring(sempresadestino.LastIndexOf("\\") + 1);

            string Cadenaconexion1 = "data source =" + cserver + ";initial catalog = " + sempresa + ";user id = " + cusr + "; password = " + cpwd + ";";
            SqlConnection _conDes = new SqlConnection();

            _conDes.ConnectionString = Cadenaconexion1;

            _conDes.Open();


            SqlCommand comfolios = new SqlCommand("select ISNULL(max(CIDFOLDIG),0) from admFoliosDigitales");
            SqlDataReader ldrfolios;

            comfolios.Connection = _conDes;
            ldrfolios = comfolios.ExecuteReader();
            ldrfolios.Read();
            foliodig = int.Parse(ldrfolios[0].ToString()) + 1;

            ldrfolios.Close();


            SqlCommand com = new SqlCommand("select ISNULL(max(ciddocumento),0) from admDocumentos");
            SqlDataReader ldr;

            com.Connection = _conDes;
            ldr = com.ExecuteReader();
            ldr.Read();
            int liddocum = int.Parse(ldr[0].ToString()) + 1;
            ldr.Close();


            SqlCommand com1 = new SqlCommand("select ISNULL(max(cidmovimiento),0) from admMovimientos");
            SqlDataReader ldr1;
            int lfolio = 1;

            com1.Connection = _conDes;
            ldr.Close();
            ldr1 = com1.ExecuteReader();
            ldr1.Read();
            lidmovim = int.Parse(ldr1[0].ToString()) + 1;
            ldr1.Close();
            foreach (DataRow zz in lds.Tables[0].Rows)
            {

                string ltexto = "";
                if (aporcodigo == 1)
                    ltexto = "select cidclienteproveedor from admClientes where ccodigocliente = '" + zz["ccodigocliente"].ToString() + "'";
                else
                    ltexto = "select cidclienteproveedor from admClientes where ctextoextra1 = '" + zz["ccodigocliente"].ToString() + "'";


                int lidclien = mRegresaId(ltexto, _conDes);
                ltexto = "select cidagente from admAgentes where ccodigoagente = '" + zz["ccodigoagente"].ToString() + "'";
                int lidagent = mRegresaId(ltexto, _conDes);
                ltexto = "select cidconceptodocumento from admConceptos where ccodigoconcepto = '" + zz["ccodigoconcepto"].ToString() + "'";
                int lidconc = mRegresaId(ltexto, _conDes);







                if (lidconc > 0)
                {
                    StringBuilder x = new StringBuilder();
                    x.AppendLine("SET QUOTED_IDENTIFIER OFF ; insert into admDocumentos ");
                    x.AppendLine("(");
                    x.AppendLine("CIDDOCUMENTO");
                    x.AppendLine(",CIDDOCUMENTODE");
                    x.AppendLine(",CIDCONCEPTODOCUMENTO");
                    x.AppendLine(",CSERIEDOCUMENTO");
                    x.AppendLine(",CFOLIO");
                    x.AppendLine(",CFECHA");
                    x.AppendLine(",CIDCLIENTEPROVEEDOR");
                    x.AppendLine(",CRAZONSOCIAL");
                    x.AppendLine(",CRFC");
                    x.AppendLine(",CIDAGENTE");
                    x.AppendLine(",CFECHAVENCIMIENTO");
                    x.AppendLine(",CFECHAPRONTOPAGO");
                    x.AppendLine(",CFECHAENTREGARECEPCION");
                    x.AppendLine(",CFECHAULTIMOINTERES");
                    x.AppendLine(",CIDMONEDA");
                    x.AppendLine(",CTIPOCAMBIO");
                    x.AppendLine(",CREFERENCIA");
                    x.AppendLine(",COBSERVACIONES");
                    x.AppendLine(",CNATURALEZA");
                    x.AppendLine(",CIDDOCUMENTOORIGEN");
                    x.AppendLine(",CPLANTILLA");
                    x.AppendLine(",CUSACLIENTE");
                    x.AppendLine(",CUSAPROVEEDOR");
                    x.AppendLine(",CAFECTADO");
                    x.AppendLine(",CIMPRESO");
                    x.AppendLine(",CCANCELADO");
                    x.AppendLine(",CDEVUELTO");
                    x.AppendLine(",CIDPREPOLIZA");
                    x.AppendLine(",CIDPREPOLIZACANCELACION");
                    x.AppendLine(",CESTADOCONTABLE");
                    x.AppendLine(",CNETO");
                    x.AppendLine(",CIMPUESTO1");
                    x.AppendLine(",CIMPUESTO2");
                    x.AppendLine(",CIMPUESTO3");
                    x.AppendLine(",CRETENCION1");
                    x.AppendLine(",CRETENCION2");
                    x.AppendLine(",CDESCUENTOMOV");
                    x.AppendLine(",CDESCUENTODOC1");
                    x.AppendLine(",CDESCUENTODOC2");
                    x.AppendLine(",CGASTO1");
                    x.AppendLine(",CGASTO2");
                    x.AppendLine(",CGASTO3");
                    x.AppendLine(",CTOTAL");
                    x.AppendLine(",CPENDIENTE");
                    x.AppendLine(",CTOTALUNIDADES");
                    x.AppendLine(",CDESCUENTOPRONTOPAGO");
                    x.AppendLine(",CPORCENTAJEIMPUESTO1");
                    x.AppendLine(",CPORCENTAJEIMPUESTO2");
                    x.AppendLine(",CPORCENTAJEIMPUESTO3");
                    x.AppendLine(",CPORCENTAJERETENCION1");
                    x.AppendLine(",CPORCENTAJERETENCION2");
                    x.AppendLine(",CPORCENTAJEINTERES");
                    x.AppendLine(",CTEXTOEXTRA1");
                    x.AppendLine(",CTEXTOEXTRA2");
                    x.AppendLine(",CTEXTOEXTRA3");
                    x.AppendLine(",CFECHAEXTRA");
                    x.AppendLine(",CIMPORTEEXTRA1");
                    x.AppendLine(",CIMPORTEEXTRA2");
                    x.AppendLine(",CIMPORTEEXTRA3");
                    x.AppendLine(",CIMPORTEEXTRA4");
                    x.AppendLine(",CDESTINATARIO");
                    x.AppendLine(",CNUMEROGUIA");
                    x.AppendLine(",CMENSAJERIA");
                    x.AppendLine(",CCUENTAMENSAJERIA");
                    x.AppendLine(",CNUMEROCAJAS");
                    x.AppendLine(",CPESO");
                    x.AppendLine(",CBANOBSERVACIONES");
                    x.AppendLine(",CBANDATOSENVIO");
                    x.AppendLine(",CBANCONDICIONESCREDITO");
                    x.AppendLine(",CBANGASTOS");
                    x.AppendLine(",CUNIDADESPENDIENTES");
                    x.AppendLine(",CTIMESTAMP");
                    x.AppendLine(",CIMPCHEQPAQ");
                    x.AppendLine(",CSISTORIG");
                    x.AppendLine(",CIDMONEDCA");
                    x.AppendLine(",CTIPOCAMCA");
                    x.AppendLine(",CESCFD");
                    x.AppendLine(",CTIENECFD");
                    x.AppendLine(",CLUGAREXPE");
                    x.AppendLine(",CMETODOPAG");
                    x.AppendLine(",CNUMPARCIA");
                    x.AppendLine(",CCANTPARCI");
                    x.AppendLine(",CCONDIPAGO");
                    x.AppendLine(",CNUMCTAPAG");
                    x.AppendLine(",CGUIDDOCUMENTO");
                    x.AppendLine(",CUSUARIO");
                    x.AppendLine(",CIDPROYECTO");
                    x.AppendLine(",CIDCUENTA");
                    x.AppendLine(",CTRANSACTIONID)");

                    x.AppendLine("values (");
                    x.AppendLine(liddocum + ",");
                    /*
                    if (aClientes == 1)
                    {
                        x.AppendLine("13,");
                        x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        x.AppendLine("40,");
                    }
                     * */

                    x.AppendLine("'" + zz["ciddocumentode"].ToString() + "',");
                    x.AppendLine("'" + lidconc.ToString() + "',");

                    x.AppendLine("'" + zz["cseriedocumento"].ToString() + "',");


                    x.AppendLine(zz["cfoliodocumento"].ToString() + ",");

                    // x.AppendLine(lfolio.ToString() + ",");
                    lfolio++;


                    string fecha = zz["cfecha"].ToString().Substring(0, 10);
                    int espacio = fecha.IndexOf(" ");
                    if (espacio > -1)
                        fecha = fecha.Substring(0, espacio);
                    DateTime dfecha = DateTime.Parse(fecha);
                    fecha = dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                    string sFecha = dfecha.Year.ToString().PadLeft(4, '0') + dfecha.Month.ToString().PadLeft(2, '0') + dfecha.Day.ToString().PadLeft(2, '0');

                    //sFecha = "2018" + dfecha.Month.ToString().PadLeft(2, '0') + "01";

                    x.AppendLine("'" + sFecha + "',");

                    x.AppendLine(lidclien.ToString() + ",");

                    int xxxxx = 0;
                    if (zz["crazonsocial"].ToString() != "")
                        xxxxx = 20;
                    x.AppendLine("'" + zz["crazonsocial"].ToString() + "',");
                    x.AppendLine("'" + zz["crfc"].ToString() + "',");
                    x.AppendLine(lidagent.ToString() + ",");

                    string fechav = zz["cfechavencimiento"].ToString().Substring(0, 10);
                    int espacio1 = fechav.IndexOf(" ");
                    if (espacio1 > -1)
                        fechav = fechav.Substring(0, espacio1);

                    DateTime dfechav = DateTime.Parse(fechav);
                    fechav = dfechav.Month.ToString().PadLeft(2, '0') + "/" + dfechav.Day.ToString().PadLeft(2, '0') + "/" + dfechav.Year.ToString().PadLeft(4, '0');
                    string sFechav = dfechav.Year.ToString().PadLeft(4, '0') + dfechav.Month.ToString().PadLeft(2, '0') + dfechav.Day.ToString().PadLeft(2, '0');
                    x.AppendLine("'" + sFechav + "',");

                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("'" + sFechav + "',");

                    x.AppendLine(zz["cidmoneda"].ToString() + ",");
                    x.AppendLine(zz["ctipocambio"].ToString() + ",");
                    x.AppendLine("'" + zz["creferencia"].ToString() + "',");
                    x.AppendLine("'" + zz["cfoliodocumento"].ToString() + "',");

                    x.AppendLine(zz["cnaturaleza"].ToString() + ",");
                    x.AppendLine(zz["ciddocumentoorigen"].ToString() + ",");
                    x.AppendLine(zz["cplantilla"].ToString() + ",");
                    x.AppendLine(zz["cusacliente"].ToString() + ",");
                    x.AppendLine(zz["cusaproveedor"].ToString() + ",");
                    //x.AppendLine(zz["cafectado"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine("0,"); // impreso

                    x.AppendLine(zz["ccancelado"].ToString() + ",");
                    x.AppendLine(zz["cdevuelto"].ToString() + ",");
                    x.AppendLine(zz["cidprepoliza"].ToString() + ",");
                    x.AppendLine(zz["cidprepolizacancelacion"].ToString() + ",");
                    x.AppendLine(zz["cestadocontable"].ToString() + ",");
                    //x.AppendLine(zz["cpendiente"].ToString() + ",");   // neto
                    //x.AppendLine("0,"); // impuesto1

                    x.AppendLine(zz["cneto"].ToString() + ",");   // neto
                    x.AppendLine(zz["cimpuesto1"].ToString() + ","); // impuesto1


                    x.AppendLine(zz["cimpuesto2"].ToString() + ",");
                    x.AppendLine(zz["cimpuesto3"].ToString() + ",");
                    x.AppendLine(zz["cretencion1"].ToString() + ",");
                    x.AppendLine(zz["cretencion2"].ToString() + ",");
                    x.AppendLine(zz["cdescuentomov"].ToString() + ",");
                    x.AppendLine(zz["cdescuentodoc1"].ToString() + ",");
                    x.AppendLine(zz["cdescuentodoc2"].ToString() + ",");
                    x.AppendLine(zz["cgasto1"].ToString() + ",");
                    x.AppendLine(zz["cgasto2"].ToString() + ",");
                    x.AppendLine(zz["cgasto3"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cdescuentoprontopago"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeimpuesto1"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeimpuesto2"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeimpuesto3"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeretencion1"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeretencion2"].ToString() + ",");
                    x.AppendLine(zz["CPORCENTAJEINTERES"].ToString() + ",");
                    x.AppendLine("'" + zz["ctextoextra1"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoextra2"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoextra3"].ToString() + "',");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine(zz["cimporteextra1"].ToString() + ",");
                    x.AppendLine(zz["cimporteextra2"].ToString() + ",");
                    x.AppendLine(zz["cimporteextra3"].ToString() + ",");
                    x.AppendLine(zz["cimporteextra4"].ToString() + ",");
                    //x.AppendLine("'" + zz["cdestinatario"].ToString() + "',");
                    x.AppendLine("'',");
                    x.AppendLine("'" + zz["cnumeroguia"].ToString() + "',");
                    x.AppendLine("'" + zz["cmensajeria"].ToString() + "',");
                    x.AppendLine("'" + zz["ccuentamensajeria"].ToString() + "',");
                    x.AppendLine(zz["cnumerocajas"].ToString() + ",");

                    x.AppendLine(zz["cpeso"].ToString() + ",");
                    x.AppendLine(zz["cbanobservaciones"].ToString() + ",");
                    x.AppendLine(zz["cbandatosenvio"].ToString() + ",");
                    x.AppendLine(zz["cbancondicionescredito"].ToString() + ",");
                    x.AppendLine(zz["cbangastos"].ToString() + ",");
                    x.AppendLine(zz["cunidadespendientes"].ToString() + ",");
                    x.AppendLine("'" + sFechav + "',");

                    x.AppendLine(zz["cimpcheqpaq"].ToString() + ",");
                    x.AppendLine(zz["csistorig"].ToString() + ",");
                    x.AppendLine(zz["cidmonedca"].ToString() + ",");
                    x.AppendLine(zz["ctipocamca"].ToString() + ",");
                    x.AppendLine(zz["cescfd"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["ctienecfd"].ToString() + ",");
                    x.AppendLine("'" + zz["clugarexpe"].ToString() + "',");
                    x.AppendLine("'" + zz["cmetodopag"].ToString() + "',");
                    x.AppendLine(zz["cnumparcia"].ToString() + ",");
                    x.AppendLine(zz["ccantparci"].ToString() + ",");
                    x.AppendLine("'" + zz["ccondipago"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumctapag"].ToString() + "'");
                    x.AppendLine(",NEWID(),'',0,0,'')");

                    comando = x.ToString();
                    SqlCommand lsql3 = new SqlCommand(comando, _conDes);
                    lsql3.ExecuteNonQuery();


                    if (comandomovtos != "")
                    {
                        mGeneraMovimientos(zz["ciddocumento"].ToString(), liddocum, lds1, _conDes);
                    }

                    if (zz["ciddocumentode"].ToString() == "5" || zz["ciddocumentode"].ToString() == "4" || zz["ciddocumentode"].ToString() == "7")
                    {
                        mFolioDigital(zz["ciddocumento"].ToString(), liddocum, lds2, _conDes, lidconc, zz["cpendiente"].ToString());

                    }

                    liddocum++;
                    //lidmovim++;
                }




            }
            //miconexion.mCerrarConexionDestino();
            _conDes.Close();

            return 0;
        }
        int foliodig = 1;

        private void mFolioDigital(string liddocumento, int liddocum, DataSet lds2, SqlConnection _conDes, int lidconc, string aTotal)
        {

            var results = from DataRow myRow in lds2.Tables[0].Rows
                          where myRow["ciddocto"].ToString() == liddocumento
                          select myRow;
            foreach (var a in results)
            {
                //string x = a["cfolio"].ToString();
                StringBuilder x = new StringBuilder();
                x.AppendLine("INSERT INTO admFoliosDigitales (");
                x.AppendLine("CIDFOLDIG");
                x.AppendLine(",CIDDOCTODE");
                x.AppendLine(",CIDCPTODOC");
                x.AppendLine(",CIDDOCTO");
                x.AppendLine(",CIDDOCALDI");
                x.AppendLine(",CIDFIRMARL");
                x.AppendLine(",CNOORDEN");
                x.AppendLine(",CSERIE");
                x.AppendLine(",CFOLIO");
                x.AppendLine(",CNOAPROB");
                x.AppendLine(",CFECAPROB");
                x.AppendLine(",CESTADO");
                x.AppendLine(",CENTREGADO");
                x.AppendLine(",CFECHAEMI");
                x.AppendLine(",CHORAEMI");
                x.AppendLine(",CEMAIL");
                x.AppendLine(",CARCHDIDIS");
                x.AppendLine(",CIDCPTOORI");
                x.AppendLine(",CFECHACANC");
                x.AppendLine(",CHORACANC");
                x.AppendLine(",CESTRAD");
                x.AppendLine(",CCADPEDI");
                x.AppendLine(",CARCHCBB");
                x.AppendLine(",CINIVIG");
                x.AppendLine(",CFINVIG");
                x.AppendLine(",CTIPO");
                x.AppendLine(",CSERIEREC");
                x.AppendLine(",CFOLIOREC");
                x.AppendLine(",CRFC");
                x.AppendLine(",CRAZON");
                x.AppendLine(",CSISORIGEN");
                x.AppendLine(",CEJERPOL");
                x.AppendLine(",CPERPOL");
                x.AppendLine(",CTIPOPOL");
                x.AppendLine(",CNUMPOL");
                x.AppendLine(",CTIPOLDESC");
                x.AppendLine(",CTOTAL");
                x.AppendLine(",CALIASBDCT");
                x.AppendLine(",CCFDPRUEBA");
                x.AppendLine(",CDESESTADO");
                x.AppendLine(",CPAGADOBAN");
                x.AppendLine(",CDESPAGBAN");
                x.AppendLine(",CREFEREN01");
                x.AppendLine(",COBSERVA01");
                x.AppendLine(",CCODCONCBA");
                x.AppendLine(",CDESCONCBA");
                x.AppendLine(",CNUMCTABAN");
                x.AppendLine(",CFOLIOBAN");
                x.AppendLine(",CIDDOCDEBA");
                x.AppendLine(",CUSUAUTBAN");
                x.AppendLine(",CUUID");
                x.AppendLine(",CUSUBAN01");
                x.AppendLine(",CAUTUSBA01");
                x.AppendLine(",CUSUBAN02");
                x.AppendLine(",CAUTUSBA02");
                x.AppendLine(",CUSUBAN03");
                x.AppendLine(",CAUTUSBA03");
                x.AppendLine(",CDESCAUT01");
                x.AppendLine(",CDESCAUT02");
                x.AppendLine(",CDESCAUT03");
                x.AppendLine(",CERRORVAL");
                x.AppendLine(",CACUSECAN");
                x.AppendLine(",CIDDOCTODSL )");

                x.AppendLine("values ( ");
                x.AppendLine(foliodig.ToString() + ",");
                x.AppendLine(a["CIDDOCTODE"].ToString() + ","); //,CIDDOCTODE
                x.AppendLine(lidconc.ToString() + ","); //,CIDCPTODOC
                x.AppendLine(liddocum.ToString()); //,,d.CIDDOCUMENTO as ciddocto
                x.AppendLine("," + a["CIDDOCALDI"].ToString());
                x.AppendLine("," + a["CIDFIRMARL"].ToString());
                x.AppendLine("," + a["CNOORDEN"].ToString());
                x.AppendLine(",'" + a["CSERIE"].ToString() + "'");
                x.AppendLine("," + a["CFOLIO"].ToString());
                x.AppendLine("," + a["CNOAPROB"].ToString());
                //x.AppendLine(",CONVERT(VARCHAR(10),'" +  a["CFECAPROB"].ToString() + "', 112) "); 

                string lcad = ",SUBSTRING('" + a["CFECAPROB"].ToString() + "',7,4) +";
                lcad += "SUBSTRING('" + a["CFECAPROB"].ToString() + "',4,2) +";
                lcad += "LEFT('" + a["CFECAPROB"].ToString() + "',2)";


                x.AppendLine(lcad);




                x.AppendLine("," + a["CESTADO"].ToString());
                x.AppendLine("," + a["CENTREGADO"].ToString());

                lcad = ",SUBSTRING('" + a["CFECHAEMI"].ToString() + "',7,4) +";
                lcad += "SUBSTRING('" + a["CFECHAEMI"].ToString() + "',4,2) +";
                lcad += "LEFT('" + a["CFECHAEMI"].ToString() + "',2)";


                x.AppendLine(lcad);
                x.AppendLine(",'" + a["CHORAEMI"].ToString() + "'");
                x.AppendLine(",'" + a["CEMAIL"].ToString() + "'");
                x.AppendLine(",'" + a["CARCHDIDIS"].ToString() + "'");
                x.AppendLine("," + a["CIDCPTOORI"].ToString());


                lcad = ",SUBSTRING('" + a["CFECHACANC"].ToString() + "',7,4) +";
                lcad += "SUBSTRING('" + a["CFECHACANC"].ToString() + "',4,2) +";
                lcad += "LEFT('" + a["CFECHACANC"].ToString() + "',2)";
                x.AppendLine(lcad);
                x.AppendLine(",'" + a["CHORACANC"].ToString() + "'");
                x.AppendLine("," + a["CESTRAD"].ToString());
                x.AppendLine(",'" + a["CCADPEDI"].ToString() + "'");
                x.AppendLine(",'" + a["CARCHCBB"].ToString() + "'");

                lcad = ",SUBSTRING('" + a["CINIVIG"].ToString() + "',7,4) +";
                lcad += "SUBSTRING('" + a["CINIVIG"].ToString() + "',4,2) +";
                lcad += "LEFT('" + a["CINIVIG"].ToString() + "',2)";
                x.AppendLine(lcad);

                lcad = ",SUBSTRING('" + a["CFINVIG"].ToString() + "',7,4) +";
                lcad += "SUBSTRING('" + a["CFINVIG"].ToString() + "',4,2) +";
                lcad += "LEFT('" + a["CFINVIG"].ToString() + "',2)";

                x.AppendLine(lcad);
                x.AppendLine(",'" + a["CTIPO"].ToString() + "'");
                x.AppendLine(",'" + a["CSERIEREC"].ToString() + "'");
                x.AppendLine("," + a["CFOLIOREC"].ToString());
                x.AppendLine(",'" + a["CRFC"].ToString() + "'");
                x.AppendLine(",'" + a["CRAZON"].ToString() + "'");
                x.AppendLine("," + a["CSISORIGEN"].ToString());
                x.AppendLine("," + a["CEJERPOL"].ToString());
                x.AppendLine("," + a["CPERPOL"].ToString());
                x.AppendLine("," + a["CTIPOPOL"].ToString());
                x.AppendLine("," + a["CNUMPOL"].ToString());
                x.AppendLine(",'" + a["CTIPOLDESC"].ToString() + "'");
                //x.AppendLine("," + a["CTOTAL"].ToString()); // debe Ser el PENDIENTE
                x.AppendLine("," + aTotal); // debe Ser el PENDIENTE
                x.AppendLine(",'" + a["CALIASBDCT"].ToString() + "'");
                x.AppendLine("," + a["CCFDPRUEBA"].ToString());
                x.AppendLine(",'" + a["CDESESTADO"].ToString() + "'");
                x.AppendLine("," + a["CPAGADOBAN"].ToString());
                x.AppendLine(",'" + a["CDESPAGBAN"].ToString() + "'");
                x.AppendLine(",'" + a["CREFEREN01"].ToString() + "'");
                x.AppendLine(",'" + a["COBSERVA01"].ToString() + "'");
                x.AppendLine(",'" + a["CCODCONCBA"].ToString() + "'");
                x.AppendLine(",'" + a["CDESCONCBA"].ToString() + "'");
                x.AppendLine(",'" + a["CNUMCTABAN"].ToString() + "'");
                x.AppendLine(",'" + a["CFOLIOBAN"].ToString() + "'");
                x.AppendLine("," + a["CIDDOCDEBA"].ToString());
                x.AppendLine(",'" + a["CUSUAUTBAN"].ToString() + "'");
                x.AppendLine(",'" + a["CUUID"].ToString() + "'");
                x.AppendLine(",'" + a["CUSUBAN01"].ToString() + "'");
                x.AppendLine("," + a["CAUTUSBA01"].ToString());

                x.AppendLine(",'" + a["CUSUBAN02"].ToString() + "'");

                x.AppendLine("," + a["CAUTUSBA02"].ToString());
                x.AppendLine(",'" + a["CUSUBAN03"].ToString() + "'");

                x.AppendLine("," + a["CAUTUSBA03"].ToString());
                x.AppendLine(",'" + a["CDESCAUT01"].ToString() + "'");
                x.AppendLine(",'" + a["CDESCAUT02"].ToString() + "'");
                x.AppendLine(",'" + a["CDESCAUT03"].ToString() + "'");
                x.AppendLine("," + a["CERRORVAL"].ToString());
                x.AppendLine(",'" + a["CACUSECAN"].ToString() + "'");
                x.AppendLine(",'" + a["CIDDOCTODSL"].ToString() + "')");

                string comando = x.ToString();
                SqlCommand lsql4 = new SqlCommand(comando, _conDes);
                lsql4.ExecuteNonQuery();
                foliodig++;



                /*
                CIDFIRMARL
                CNOORDEN
                CSERIE
                dd.CFOLIO
                CNOAPROB
                CFECAPROB
                CESTADO
                CENTREGADO
                CFECHAEMI
                CHORAEMI
                CEMAIL
                CARCHDIDIS
                CIDCPTOORI
                CFECHACANC
                CHORACANC
                CESTRAD
                CCADPEDI
                CARCHCBB
                CINIVIG
                CFINVIG
                CTIPO
                CSERIEREC
                CFOLIOREC
                do.CRFC
                CRAZON
                CSISORIGEN
                CEJERPOL
                CPERPOL
                CTIPOPOL
                CNUMPOL
                CTIPOLDESC
                do.CTOTAL
                CALIASBDCT
                CCFDPRUEBA
                CDESESTADO
                CPAGADOBAN
                CDESPAGBAN
                CREFEREN01
                COBSERVA01
                CCODCONCBA
                CDESCONCBA
                CNUMCTABAN
                CFOLIOBAN
                CIDDOCDEBA
                CUSUAUTBAN
                CUUID
                CUSUBAN01
                CAUTUSBA01
                CUSUBAN02
                CAUTUSBA02
                CUSUBAN03
                CAUTUSBA03
                CDESCAUT01
                CDESCAUT02
                CDESCAUT03
                CERRORVAL
                CACUSECAN
                CIDDOCTODSL*/

            }
        }



        private void mFolioDigital1(RegDocto ldocto, int lidconc, string ltotal)
        {


            SqlCommand comfolios = new SqlCommand("select ISNULL(max(CIDFOLDIG),0) from admFoliosDigitales");
            SqlDataReader ldrfolios;

            comfolios.Connection = miconexion._conexion1;
            ldrfolios = comfolios.ExecuteReader();
            ldrfolios.Read();
            foliodig = int.Parse(ldrfolios[0].ToString()) + 1;
            ldrfolios.Close();


            //string x = a["cfolio"].ToString();
            StringBuilder x = new StringBuilder();
            x.AppendLine("INSERT INTO admFoliosDigitales (");
            x.AppendLine("CIDFOLDIG");
            x.AppendLine(",CIDDOCTODE");
            x.AppendLine(",CIDCPTODOC");
            x.AppendLine(",CIDDOCTO");
            x.AppendLine(",CIDDOCALDI");
            x.AppendLine(",CIDFIRMARL");
            x.AppendLine(",CNOORDEN");
            x.AppendLine(",CSERIE");
            x.AppendLine(",CFOLIO");
            x.AppendLine(",CNOAPROB");
            x.AppendLine(",CFECAPROB");
            x.AppendLine(",CESTADO");
            x.AppendLine(",CENTREGADO");
            x.AppendLine(",CFECHAEMI");
            x.AppendLine(",CHORAEMI");
            x.AppendLine(",CEMAIL");
            x.AppendLine(",CARCHDIDIS");
            x.AppendLine(",CIDCPTOORI");
            x.AppendLine(",CFECHACANC");
            x.AppendLine(",CHORACANC");
            x.AppendLine(",CESTRAD");
            x.AppendLine(",CCADPEDI");
            x.AppendLine(",CARCHCBB");
            x.AppendLine(",CINIVIG");
            x.AppendLine(",CFINVIG");
            x.AppendLine(",CTIPO");
            x.AppendLine(",CSERIEREC");
            x.AppendLine(",CFOLIOREC");
            x.AppendLine(",CRFC");
            x.AppendLine(",CRAZON");
            x.AppendLine(",CSISORIGEN");
            x.AppendLine(",CEJERPOL");
            x.AppendLine(",CPERPOL");
            x.AppendLine(",CTIPOPOL");
            x.AppendLine(",CNUMPOL");
            x.AppendLine(",CTIPOLDESC");
            x.AppendLine(",CTOTAL");
            x.AppendLine(",CALIASBDCT");
            x.AppendLine(",CCFDPRUEBA");
            x.AppendLine(",CDESESTADO");
            x.AppendLine(",CPAGADOBAN");
            x.AppendLine(",CDESPAGBAN");
            x.AppendLine(",CREFEREN01");
            x.AppendLine(",COBSERVA01");
            x.AppendLine(",CCODCONCBA");
            x.AppendLine(",CDESCONCBA");
            x.AppendLine(",CNUMCTABAN");
            x.AppendLine(",CFOLIOBAN");
            x.AppendLine(",CIDDOCDEBA");
            x.AppendLine(",CUSUAUTBAN");
            x.AppendLine(",CUUID");
            x.AppendLine(",CUSUBAN01");
            x.AppendLine(",CAUTUSBA01");
            x.AppendLine(",CUSUBAN02");
            x.AppendLine(",CAUTUSBA02");
            x.AppendLine(",CUSUBAN03");
            x.AppendLine(",CAUTUSBA03");
            x.AppendLine(",CDESCAUT01");
            x.AppendLine(",CDESCAUT02");
            x.AppendLine(",CDESCAUT03");
            x.AppendLine(",CERRORVAL");
            x.AppendLine(",CACUSECAN");
            x.AppendLine(",CIDDOCTODSL )");

            x.AppendLine("values ( ");
            x.AppendLine(foliodig.ToString() + ",");
            x.AppendLine("4,"); //,CIDDOCTODE
            x.AppendLine(lidconc.ToString() + ","); //,CIDCPTODOC
            x.AppendLine(ldocto.cIdDocto.ToString()); //,,d.CIDDOCUMENTO as ciddocto
            x.AppendLine(",0"); //CIDDOCALDI
            x.AppendLine(",0"); //CIDFIRMARL
            x.AppendLine(",0");
            x.AppendLine(",'" + ldocto.cSerie + "'");
            x.AppendLine("," + ldocto.cFolio.ToString());
            x.AppendLine(",0");
            //x.AppendLine(",CONVERT(VARCHAR(10),'" +  a["CFECAPROB"].ToString() + "', 112) "); 

            string lcad = ",'1899-12-30 00:00:00.000'";


            x.AppendLine(lcad);




            x.AppendLine(",2"); //"," + a["CESTADO"].ToString());
            x.AppendLine(",1"); //+ a["CENTREGADO"].ToString());

            //lcad = ",SUBSTRING('" + a["CFECHAEMI"].ToString() + "',7,4) +";
            //lcad += "SUBSTRING('" + a["CFECHAEMI"].ToString() + "',4,2) +";
            //lcad += "LEFT('" + a["CFECHAEMI"].ToString() + "',2)";

            lcad = ",'" + ldocto.cFecha.ToString("yyyyMMdd") + "'";

            x.AppendLine(lcad);
            //x.AppendLine(",'" + a["CHORAEMI"].ToString() + "'");
            lcad = ",'" + ldocto.cFecha.ToString("hh:mm:ss") + "'";

            x.AppendLine(lcad);
            x.AppendLine(",''");//'" + a["CEMAIL"].ToString() + "'");
            x.AppendLine(",''"); //+ a["CARCHDIDIS"].ToString() + "'");
            x.AppendLine(",0");// + a["CIDCPTOORI"].ToString());


            x.AppendLine(",'1899-12-30 00:00:00.000'");
            x.AppendLine(",''"); //+ a["CHORACANC"].ToString() + "'");
            x.AppendLine(",3");// + a["CESTRAD"].ToString());
            x.AppendLine(",''");// + a["CCADPEDI"].ToString() + "'");
            x.AppendLine(",''");// + a["CARCHCBB"].ToString() + "'");

            //lcad = ",SUBSTRING('" + a["CINIVIG"].ToString() + "',7,4) +";
            //lcad += "SUBSTRING('" + a["CINIVIG"].ToString() + "',4,2) +";
            ///lcad += "LEFT('" + a["CINIVIG"].ToString() + "',2)";
            x.AppendLine(",'1899-12-30 00:00:00.000'");

            //lcad = ",SUBSTRING('" + a["CFINVIG"].ToString() + "',7,4) +";
            //lcad += "SUBSTRING('" + a["CFINVIG"].ToString() + "',4,2) +";
            //lcad += "LEFT('" + a["CFINVIG"].ToString() + "',2)";

            x.AppendLine(",'1899-12-30 00:00:00.000'");
            x.AppendLine(",''"); //+ a["CTIPO"].ToString()+"'");
            x.AppendLine(",''");// + a["CSERIEREC"].ToString()+"'");
            x.AppendLine(",0");// + a["CFOLIOREC"].ToString());
            x.AppendLine(",'" + ldocto.cRFC + "'");
            x.AppendLine(",'" + ldocto.cRazonSocial + "'");
            x.AppendLine(",0");// + a["CSISORIGEN"].ToString());
            x.AppendLine(",0"); //+ a["CEJERPOL"].ToString());
            x.AppendLine(",0"); //+ a["CPERPOL"].ToString());
            x.AppendLine(",0"); //+ a["CTIPOPOL"].ToString());
            x.AppendLine(",0"); //+ a["CNUMPOL"].ToString());
            x.AppendLine(",''");// + a["CTIPOLDESC"].ToString()+ "'");
                                //x.AppendLine("," + a["CTOTAL"].ToString()); // debe Ser el PENDIENTE
            x.AppendLine(",0"); // debe Ser el PENDIENTE
            x.AppendLine(",''");// + a["CALIASBDCT"].ToString()+"'");
            x.AppendLine(",0");// + a["CCFDPRUEBA"].ToString());
            x.AppendLine(",''");// + a["CDESESTADO"].ToString()+"'");
            x.AppendLine(",0");// + a["CPAGADOBAN"].ToString());
            x.AppendLine(",''"); //+ a["CDESPAGBAN"].ToString()+"'");
            x.AppendLine(",''");// + a["CREFEREN01"].ToString()+"'");
            x.AppendLine(",''");// + a["COBSERVA01"].ToString()+"'");
            x.AppendLine(",'P01'");// + a["CCODCONCBA"].ToString()+"'");
            x.AppendLine(",''");// + a["CDESCONCBA"].ToString()+"'");
            x.AppendLine(",''");// + a["CNUMCTABAN"].ToString()+"'");
            x.AppendLine(",''"); //+ a["CFOLIOBAN"].ToString()+"'");
            x.AppendLine(",0");// + a["CIDDOCDEBA"].ToString());
            x.AppendLine(",''");// + a["CUSUAUTBAN"].ToString()+"'");
            x.AppendLine(",'" + ldocto.cTextoExtra1 + "'");
            x.AppendLine(",''");// + a["CUSUBAN01"].ToString()+"'");
            x.AppendLine(",0");// + a["CAUTUSBA01"].ToString());

            x.AppendLine(",''");// + a["CUSUBAN02"].ToString()+"'");

            x.AppendLine(",0");// + a["CAUTUSBA02"].ToString());
            x.AppendLine(",''");// + a["CUSUBAN03"].ToString()+"'");

            x.AppendLine(",0");// + a["CAUTUSBA03"].ToString());
            x.AppendLine(",''");// + a["CDESCAUT01"].ToString()+ "'");
            x.AppendLine(",''");// + a["CDESCAUT02"].ToString()+ "'");
            x.AppendLine(",''");// + a["CDESCAUT03"].ToString()+ "'");
            x.AppendLine(",0");// + a["CERRORVAL"].ToString());
            x.AppendLine(",''");// + a["CACUSECAN"].ToString()+"'");
            x.AppendLine(",'')");// + a["CIDDOCTODSL"].ToString() + "')");

            string comando = x.ToString();

            comando = "update admFoliosDigitales set CUUID = '" + ldocto.cTextoExtra1 + "', cestado = 2, centregado = 1 where ciddocto =" + ldocto.cIdDocto;
            SqlCommand lsql4 = new SqlCommand(comando, miconexion._conexion1);
            lsql4.ExecuteNonQuery();
            foliodig++;



            /*
            CIDFIRMARL
            CNOORDEN
            CSERIE
            dd.CFOLIO
            CNOAPROB
            CFECAPROB
            CESTADO
            CENTREGADO
            CFECHAEMI
            CHORAEMI
            CEMAIL
            CARCHDIDIS
            CIDCPTOORI
            CFECHACANC
            CHORACANC
            CESTRAD
            CCADPEDI
            CARCHCBB
            CINIVIG
            CFINVIG
            CTIPO
            CSERIEREC
            CFOLIOREC
            do.CRFC
            CRAZON
            CSISORIGEN
            CEJERPOL
            CPERPOL
            CTIPOPOL
            CNUMPOL
            CTIPOLDESC
            do.CTOTAL
            CALIASBDCT
            CCFDPRUEBA
            CDESESTADO
            CPAGADOBAN
            CDESPAGBAN
            CREFEREN01
            COBSERVA01
            CCODCONCBA
            CDESCONCBA
            CNUMCTABAN
            CFOLIOBAN
            CIDDOCDEBA
            CUSUAUTBAN
            CUUID
            CUSUBAN01
            CAUTUSBA01
            CUSUBAN02
            CAUTUSBA02
            CUSUBAN03
            CAUTUSBA03
            CDESCAUT01
            CDESCAUT02
            CDESCAUT03
            CERRORVAL
            CACUSECAN
            CIDDOCTODSL*/


        }


        int lidmovim = 1;

        private string mRegresaIdProducto(SqlConnection _conDes, string lcodigo)
        {
            //string lidproducto = ldr2[0].ToString();
            string ltexto = "";
            //
            ltexto = "select isnull(cidproducto,0) from admProductos where ccodigoproducto = '" + lcodigo + "'";
            SqlCommand com2 = new SqlCommand(ltexto);
            SqlDataReader ldr2;
            com2.Connection = _conDes;
            ldr2 = com2.ExecuteReader();
            if (ldr2.HasRows)
            {
                ldr2.Read();
                ltexto = ldr2[0].ToString();

            }
            else
            {
                ltexto = "0";
            }
            ldr2.Close();
            return ltexto;
        }


        private string mRegresaIdAlmacen(SqlConnection _conDes, string lcodigo)
        {
            //string lidproducto = ldr2[0].ToString();
            string ltexto = "";
            //
            ltexto = "select isnull(cidalmacen,0) from admAlmacenes where ccodigoalmacen = '" + lcodigo + "'";
            SqlCommand com2 = new SqlCommand(ltexto);
            SqlDataReader ldr2;
            com2.Connection = _conDes;
            ldr2 = com2.ExecuteReader();
            if (ldr2.HasRows)
            {
                ldr2.Read();
                ltexto = ldr2[0].ToString();
            }
            else
            {
                ltexto = "0";
            }
            ldr2.Close();
            return ltexto;
        }
        private int mGeneraMovimientos(string liddocumento, int liddocum, DataSet lds1, SqlConnection _conDes)
        {

            var results = from DataRow myRow in lds1.Tables[0].Rows
                          where myRow["ciddocumento"].ToString() == liddocumento
                          select myRow;



            int lnumeromov = 100;

            foreach (var a in results)
            {
                StringBuilder x = new StringBuilder();
                x.AppendLine("insert into admmovimientos (");
                x.AppendLine("            CIDMOVIMIENTO");
                x.AppendLine(",CIDDOCUMENTO");
                x.AppendLine(",CNUMEROMOVIMIENTO");
                x.AppendLine(",CIDDOCUMENTODE");
                x.AppendLine(",CIDPRODUCTO");
                x.AppendLine(",CIDALMACEN");
                x.AppendLine(",CUNIDADES");
                x.AppendLine(",CUNIDADESNC");
                x.AppendLine(",CUNIDADESCAPTURADAS");
                x.AppendLine(",CIDUNIDAD");
                x.AppendLine(",CIDUNIDADNC");
                x.AppendLine(",CPRECIO");
                x.AppendLine(",CPRECIOCAPTURADO");
                x.AppendLine(",CCOSTOCAPTURADO");
                x.AppendLine(",CCOSTOESPECIFICO");
                x.AppendLine(",CNETO");
                x.AppendLine(",CIMPUESTO1");
                x.AppendLine(",CPORCENTAJEIMPUESTO1");
                x.AppendLine(",CIMPUESTO2");
                x.AppendLine(",CPORCENTAJEIMPUESTO2");
                x.AppendLine(",CIMPUESTO3");
                x.AppendLine(",CPORCENTAJEIMPUESTO3");
                x.AppendLine(",CRETENCION1");
                x.AppendLine(",CPORCENTAJERETENCION1");
                x.AppendLine(",CRETENCION2");
                x.AppendLine(",CPORCENTAJERETENCION2");
                x.AppendLine(",CDESCUENTO1");
                x.AppendLine(",CPORCENTAJEDESCUENTO1");
                x.AppendLine(",CDESCUENTO2");
                x.AppendLine(",CPORCENTAJEDESCUENTO2");
                x.AppendLine(",CDESCUENTO3");
                x.AppendLine(",CPORCENTAJEDESCUENTO3");
                x.AppendLine(",CDESCUENTO4");
                x.AppendLine(",CPORCENTAJEDESCUENTO4");
                x.AppendLine(",CDESCUENTO5");
                x.AppendLine(",CPORCENTAJEDESCUENTO5");
                x.AppendLine(",CTOTAL");
                x.AppendLine(",CPORCENTAJECOMISION");
                x.AppendLine(",CREFERENCIA");
                x.AppendLine(",COBSERVAMOV");
                x.AppendLine(",CAFECTAEXISTENCIA");
                x.AppendLine(",CAFECTADOSALDOS");
                x.AppendLine(",CAFECTADOINVENTARIO");
                x.AppendLine(",CFECHA");
                x.AppendLine(",CMOVTOOCULTO");
                x.AppendLine(",CIDMOVTOOWNER");
                x.AppendLine(",CIDMOVTOORIGEN");
                x.AppendLine(",CUNIDADESPENDIENTES");
                x.AppendLine(",CUNIDADESNCPENDIENTES");
                x.AppendLine(",CUNIDADESORIGEN");
                x.AppendLine(",CUNIDADESNCORIGEN");
                x.AppendLine(",CTIPOTRASPASO");
                x.AppendLine(",CIDVALORCLASIFICACION");
                x.AppendLine(",CTEXTOEXTRA1");
                x.AppendLine(",CTEXTOEXTRA2");
                x.AppendLine(",CTEXTOEXTRA3");
                x.AppendLine(",CFECHAEXTRA");
                x.AppendLine(",CIMPORTEEXTRA1");
                x.AppendLine(",CIMPORTEEXTRA2");
                x.AppendLine(",CIMPORTEEXTRA3");
                x.AppendLine(",CIMPORTEEXTRA4");
                x.AppendLine(",CTIMESTAMP");
                x.AppendLine(",CGTOMOVTO");
                x.AppendLine(",CSCMOVTO");
                x.AppendLine(",CCOMVENTA");
                x.AppendLine(",CIDMOVTODESTINO");
                x.AppendLine(",CNUMEROCONSOLIDACIONES )");

                x.AppendLine("values (");

                x.AppendLine(lidmovim.ToString() + ",");
                x.AppendLine(liddocum.ToString() + ",");

                x.AppendLine(lnumeromov.ToString() + ",");



                lnumeromov += 100;

                x.AppendLine(a["ciddocumentode"].ToString() + ",");

                x.AppendLine(mRegresaIdProducto(_conDes, a["ccodigoproducto"].ToString()) + ","); // producto

                x.AppendLine(mRegresaIdAlmacen(_conDes, a["ccodigoalmacen"].ToString()) + ","); // almacen

                //x.AppendLine("32,");

                x.AppendLine(a["CUNIDADES"] + ",");
                x.AppendLine(a["CUNIDADESNC"] + ",");
                x.AppendLine(a["CUNIDADESCAPTURADAS"] + ",");
                x.AppendLine(a["CIDUNIDAD"] + ",");
                x.AppendLine(a["CIDUNIDADNC"] + ",");
                x.AppendLine(a["CPRECIO"] + ",");
                x.AppendLine(a["CPRECIOCAPTURADO"] + ",");
                x.AppendLine(a["CCOSTOCAPTURADO"] + ",");
                x.AppendLine(a["CCOSTOESPECIFICO"] + ",");
                x.AppendLine(a["CNETO"] + ",");
                x.AppendLine(a["CIMPUESTO1"] + ",");
                x.AppendLine(a["CPORCENTAJEIMPUESTO1"] + ",");
                x.AppendLine(a["CIMPUESTO2"] + ",");
                x.AppendLine(a["CPORCENTAJEIMPUESTO2"] + ",");
                x.AppendLine(a["CIMPUESTO3"] + ",");
                x.AppendLine(a["CPORCENTAJEIMPUESTO3"] + ",");
                x.AppendLine(a["CRETENCION1"] + ",");
                x.AppendLine(a["CPORCENTAJERETENCION1"] + ",");
                x.AppendLine(a["CRETENCION2"] + ",");
                x.AppendLine(a["CPORCENTAJERETENCION2"] + ",");
                x.AppendLine(a["CDESCUENTO1"] + ",");
                x.AppendLine(a["CPORCENTAJEDESCUENTO1"] + ",");
                x.AppendLine(a["CDESCUENTO2"] + ",");
                x.AppendLine(a["CPORCENTAJEDESCUENTO2"] + ",");
                x.AppendLine(a["CDESCUENTO3"] + ",");
                x.AppendLine(a["CPORCENTAJEDESCUENTO3"] + ",");
                x.AppendLine(a["CDESCUENTO4"] + ",");
                x.AppendLine(a["CPORCENTAJEDESCUENTO4"] + ",");
                x.AppendLine(a["CDESCUENTO5"] + ",");
                x.AppendLine(a["CPORCENTAJEDESCUENTO5"] + ",");
                x.AppendLine(a["CTOTAL"] + ",");
                x.AppendLine(a["CPORCENTAJECOMISION"] + ",");
                x.AppendLine("'" + a["CREFERENCIA"] + "',");
                x.AppendLine("'" + a["COBSERVAMOV"] + "',");
                //x.AppendLine(a["CAFECTAEXISTENCIA"] + ",");
                x.AppendLine("1,");
                //x.AppendLine(a["CAFECTADOSALDOS"] + ",");
                x.AppendLine("0,");
                //x.AppendLine(a["CAFECTADOINVENTARIO"] + ",");
                x.AppendLine("0,");
                x.AppendLine("'" + a["CFECHA"] + "',");
                x.AppendLine(a["CMOVTOOCULTO"] + ",");
                x.AppendLine(a["CIDMOVTOOWNER"] + ",");
                x.AppendLine(a["CIDMOVTOORIGEN"] + ",");
                x.AppendLine(a["CUNIDADESPENDIENTES"] + ",");
                x.AppendLine(a["CUNIDADESNCPENDIENTES"] + ",");
                x.AppendLine(a["CUNIDADESORIGEN"] + ",");
                x.AppendLine(a["CUNIDADESNCORIGEN"] + ",");
                x.AppendLine(a["CTIPOTRASPASO"] + ",");
                x.AppendLine(a["CIDVALORCLASIFICACION"] + ",");
                x.AppendLine("'" + a["CTEXTOEXTRA1"] + "',");
                x.AppendLine("'" + a["CTEXTOEXTRA2"] + "',");
                x.AppendLine("'" + a["CTEXTOEXTRA3"] + "',");
                x.AppendLine("'" + a["CFECHAEXTRA"] + "',");
                x.AppendLine(a["CIMPORTEEXTRA1"] + ",");
                x.AppendLine(a["CIMPORTEEXTRA2"] + ",");
                x.AppendLine(a["CIMPORTEEXTRA3"] + ",");
                x.AppendLine(a["CIMPORTEEXTRA4"] + ",");
                x.AppendLine("'" + a["CTIMESTAMP"] + "',");
                x.AppendLine(a["CGTOMOVTO"] + ",");
                x.AppendLine("'" + a["CSCMOVTO"] + "',");
                x.AppendLine(a["CCOMVENTA"] + ",");
                x.AppendLine(a["CIDMOVTODESTINO"] + ",");
                x.AppendLine(a["CNUMEROCONSOLIDACIONES"] + ")");


                string comando = x.ToString();
                SqlCommand lsql4 = new SqlCommand(comando, _conDes);
                lsql4.ExecuteNonQuery();
                lidmovim++;

            }


            return 0;
        }


        public int mValidaSQLConexion(string server, string bd, string user, string psw)
        {
            Cadenaconexion = "data source =" + server + ";initial catalog =" + bd + ";user id = " + user + "; password = " + psw + ";";
            SqlConnection _con = new SqlConnection();
            cserver = server;
            cbd = bd;
            cusr = user;
            cpwd = psw;

            _con.ConnectionString = Cadenaconexion;
            try
            {
                _con.Open();
                // si se conecto grabar los datos en el cnf
                _con.Close();
                return 1;
            }
            catch (Exception ee)
            {
                return 0;
            }
        }

        public void mLlenarinfoMicroplane(int afolioinicial, int afoliofinal)
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
            " FROM oehdrhst_sql h WITH (NOLOCK)" +
            " join cicmpy c WITH (NOLOCK) on c.cmp_code = h.cus_no " +
            " join oelinhst_sql l WITH (NOLOCK) on l.inv_no = h.inv_no " +
            " join imitmidx_sql p WITH (NOLOCK) on p.item_no = l.item_no ";

            if (afoliofinal == 0)
                ssql += " where h.inv_no > " + afolioinicial.ToString();
            else
                ssql += " where h.inv_no >= " + afolioinicial.ToString() + " and h.inv_no <= " + afoliofinal.ToString();

            ssql += " and l.qty_to_ship > 0 ";
            ssql += " order by h.inv_no asc ";


            ssql = "SELECT  inv_dt,  h.inv_no,  c.cmp_code,c.cmp_name, c.textfield1, c.TaxCode  " +
                    ", h.curr_cd, h.curr_trx_rt  , h.bill_to_addr_1,  h.bill_to_city, h.bill_to_country,   h.bill_to_no, h.bill_to_state, h.bill_to_zip   " +
                    ", l.item_no, l.item_desc_1,     " +
                    "l.unit_price, l.discount_pct,   " +
                    "l.qty_ordered,     " +
                    "l.qty_to_ship     " +
                    ",p.item_note_1,     " +
                    "p.item_note_5,     " +
                    "p.item_note_2     " +
                    ",p.item_note_3     " +
                    "FROM oehdrhst_sql h WITH (NOLOCK) join cicmpy c WITH (NOLOCK) on c.cmp_code = h.cus_no      " +
                    "join     " +
                    "(    " +
                    "select inv_no,item_no, sum(qty_to_ship) as qty_to_ship,item_desc_1,    " +
                    "max(unit_price) as unit_price, max(discount_pct) as discount_pct ,     " +
                    "max(qty_ordered) as qty_ordered    " +
                    " from oelinhst_sql ";

            if (afoliofinal == 0)
                ssql += " where inv_no > " + afolioinicial.ToString();
            else
                ssql += " where inv_no >= " + afolioinicial.ToString() + " and inv_no <= " + afoliofinal.ToString();

            //where inv_no >= 7142 and inv_no <= 7142 and qty_to_ship > 0
            ssql += " and qty_to_ship > 0 group by inv_no,item_no,item_desc_1 " +
") as l    " +
"on  l.inv_no = h.inv_no      " +
"join imitmidx_sql p WITH (NOLOCK) on p.item_no = l.item_no      ";



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

                        string lcliente = dr["cmp_code"].ToString().Trim();

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
                            lDocto.cCodigoCliente = dr["cmp_code"].ToString().Trim();
                            lDocto.cRazonSocial = dr["cmp_code"].ToString().Trim();
                            lDocto._RegCliente.Codigo = dr["cmp_code"].ToString().Trim();
                            lDocto._RegCliente.RazonSocial = dr["cmp_name"].ToString();
                            lcliente = lDocto.cCodigoCliente;


                            // leer el texto extra 1 del cliente




                            lDocto.cCodigoConcepto = lConcepto;
                            //lDocto.cMetodoPago = "02";

                            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");
                            lDocto.cFolio = long.Parse(dr["inv_no"].ToString());

                            lDocto.cFecha = DateTime.Parse(dr["inv_dt"].ToString());

                            // lDocto.cFecha = DateTime.Today;


                            clienteleido = lcliente;
                            folioleido = lfolio;
                            lDocto.cMoneda = dr["curr_cd"].ToString();
                            lDocto.cTipoCambio = decimal.Parse(dr["curr_trx_rt"].ToString().Trim());


                            lDocto.cMetodoPago = "PPD";

                        }

                        RegMovto regmov = new RegMovto();
                        regmov.cCodigoProducto = dr["item_no"].ToString();
                        regmov._RegProducto.Nombre = dr["item_desc_1"].ToString().Trim();

                        regmov._RegProducto.noIdentificacion = dr["item_note_1"].ToString().Trim();
                        regmov._RegProducto.CodigoMedidaPesoSAT = dr["item_note_5"].ToString().Trim();
                        regmov._RegProducto.ComercioExterior = dr["item_note_2"].ToString().Trim();
                        regmov._RegProducto.UnidadMicroplaneComercioExterior = dr["item_note_3"].ToString().Trim();


                        regmov.cPorcent01 = decimal.Parse(dr["discount_pct"].ToString().Trim());
                        regmov.cUnidades = decimal.Parse(dr["qty_to_ship"].ToString().Trim());
                        regmov.cCodigoAlmacen = "1";
                        regmov.cPrecio = decimal.Parse(dr["unit_price"].ToString().Trim());
                        lDocto._RegMovtos.Add(regmov);
                    }
                    else
                    {
                        noseguir = true;
                        if (lDocto.cCodigoCliente != "")
                        {
                            _RegDoctos.Add(lDocto);
                            //lDocto = new RegDocto();
                        }
                    }

                }
            }
            dr.Close();


        }

        public string mLlenarinfoXML(string archivo)
        {
            string lFolio = "";
            try
            {
                _RegDoctos.Clear();

                DirectoryInfo dirInfo = new DirectoryInfo(@archivo);
                FileSystemInfo[] allFiles = dirInfo.GetFileSystemInfos();
                var orderedFiles = allFiles.OrderBy(f => f.Name);

                foreach (var fi in orderedFiles)
                {
                    RegDocto lDocto = new RegDocto();
                    XmlDocument xDoc = new XmlDocument();
                    xDoc.Load(archivo + "\\" + fi.Name);

                    XmlNodeList xComprobante = xDoc.GetElementsByTagName("cfdi:Comprobante");

                    string lTipoComprobante = "";
                    foreach (XmlElement nodo in xComprobante)
                    {
                        lDocto.cFecha = DateTime.Parse(nodo.GetAttribute("Fecha"));
                        /*if (lDocto.cFecha < DateTime.Now.AddHours(-72))
                            lDocto.cFecha = DateTime.Today;*/
                        string ltipocambio = nodo.GetAttribute("TipoCambio").ToString();
                        if (ltipocambio != "")
                            lDocto.cTipoCambio = Decimal.Parse(ltipocambio);
                        lDocto.cMoneda = nodo.GetAttribute("Moneda");
                        lDocto.cMetodoPago = nodo.GetAttribute("MetodoPago");

                        if (lDocto.cMoneda == "MXN")
                            lDocto.cMoneda = "Peso Mexicano";

                        lDocto.cFormaPago = nodo.GetAttribute("FormaPago");
                        lTipoComprobante = nodo.GetAttribute("TipoDeComprobante");

                        lFolio = nodo.GetAttribute("Folio");
                        lDocto.cFolio = long.Parse(lFolio);

                    }

                    XmlNodeList xEmisor = ((XmlElement)xComprobante[0]).GetElementsByTagName("cfdi:Emisor");
                    XmlNodeList xReceptor = ((XmlElement)xComprobante[0]).GetElementsByTagName("cfdi:Receptor");
                    XmlNodeList xConceptos = ((XmlElement)xComprobante[0]).GetElementsByTagName("cfdi:Conceptos");

                    XmlNodeList xComplemento = ((XmlElement)xComprobante[0]).GetElementsByTagName("cfdi:Complemento");

                    lDocto.cNombreArchivo = fi.Name;


                    foreach (XmlElement nodo in xReceptor)
                    {
                        lDocto.cRFC = nodo.GetAttribute("Rfc");
                        lDocto.cRazonSocial = nodo.GetAttribute("Nombre");
                        lDocto.cUsoCFDI = nodo.GetAttribute("UsoCFDI");
                        //lFolio = nodo.GetAttribute("Folio");
                        //long lFoliox = mBuscarUltimoFolioConcepto("4", GetSettingValueFromAppConfigForDLL("Concepto"), ref cserie);
                        string cserie = "";
                        //long lFoliox = mBuscarUltimoFolioConcepto("4", "4", ref cserie);
                        //lDocto.cFolio = long.Parse(nodo.GetAttribute("Folio").ToString());
                        lDocto.cCodigoCliente = nodo.GetAttribute("Rfc");
                        if (lTipoComprobante == "I")
                            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");

                        if (lTipoComprobante == "E")
                            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoD");

                        if (lTipoComprobante == "P")
                            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoP");

                        lDocto._RegCliente.Codigo = nodo.GetAttribute("Rfc");
                        lDocto._RegCliente.RazonSocial = nodo.GetAttribute("Nombre");

                        XmlNodeList xDomFiscal = ((XmlElement)nodo).GetElementsByTagName("cfdi:DomicilioFiscal");

                        foreach (XmlElement nodoDomFiscal in xDomFiscal)
                        {
                            lDocto._RegDireccion.cCodigoPostal = nodoDomFiscal.GetAttribute("codigoPostal");
                        }

                        lDocto.cRegimenFiscal = nodo.GetAttribute("RegimenFiscal");

                        XmlNodeList xRegFiscal = ((XmlElement)nodo).GetElementsByTagName("cfdi:RegimenFiscal");
                        foreach (XmlElement nodoRegFiscal in xRegFiscal)
                        {
                            lDocto.cRegimenFiscal = nodoRegFiscal.GetAttribute("Regimen");
                        }


                        //lDocto.cFecha = 



                    }

                    foreach (XmlElement nodoReceptor in xReceptor)
                    {
                        lDocto.cRFC = nodoReceptor.GetAttribute("Rfc");
                        lDocto.cRazonSocial = nodoReceptor.GetAttribute("Nombre");
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
                            regmov._RegProducto.Nombre = nodoConcepto.GetAttribute("Descripcion");

                            regmov.cDescuento = decimal.Parse(nodoConcepto.GetAttribute("Descuento"));

                            regmov.cUnidades = decimal.Parse(nodoConcepto.GetAttribute("Cantidad"));
                            regmov.cCodigoAlmacen = "1";
                            regmov.cPrecio = decimal.Parse(nodoConcepto.GetAttribute("ValorUnitario"));


                            //int HashCode = regmov._RegProducto.Nombre.GetHashCode();

                            regmov._RegProducto.Codigo = nodoConcepto.GetAttribute("NoIdentificacion");

                            regmov.cCodigoProducto = regmov._RegProducto.Codigo;

                            regmov._RegProducto.CodigoMedidaPesoSAT = nodoConcepto.GetAttribute("ClaveUnidad");

                            regmov._RegProducto.noIdentificacion = nodoConcepto.GetAttribute("ClaveProdServ");

                            XmlNodeList xImpuesto = ((XmlElement)nodoConcepto).GetElementsByTagName("cfdi:Impuestos");
                            foreach (XmlElement nodoImpuestos in xImpuesto)
                            {
                                XmlNodeList xTraslados = ((XmlElement)nodoImpuestos).GetElementsByTagName("cfdi:Traslados");
                                foreach (XmlElement nodoTraslados in xTraslados)
                                {
                                    XmlNodeList xTraslado = ((XmlElement)nodoTraslados).GetElementsByTagName("cfdi:Traslado");
                                    foreach (XmlElement nodoTraslado in xTraslado)
                                    {
                                        string limpuesto1 = nodoTraslado.GetAttribute("Importe");
                                        regmov.cImpuesto = decimal.Parse(limpuesto1);
                                        //regmov.cImpuesto = (regmov.cUnidades * regmov.cPrecio) * decimal.Parse(limpuesto1);
                                    }
                                }
                            }



                            lDocto._RegMovtos.Add(regmov);


                        }
                    }


                    foreach (XmlElement nodoComplemento in xComplemento)
                    {
                        XmlNodeList xpago10 = ((XmlElement)nodoComplemento).GetElementsByTagName("pago10:Pagos");
                        foreach (XmlElement nodoPago in xpago10)
                        {
                            XmlNodeList xpago = ((XmlElement)nodoPago).GetElementsByTagName("pago10:Pago");



                            foreach (XmlElement nodoPagoRel in xpago)
                            {

                                lDocto.cFecha = DateTime.Parse(nodoPagoRel.GetAttribute("FechaPago").ToString());
                                lDocto.cNeto = double.Parse(nodoPagoRel.GetAttribute("Monto").ToString());
                                XmlNodeList xpagorelacionado = ((XmlElement)nodoPago).GetElementsByTagName("pago10:DoctoRelacionado");

                                foreach (XmlElement nodoPagoRelacionado in xpagorelacionado)
                                {
                                    RegDocto cargo = new RegDocto();
                                    cargo.cSerie = nodoPagoRelacionado.GetAttribute("Serie");
                                    cargo.cFolio = long.Parse(nodoPagoRelacionado.GetAttribute("Folio").ToString());
                                    cargo.cNeto = double.Parse(nodoPagoRelacionado.GetAttribute("ImpPagado").ToString());
                                    cargo.cTipoCambio = decimal.Parse(nodoPagoRelacionado.GetAttribute("TipoCambioDR").ToString());
                                    lDocto.relacionados.Add(cargo);
                                }
                            }
                        }
                    }



                    _RegDoctos.Add(lDocto);

                }

                return "";
            }
            catch (Exception eeee)
            {
                return "Llenar Datos, Documento" + lFolio + " " + eeee.Message;
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

                if (cfdi == 0)
                    sqlstring = "select ccodigoconcepto,cnombreconcepto,cverfacele from admConceptos where ciddocumentode = " + aIdDocumentoDe + " and cescfd = 0";



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
            lsql.CommandText = "select cnombrep01 from mgw10005 where ccodigop01 = '" + codigo + "'";
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
            return lAlmacen;
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

      /*  public string mLLenarInfoPedidosFacturas(string archivo)
        {
            //string archivo1 = @archivo;
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @archivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");

            // OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @archivo + ";Extended Properties=" + Convert.ToChar(34).ToString() + @"Excel 8.0" + Convert.ToChar(34).ToString() + ";");

            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            try
            {
                conn.Open();

                cmd.Connection = conn;
                cmd.CommandText = "SELECT * FROM [Hoja1$]";

                cmd.ExecuteNonQuery();
            }
            catch (Exception eeeee)
            {
                return eeeee.Message;
            }

            System.Data.OleDb.OleDbDataReader dr;
            dr = cmd.ExecuteReader();
            Boolean noseguir = false;
            _RegDoctos.Clear();
            List<RegDocto> doctos = new List<RegDocto>();
            RegDocto lDocto = new RegDocto();
            if (dr.HasRows)
            {
                long lfolioleido = 0;
                string cserie = "";
                //dr.Read();
                long lFoliox;
                while (noseguir == false)
                {

                    dr.Read();

                    string lcliente = dr["Cliente ID"].ToString();
                    if (lcliente == "")
                        break;

                    try
                    {
                        lFoliox = long.Parse(dr["Folio dispensador"].ToString());
                    }
                    catch (Exception eee)
                    {
                        lFoliox = long.Parse(dr["Folio"].ToString());
                    }


                    if (lFoliox != lfolioleido)
                    {
                        if (lDocto.cCodigoCliente != "")
                        {
                            _RegDoctos.Add(lDocto);
                            lDocto = new RegDocto();
                        }


                        //lDocto.cSerie = cserie;
                        lDocto.cCodigoCliente = dr["Cliente ID"].ToString();
                        //lcliente = lDocto.cCodigoCliente;
                        lDocto.cCodigoConcepto = "2";


                        lDocto.cFolio = lFoliox;
                        //--lFoliox++;
                        try
                        {
                            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoP");
                            lDocto.cFecha = DateTime.Parse(dr["Fecha Ticket"].ToString());

                        }
                        catch (Exception eeeeee)
                        {
                            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");
                            lDocto.cFecha = DateTime.Parse(dr["Fecha"].ToString());
                        }
                        try
                        {
                            lDocto.cReferencia = dr["Referencia "].ToString();
                        }
                        catch (Exception iiii)
                        { }
                        lfolioleido = lFoliox;
                        lDocto.cMoneda = "Peso Mexicano";
                        lDocto.cTipoCambio = 1;
                    }

                    RegMovto regmov = new RegMovto();
                    //                    regmov.cCodigoProducto = dr["Producto"].ToString();


                    try
                    {
                        regmov.cCodigoProducto = @"001";

                        regmov.cUnidades = decimal.Parse(dr["Litros"].ToString());
                    }
                    catch (Exception eeeeee)
                    {
                        regmov.cUnidades = 1;
                        regmov.cCodigoProducto = @"(Ninguno)                     ";
                    }
                    regmov.cCodigoAlmacen = "1";

                    try
                    {
                        regmov.cPrecio = decimal.Parse(dr["Precio x Litro"].ToString());
                    }
                    catch (Exception yyyyyyy)
                    {
                        regmov.cPrecio = decimal.Parse(dr["Subtotal"].ToString());
                        regmov.cImpuesto = decimal.Parse(dr["Importe IVA"].ToString());
                        regmov.cTotal = decimal.Parse(dr["Total"].ToString());

                    }
                    //regmov.cObservaciones = dr["DESCRIPCION"].ToString();
                    lDocto._RegMovtos.Add(regmov);

                    //dr.Read();

                }


                if (lDocto.cCodigoCliente != "")

                    _RegDoctos.Add(lDocto);

            }
            return "";

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
                long lFoliox = mBuscarUltimoFolioConcepto("4", GetSettingValueFromAppConfigForDLL("Concepto"), ref cserie);
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

            using (var webClient = new System.Net.WebClient())
            {
                var json = webClient.DownloadString(url);
                // Now parse with JSON.Net
            }






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
            listanueva.Sort(delegate (RegElemento x, RegElemento y)
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
                    string tipo = newelemento.type.Substring(0, 5);
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

        List<MovimientosCartaPorte> listacartaporte = new List<MovimientosCartaPorte>();


        public string mLlenarTraslado(string archivo)
        {
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @archivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");

            // OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @archivo + ";Extended Properties=" + Convert.ToChar(34).ToString() + @"Excel 8.0" + Convert.ToChar(34).ToString() + ";");

            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            try
            {
                conn.Open();

                cmd.Connection = conn;
                cmd.CommandText = "SELECT * FROM [Hoja1$]";

                cmd.ExecuteNonQuery();
            }
            catch (Exception eeeee)
            {
                MessageBox.Show(eeeee.Message);
                return eeeee.Message;
            }

            System.Data.OleDb.OleDbDataReader dr;
            dr = cmd.ExecuteReader();
            Boolean noseguir = false;
            _RegDoctos.Clear();
            List<RegDocto> doctos = new List<RegDocto>();
            RegDocto lDocto = new RegDocto();

            long lfolioleido = 0;
            string cserie = "";
            //dr.Read();
            long lFoliox;

            string lcliente = "CL001";
            lFoliox = mBuscarUltimoFolioConcepto("4", GetSettingValueFromAppConfigForDLL("Concepto"), ref cserie);

            lDocto.cCodigoCliente = lcliente;
            lDocto.cFolio = lFoliox;
            //lcliente = lDocto.cCodigoCliente;
            lDocto.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto");


            lDocto.cFecha = DateTime.Today;

            lDocto.cMoneda = "Peso Mexicano";
            lDocto.cTipoCambio = 1;

            while (dr.HasRows == true)
            {

                dr.Read();




                RegMovto regmov = new RegMovto();
                //                    regmov.cCodigoProducto = dr["Producto"].ToString();


                try
                {

                    regmov.cCodigoProducto = dr["Codigo"].ToString();
                }
                catch (Exception eee)
                {
                    break;
                }



                regmov.cUnidades = decimal.Parse(dr["Cantidad"].ToString());

                regmov.cCodigoAlmacen = "1";

                regmov.cPrecio = decimal.Parse(dr["PrecioU"].ToString());
                regmov._RegProducto.Codigo = dr["Codigo"].ToString();
                regmov._RegProducto.Nombre = dr["Descripcion"].ToString();
                int l = 0;
                if (dr["Descripcion"].ToString() == "BTS-330/350 SOLUCION DE LAVADO 1 LT")
                    l = 1;

                regmov.cNombreProducto = dr["Descripcion"].ToString();
                regmov._RegProducto.CodigoMedidaPesoSAT = dr["Udm"].ToString();
                regmov._RegProducto.CodigoSAT = dr["CodigoSAT"].ToString();
                regmov.cUnidad = dr["UnidadMedida"].ToString();



                regmov.traslado.PesoEnKg = dr["PesoEnKg"].ToString();
                regmov.traslado.Moneda = dr["Moneda"].ToString();
                regmov.traslado.Pedimento = dr["Pedimento"].ToString().Replace(" ", "");
                regmov.traslado.ValorMercancia = dr["ValorMercancia"].ToString();

                regmov.traslado.materialpeligroso = dr["MaterialPeligroso"].ToString();
                regmov.traslado.cvematerialpeligroso = dr["ClaveSatMaterialP"].ToString();






                lDocto._RegMovtos.Add(regmov);

                //                dr.Read();

            }
            _RegDoctos.Add(lDocto);




            return "";
        }


        public int mLlenarinfoFresko(string url)
        {
            string json1 = File.ReadAllText(url);

            try
            {
                var jsonObject = JsonConvert.DeserializeObject<Root>(json1);
            }
            catch (Exception ee)
            {
                return -1;
            }
            //var files1 = JsonConvert.DeserializeObject<prods>(json1);

            RegDocto doc = new RegDocto();
            doc.cCodigoCliente = "1";
            doc.cCodigoConcepto = "5";
            doc.cFecha = DateTime.Today;

            var data = JsonConvert.DeserializeObject<Root>(json1);
            listacartaporte.Clear();
            _RegDoctos.Clear();
            foreach (Cita x in data.citas)
            {
                foreach (SecuenciasDeEntrega y in x.secuencias_de_Entrega)
                {
                    MovimientosCartaPorte mov = new MovimientosCartaPorte();




                    foreach (Contenido z in y.contenido)
                    {

                        mov = new MovimientosCartaPorte();
                        mov.PesoEnKg = z.peso.ToString();
                        mov.Cantidad = z.cantidad.ToString();
                        if (z.claveUnidadCompra == null)
                            z.claveUnidadCompra = "";
                        mov.ClaveUnidad = z.claveUnidadCompra.ToString();
                        mov.BienesTransp = z.claveSat.ToString();
                        if (z.embalaje == null)
                            z.embalaje = "";
                        mov.Embalaje = z.embalaje.ToString();
                        if (z.claveProductoPeligroso == null)
                            z.claveProductoPeligroso = "";
                        mov.CveMaterialPeligroso = z.claveProductoPeligroso.ToString();
                        mov.Descripcion = z.descripcion;
                        listacartaporte.Add(mov);

                        RegMovto movi = new RegMovto();
                        movi.cNombreProducto = z.descripcion;
                        movi.cUnidades = decimal.Parse(z.cantidad);
                        movi.cUnidad = z.claveUnidadCompra;
                        movi.cCodigoAlmacen = "1";
                        doc._RegMovtos.Add(movi);
                    }


                }


            }
            doc.listacartaporte = listacartaporte;
            _RegDoctos.Add(doc);




            //selectedCollection = data.ToList();

            return 1;

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
            connect.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.GetDirectoryName(archivo1) + ";Extended Properties='Text;HDR=Yes;FMT=Delimited;IMEX=1';Persist Security Info=False";
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
            mRegresarPrincipales(lCodigoConcepto, ref lidconce, ref tipocfd, ref cserie, ref lnat);
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
                lresp = mGrabarMovimientos(lIdDocumento, 3, 0);
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

        public List<string> mGrabarDoctosFresko(bool incluyetimbrado, int tipo)
        {
            List<string> lista = new List<string>();
            String cseriex = "";
            String cseriey = "";
            long lFoliox = mBuscarUltimoFolioConcepto("4", "5", ref cseriex);
            long lFolioy = mBuscarUltimoFolioConcepto("4", "4", ref cseriey);
            miconexion.mAbrirConexionDestino(1);
            long aFolio = 0;
            aFolio = 1;
            _RegDoctoOrigen = _RegDoctos[0];
            _RegDoctoOrigen.cCodigoConcepto = "5";
            _RegDoctoOrigen.cSerie = cseriex;




            string lresp2 = mGrabarAdmNew(lFoliox, 1, incluyetimbrado, tipo);
            if (lresp2 != "")
                lista.Add(lresp2);

            _RegDoctoOrigen._RegMovtos.Clear();
            _RegDoctoOrigen.cCodigoConcepto = "4";
            _RegDoctoOrigen.cSerie = cseriey;
            lresp2 = mGrabarAdmNew(lFolioy, 1, incluyetimbrado, tipo);
            if (lresp2 != "")
                lista.Add(lresp2);

            //generar .ini
            mGrabarIni(lFolioy);



            miconexion.mCerrarConexionOrigen(1);
            return lista;
        }


        public List<string> mGrabarDoctosTraslado(bool incluyetimbrado, int tipo)
        {
            List<string> lista = new List<string>();
            String cseriex = "";
            String cseriey = "";
            miconexion.mAbrirConexionDestino(1);
            long aFolio = 0;
            aFolio = 1;
            _RegDoctoOrigen = _RegDoctos[0];




            string lresp2 = mGrabarTraslado(_RegDoctoOrigen.cFolio, 1, incluyetimbrado, tipo);
            if (lresp2 != "")
                lista.Add(lresp2);



            miconexion.mCerrarConexionOrigen(1);
            return lista;
        }


        private void mGrabarIni(long lFolioy)
        {
            //string jsonFilePath = @"C:\Users\147026\Downloads\INTERFAZ COMPLEMENTO CARTA PORTE\INTERFAZ\FRESKO\Carta.ini";

            string lcarpeta = miconexion.rutadestino;
            string larchivo = "CFDI 3.3 INGRESOS COMPLEMENTO CARTA PORTE" + "__" + lFolioy.ToString() + ".ini";
            string targetfile = Path.Combine(lcarpeta, larchivo);
            StreamWriter sw = new StreamWriter(targetfile);
            int index = 1;
            foreach (MovimientosCartaPorte mov in _RegDoctoOrigen.listacartaporte)
            {

                sw.WriteLine("[" + index.ToString() + "]");
                sw.WriteLine("BienesTransp=" + mov.BienesTransp);
                sw.WriteLine("Descripcion=" + mov.Descripcion);
                sw.WriteLine("Cantidad=" + mov.Cantidad);
                sw.WriteLine("ClaveUnidad=" + mov.ClaveUnidad);
                sw.WriteLine("Unidad=" + mov.Unidad);
                sw.WriteLine("CveMaterialPeligroso=" + mov.CveMaterialPeligroso);
                sw.WriteLine("Embalaje=" + mov.Embalaje);
                sw.WriteLine("DescripEmbalaje=" + mov.DescripEmbalaje);
                sw.WriteLine("PesoEnKg=" + mov.PesoEnKg);
                sw.WriteLine("ValorMercancia=" + mov.ValorMercancia);
                sw.WriteLine("Moneda=MXN");
                sw.WriteLine("FraccionArancelaria=" + mov.FraccionArancelaria);
                sw.WriteLine("UUIDComercioExt=" + mov.UUIDComercioExt);
                sw.WriteLine("Pedimentos=" + mov.Pedimentos);
                index++;

            }

            sw.Close();
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
            _con = new OleDbConnection();



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
            string lresp = mGrabarMovimientos(lIdDocumento, 3, 0);
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

        private double mRegresarFolio(int lDocumentoModelo)
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
                catch (Exception ee)
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
        protected string mModificaDatosClienteFlexo()
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

        protected string mLlenarDoctos(OleDbDataReader aReader)
        {
            _RegDoctos.Clear();
            string lfolio = "";
            aReader.Read();
            int lbandera = 1;
            while (lbandera == 1 && aReader.HasRows)
            {
                RegDocto x = new RegDocto();
                List<RegMovto> movtos = new List<RegMovto>();

                x.cAgente = "(Ninguno)";
                try
                {
                    x.cReferencia = aReader["cReferen01"].ToString();
                }
                catch (Exception dddd)
                {

                }
                x.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento");
                x.cFecha = DateTime.Parse(aReader["cfecha"].ToString());
                x.sMensaje = "";
                x.cMoneda = "Pesos";
                x.cTextoExtra1 = aReader["cObserva01"].ToString();

                string sfoliodocto = aReader["cfolio"].ToString();
                long lfoliodocto = 0;
                string lserie = "";
                try
                {
                    lfoliodocto = long.Parse(aReader["cfolio"].ToString());

                }
                catch (Exception eee)
                {
                    lserie = sfoliodocto.Substring(sfoliodocto.Length - 1);
                    sfoliodocto = sfoliodocto.Substring(0, sfoliodocto.Length - 1);
                }



                x.cFolio = long.Parse(sfoliodocto);
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
            string lfolio = "0";
            if (atipo == 1 || atipo == 2)
            {
                lfolio = aReader["cfolio"].ToString();
                _RegDoctoOrigen.cFolio = long.Parse(lfolio);
            }
            if (aReader["cliente"].ToString() == string.Empty)
                return "Falta Codigo de cliente en documento " + aFolio;
            else
                _RegDoctoOrigen.cCodigoCliente = aReader["cliente"].ToString();

            _RegDoctoOrigen.cFecha = DateTime.Parse(aReader["cfecha"].ToString());
            _RegDoctoOrigen.cFecha = DateTime.Parse(DateTime.Today.ToString());
            if (mchecarvalido() == false)
                return "";



            //_RegDoctoOrigen.cFolio = long.Parse (aReader["cfolio"].ToString()) ;
            if (aReader["cRFC"].ToString() == string.Empty)
                return "Cliente sin RFC en documento " + aFolio;
            else
                if (!(aReader["cRFC"].ToString().Length == 12 || aReader["cRFC"].ToString().Length == 13))
                return "El RFC tiene una longitud incorrecta en el documento " + aFolio;
            else
                _RegDoctoOrigen.cRFC = aReader["cRFC"].ToString();


            if (atipo == 1)
            {
                _RegDoctoOrigen.cAgente = aReader["Agente"].ToString();
                _RegDoctoOrigen.cCond = aReader["condpago"].ToString();

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
            _RegDoctoOrigen.cTipoCambio = decimal.Parse(aReader["TipoCambio"].ToString());

            if (atipo != 1)
                _RegDoctoOrigen.cReferencia = aReader["cReferen01"].ToString();
            else
                _RegDoctoOrigen.cReferencia = aReader["cReferen01"].ToString();



            if (aReader["cnombrec01"].ToString().Trim() == string.Empty)
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


            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;

            lsql.CommandText = mRegresarConsultaMovimientos(aFuente, lfolio, atipo);


            lsql.Connection = (OleDbConnection)_con;
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
        public Boolean mBuscar(long aFolio, string aConcepto, string aSerie, int aTipo)
        {
            Boolean lRespuesta = false;
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            OleDbParameter lparametrofolio = new OleDbParameter("@p2", aFolio);
            OleDbParameter lparametrodocumentode = new OleDbParameter("@p1", aConcepto);

            lsql.CommandText = "Select m2.ccodigoc01 as cliente,m6.ccodigoc01 as concepto, m6.cidconce01, m8.cfecha,m8.cfolio, m8.ciddocum01 " +
                " from mgw10008 m8 join mgw10002 m2 on m8.cidclien01 = m2.cidclien01 " +
                " join mgw10006 m6 on m8.cidconce01 = m6.cidconce01 " +
                " and m6.ccodigoc01 =  '" + aConcepto + "'" +
                " where cfolio = " + aFolio +
            " and cseriedo01 = '" + aSerie + "'";
            //lsql.Parameters.Add(lparametrodocumentode);
            //lsql.Parameters.Add(lparametrofolio);
            if (aTipo == 0)
                lsql.Connection = miconexion.mAbrirConexionOrigen();
            else
                lsql.Connection = miconexion.mAbrirConexionDestino();


            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lreader.Read();
                //mLlenarDocto(lreader);

                lRespuesta = true;

            }
            miconexion.mCerrarConexionOrigen();
            lreader.Close();
            return lRespuesta;






        }



        public string mGrabarDestinos()
        {
            ClassConexion miconexion = new ClassConexion();
            string lregresa = "";
            miconexion.aRutaExe = aRutaExe;
            miconexion.mAbrirConexionDestino(1);
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
            lret = fSetDatoDocumento("cFecha", lfechavenc);
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
            string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoFactura").ToString().Trim();
            long lIdDocumento;
            RegProveedor lRegProveedor = new RegProveedor();
            lRegProveedor = mBuscarCliente(GetSettingValueFromAppConfigForDLL("Cliente").ToString().Trim(), 0, 0);


            fInsertarDocumento();
            lret = fSetDatoDocumento("cFecha", DateTime.Today.ToString());
            lret = fSetDatoDocumento("cCodigoConcepto", lCodigoConcepto);
            lret = fSetDatoDocumento("cSerieDocumento", GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim());
            lret = fSetDatoDocumento("cCodigoCteProv", lRegProveedor.Codigo);
            lret = fSetDatoDocumento("cRazonSocial", lRegProveedor.RazonSocial);
            lret = fSetDatoDocumento("cRFC", lRegProveedor.RFC);
            lret = fSetDatoDocumento("cIdMoneda", "1");
            lret = fSetDatoDocumento("cTipoCambio", "1");
            lret = fSetDatoDocumento("cReferencia", "Por Programa");
            lret = fSetDatoDocumento("cFolio", GetSettingValueFromAppConfigForDLL("FolioFactura").ToString().Trim());
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
            lRegDireccion = mBuscarDireccion(lRegProveedor.Id, 0);

            if (!string.IsNullOrEmpty(lRegDireccion.cNombreCalle))
            {
                lret = fInsertaDireccion();
                lret = fSetDatoDireccion("cIdCatalogo", lIdDocumento.ToString());
                lret = fSetDatoDireccion("cTipoCatalogo", "3");
                lret = fSetDatoDireccion("cTipoDireccion", "0");
                lret = fSetDatoDireccion("cNombreCalle", lRegDireccion.cNombreCalle);
                lret = fSetDatoDireccion("cNumeroExterior", lRegDireccion.cNumeroExterior);
                lret = fSetDatoDireccion("cNumeroInterior", lRegDireccion.cNumeroInterior);
                lret = fSetDatoDireccion("cColonia", lRegDireccion.cColonia);
                lret = fSetDatoDireccion("cCodigoPostal", lRegDireccion.cCodigoPostal);
                lret = fSetDatoDireccion("cEstado", lRegDireccion.cEstado);
                lret = fSetDatoDireccion("cPais", lRegDireccion.cPais);
                lret = fSetDatoDireccion("cCiudad", lRegDireccion.cCiudad);
                lret = fGuardaDireccion();
            }


            long lNumeroMov = 100;

            foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
            {
                //barra.Avanzar();
                lret = fInsertarMovimiento();
                lret = fSetDatoMovimiento("cIdDocumento", lIdDocumento.ToString());
                lret = fSetDatoMovimiento("cNumeroMovimiento", lNumeroMov.ToString());

                lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
                lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);
                lret = fSetDatoMovimiento("cUnidades", x.cUnidades.ToString());
                lret = fSetDatoMovimiento("cPrecio", x.cPrecio.ToString());
                //lret = fSetDatoMovimiento("cPorcentajeImpuesto1", z.Cells[17].Value.ToString());
                //w = decimal.Parse(z.Cells[4].Value.ToString()) * decimal.Parse(z.Cells[6].Value.ToString());
                lret = fSetDatoMovimiento("cImpuesto1", x.cImpuesto.ToString());

                lret = fGuardaMovimiento();
                lNumeroMov += 100;

            }
            long lrespuesta = 0;
            if (lret == 0)
                lrespuesta = fAfectaDocto_Param(lCodigoConcepto, GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim(), double.Parse(_RegDoctoOrigen.cFolio.ToString()), true);
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
            OleDbConnection lconexion = new OleDbConnection();
            if (aTipo == 0)
                lconexion = miconexion.mAbrirConexionOrigen();
            else
                lconexion = miconexion.mAbrirConexionDestino();

            string lcadena = "select m8.ciddocum01,m2.crazonso01, m2.crfc from mgw10008 m8 " +
            " join mgw10002 m2 on m8.cidclien01 = m2.cidclien01 " +
            " join mgw10006 m6 on m8.cidconce01 = m6.cidconce01 " +
            " where m6.ccodigoc01 = '" + aConcepto + "' and m8.cfolio = " + afolio.ToString() +
            " and cseriedo01 = '" + aSerie + "'";

            OleDbCommand lsql = new OleDbCommand(lcadena, lconexion);
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

            return lIdDocumento;

        }
        private RegDireccion mBuscarDireccion(long aCliente, int aTipo)
        {
            string sql;
            OleDbConnection lconexion = new OleDbConnection();
            RegDireccion lreg = new RegDireccion();
            lconexion = miconexion.mAbrirConexionOrigen();
            sql = "select * from mgw10011 where cidcatal01 = " + aCliente +
                        " and ctipocat01 = 1 and ctipodir01 = " + aTipo;
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

            return lreg;

        }

        protected virtual string GetSettingValueFromAppConfigForDLL(string aNombreSetting)
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
                Directory.SetCurrentDirectory(lrutadminpaq);
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

        public List<RegConcepto> mCargarConceptos(long aIdDocumentoDe, int aTipo, int cfdi, int cartaporte = 0)
        {
            List<RegConcepto> _RegFacturas = new List<RegConcepto>();
            OleDbConnection lconexion = new OleDbConnection();
            if (aTipo == 0)
                lconexion = miconexion.mAbrirConexionOrigen();
            else
                lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {


                string sqlstring = "select ccodigoc01,cnombrec01,cverfacele from mgw10006 where ciddocum01 = " + aIdDocumentoDe;
                if (cfdi == 1)
                    sqlstring = "select ccodigoc01,cnombrec01,cverfacele from mgw10006 where ciddocum01 = " + aIdDocumentoDe + " and cescfd = 1";

                if (cartaporte == 1)
                    sqlstring = "select ccodigoc01,cnombrec01,cverfacele from mgw10006 where ciddocum01 = 4 and ccartapor = 1";

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

        public List<RegOrigen> mCargarDocumentosComercialDoctoDeCliente(int aDocumentoDe, long aIdCliente)
        {
            List<RegOrigen> _RegOrigenes = new List<RegOrigen>();
            SqlConnection lconexion = new SqlConnection();
            lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {

                //OleDbCommand lsql = new OleDbCommand("select ccodigoc01,cnombrec01 from mgw10006 where ciddocum01 = " + aIdDocumentoDe + " and cescfd = 1 and cnombrec01 = 'CFDI'", lconexion);
                // este es para flexo
                SqlCommand lsql = new SqlCommand("select  FORMAT(m8.cfecha,'dd-MM-yyyy') as cfecha, m8.cfolio, m2.crazonsocial crazonso01, m2.cidclienteproveedor cidclien01, m8.ctotal,m2.ccodigocliente as cliente, " +
                    //"m10.cprecioc01 as precio, " +
                    "m2.crfc as rfc, m8.ciddocumento " + //, m10.cunidades as unidades, m10.cneto as TotalMov " +
                    "from admdocumentos m8 " +
//" join mgw10010 m10 on m10.ciddocum01 = m8.ciddocum01 " +
" join admclientes m2 on m2.cidclienteproveedor = m8.cidclienteproveedor " +
" where m2.cidclienteproveedor = " + aIdCliente.ToString() +
" and m8.ciddocumentode = " + aDocumentoDe.ToString() +
" AND m8.ccancelado = 0 and m8.cunidadespendientes > 0 ", lconexion);

                RegOrigen lRegOrigen = new RegOrigen();

                //lsql.Parameters.Add("@folio", 
                SqlDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _RegOrigenes.Clear();
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        lRegOrigen = new RegOrigen();
                        /*lRegOrigen.CodigoProducto = lreader["ccodigop01"].ToString().Trim();
                        lRegOrigen.NombreProducto = lreader["cnombrep01"].ToString().Trim();
                        lRegOrigen.cidproducto = int.Parse(lreader["cidprodu01"].ToString());
                        lRegOrigen.IEPS = decimal.Parse(lreader["cimpuesto2"].ToString());
                        lRegOrigen.IEPS2 = decimal.Parse(lreader["cimpuesto3"].ToString());
                        lRegOrigen.Descuento = decimal.Parse(lreader["cimporte01"].ToString());
                        */
                        lRegOrigen.cTotal = Math.Round(decimal.Parse(lreader["cTotal"].ToString()), 2);
                        lRegOrigen.Folio = lreader["cFolio"].ToString();
                        lRegOrigen.Fecha = lreader["cFecha"].ToString();

                        lRegOrigen.cIdClien01 = long.Parse(lreader["cidclien01"].ToString());
                        lRegOrigen.RazonSocial = lreader["cRazonSo01"].ToString();
                        lRegOrigen.CodigoCliente = lreader["cliente"].ToString();
                        //lRegOrigen.Precio = Math.Round(decimal.Parse(lreader["precio"].ToString()), 2);
                        //lRegOrigen.Precio2 = Math.Round(decimal.Parse(lreader["precio"].ToString()), 2);
                        //lRegOrigen.TotalMov = Math.Round(decimal.Parse(lreader["TotalMov"].ToString()), 2);
                        //lRegOrigen.Cantidad = Math.Round(decimal.Parse(lreader["Unidades"].ToString()), 2);
                        //lRegOrigen.TotalMov2 = decimal.Parse(lreader["TotalMov"].ToString());
                        lRegOrigen.ciddocumento = int.Parse(lreader["ciddocumento"].ToString());


                        lRegOrigen.CodigoCliente = lreader["cliente"].ToString();
                        lRegOrigen.RFC = lreader["rfc"].ToString();

                        _RegOrigenes.Add(lRegOrigen);
                    }
                }

                lreader.Close();
                if (aDocumentoDe == 2 && lRegOrigen.Folio != null)

                {
                    lsql = new SqlCommand("select  m8.cpendiente from admdocumentos m8 " +
" where m8.cfolio = " + lRegOrigen.Folio +
" and m8.ciddocumentode = 4", lconexion);
                    lreader = lsql.ExecuteReader();
                    if (lreader.HasRows)

                        if (lreader.Read())
                            lRegOrigen.cpendiente = double.Parse(lreader["cpendiente"].ToString());
                }
                lreader.Close();
            }


            return _RegOrigenes;



        }


        public List<RegDocto> mCargarDocumentosComercialReferencia(string aConcepto, string aReferencia, ref DataTable dt, ref DataTable dt2, int pt = 0)
        {
            List<RegDocto> _RegDocto = new List<RegDocto>();
            SqlConnection lconexion = new SqlConnection();
            lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {


                string sql = "select *, CIMPORTEEXTRA2*cuantos unidadessalidamt  from (" +
                    "select *  , round(min(x.division) over(partition by cfolio, ctextoextra3 order by ctextoextra3),0,1) cuantos " +
                " from " +
                "( " +
                    "select d.cfolio, p.crazonsocial, d.CREFERENCIA as pedido," +
                    "pr.CNOMBREPRODUCTO, " +
                    " a.cidalmacen cIdAlmacen " +
                    " ,m.cunidades, m.cprecio  " +
" , p.CCODIGOCLIENTE, pr.ccodigoproducto, a.CCODIGOALMACEN, m.cidmovimiento, m.cimporteextra1, a.cnombrealmacen, m.ctextoextra3, m.cimporteextra2 " +
" , ROW_NUMBER() over(partition by d.cfolio, m.ctextoextra3 order by m.ctextoextra3) orden " +
" , sum(m.CIMPORTEEXTRA1*m.cprecio) over(partition by d.cfolio, m.ctextoextra3 order by m.ctextoextra3) costo ";

                if (aConcepto == "340")
                    sql += " , mped.cprecio cpreciopedido, cped.ccodigocliente ccodigoclientepedido";

                //  if (aConcepto != "340")
                //sql += " , min(isnull(m.CIMPORTEEXTRA1,1) / isnull(m.CIMPORTEEXTRA2,1)) over(partition by d.cfolio, m.ctextoextra3 order by m.ctextoextra3) cuantos ";
                sql += ",isnull(m.CIMPORTEEXTRA1, 1) / isnull(m.CIMPORTEEXTRA2, 1) division ";

                sql += " from admdocumentos d " +
" join admConceptos c on d.cidconceptodocumento = c.cidconceptodocumento " +
" join admClientes p on p.CIDCLIENTEPROVEEDOR = d.CIDCLIENTEPROVEEDOR " +
" join admMovimientos m on d.CIDDOCUMENTO = m.CIDDOCUMENTO " +
" join admproductos pr on pr.cidproducto = m.cidproducto" +
" join admAlmacenes a on a.CIDALMACEN = m.CIDALMACEN";

                if (aConcepto == "340")
                    sql += " join admdocumentos ped on ped.cfolio = d.CREFERENCIA and ped.CIDDOCUMENTODE = 2 " +
" join admMovimientos mped on ped.CIDDOCUMENTO = mped.CIDDOCUMENTO " +
" and m.CIDPRODUCTO = mped.CIDPRODUCTO " +
" join admClientes cped on cped.CIDCLIENTEPROVEEDOR = ped.CIDCLIENTEPROVEEDOR ";


                sql += " where c.CCODIGOCONCEPTO = '" + aConcepto + "'" +
" and d.CREFERENCIA = '" + aReferencia + "'" +
" and m.cimporteextra1 > 0 " +
" ) x " +
" ) y " +
" order by cfolio";


                if (pt == 1)
                {
                    sql = "select *, CIMPORTEEXTRA2*cuantos unidadessalidamt " +
" from " +
" ( " +
" select * " +
" , round(min(y.division) over(partition by cfolio, ctextoextra3 order by ctextoextra3), 0, 1) cuantos " +
" from( " +
"     select * " +
"     , isnull(x.cimporteextra1, 1) / isnull(x.CIMPORTEEXTRA2, 1) division " +
"     , max(costocalc) over(order by x.ctextoextra3) costo " +
"     from " +
"     ( " +
" select * from " +
" ( " +
" select d.cfolio, p.crazonsocial, d.CREFERENCIA as pedido, pr.CNOMBREPRODUCTO, " +
" a.cidalmacen cIdAlmacen, m.cunidades, m.cprecio " +
" , p.CCODIGOCLIENTE, pr.ccodigoproducto, a.CCODIGOALMACEN, " +
" m.cidmovimiento,  a.cnombrealmacen, m.ctextoextra3, m.cimporteextra2, " +
" ROW_NUMBER() over(partition by d.cfolio, m.ctextoextra3 order by m.ctextoextra3) orden, " +
" sum(m.CIMPORTEEXTRA1 * m.cprecio) over(order by pr.ccodigoproducto) costocalc, " +
"  sum(m.CIMPORTEEXTRA1) over(partition by ccodigoproducto) cimporteextra1" +
" , ROW_NUMBER() over(partition by m.ctextoextra3, ccodigoproducto order by m.ctextoextra3, cfolio) orden2 " +
" from admdocumentos d " +
"     join admConceptos c on d.cidconceptodocumento = c.cidconceptodocumento " +
"     join admClientes p on p.CIDCLIENTEPROVEEDOR = d.CIDCLIENTEPROVEEDOR " +
"     join admMovimientos m on d.CIDDOCUMENTO = m.CIDDOCUMENTO " +
"     join admproductos pr on pr.cidproducto = m.cidproducto " +
"     join admAlmacenes a on a.CIDALMACEN = m.CIDALMACEN " +
"     where c.CCODIGOCONCEPTO = '" + aConcepto + "' and d.CREFERENCIA = '" + aReferencia + "' and m.cimporteextra1 > 0 " +
" ) a  " +
" where a.orden2 = 1) " +
"     x  " +
" ) y  " +
" ) z order by cfolio ";

                }

                string sql2 = " select p.crazonsocial,cfecha" +
                "                 from admdocumentos d" +
                " join admConceptos c on d.cidconceptodocumento = c.cidconceptodocumento" +
                " join admClientes p on p.CIDCLIENTEPROVEEDOR = d.CIDCLIENTEPROVEEDOR" +
                " where c.CCODIGOCONCEPTO = '2' and d.cfolio = " + aReferencia;



                SqlCommand lsql = new SqlCommand(sql);

                SqlCommand lsql2 = new SqlCommand(sql2);



                RegDocto lRegDocto = new RegDocto();
                lsql.Connection = miconexion._conexion1;
                SqlDataReader lreader;



                DataSet ds = new DataSet();
                DataTable dt11 = ds.Tables.Add("uno");


                DataTable dt21 = ds.Tables.Add("dos");

                SqlDataAdapter adapter = new SqlDataAdapter(lsql.CommandText, lsql.Connection);
                SqlDataAdapter adapter2 = new SqlDataAdapter(lsql2.CommandText, lsql.Connection);

                adapter.Fill(dt11);

                adapter2.Fill(dt21);


                dt = dt11;
                dt2 = dt21;



                lreader = lsql.ExecuteReader();
                _RegDocto.Clear();
                long lfolio = 0;
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {

                        if (lfolio != long.Parse(lreader["cfolio"].ToString().Trim()))
                        {
                            if (lfolio > 0)
                            {
                                _RegDocto.Add(lRegDocto);
                            }
                            lRegDocto = new RegDocto();
                            lRegDocto.cFolio = -1;
                            lRegDocto.cFecha = System.DateTime.Now;
                            lRegDocto.cCodigoConcepto = "21";
                            lRegDocto.cCodigoCliente = lreader["CCODIGOCLIENTE"].ToString().Trim();
                        }

                        RegMovto m = new RegMovto();
                        m.cCodigoAlmacen = lreader["CCODIGOALMACEN"].ToString().Trim();
                        m.cCodigoProducto = lreader["CCODIGOPRODUCTO"].ToString().Trim();
                        m.cUnidades = decimal.Parse(lreader["cunidades"].ToString().Trim());
                        m.cUnidades = decimal.Parse(lreader["cimporteextra1"].ToString().Trim());
                        m.cPrecio = decimal.Parse(lreader["cprecio"].ToString().Trim());
                        m.ctextoextra3 = lreader["Ctextoextra3"].ToString().Trim();
                        m.cimporteextra2 = decimal.Parse(lreader["cimporteextra2"].ToString().Trim());
                        m.cDescuento = decimal.Parse(lreader["costo"].ToString().Trim());
                        lRegDocto._RegMovtos.Add(m);


                        lfolio = long.Parse(lreader["cfolio"].ToString().Trim());





                        /*lRegOrigen.CodigoProducto = lreader["ccodigop01"].ToString().Trim();
                        */
                    }
                    if (lfolio > 0)
                    {
                        _RegDocto.Add(lRegDocto);
                    }
                }

                lreader.Close();
            }
            return _RegDocto;

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
                        lRegOrigen.IEPS = decimal.Parse(lreader["cimpuesto2"].ToString());
                        lRegOrigen.IEPS2 = decimal.Parse(lreader["cimpuesto3"].ToString());
                        lRegOrigen.Descuento = decimal.Parse(lreader["cimporte01"].ToString());

                        lRegOrigen.cTotal = Math.Round(decimal.Parse(lreader["cTotal"].ToString()), 2);



                        lRegOrigen.cIdClien01 = long.Parse(lreader["cidclien01"].ToString());
                        lRegOrigen.RazonSocial = lreader["cRazonSo01"].ToString();
                        lRegOrigen.CodigoCliente = lreader["cliente"].ToString();
                        lRegOrigen.Precio = Math.Round(decimal.Parse(lreader["precio"].ToString()), 2);
                        lRegOrigen.Precio2 = Math.Round(decimal.Parse(lreader["precio"].ToString()), 2);
                        lRegOrigen.TotalMov = Math.Round(decimal.Parse(lreader["TotalMov"].ToString()), 2);
                        lRegOrigen.Cantidad = Math.Round(decimal.Parse(lreader["Unidades"].ToString()), 2);
                        //lRegOrigen.TotalMov2 = decimal.Parse(lreader["TotalMov"].ToString());

                        if (lRegOrigen.Descuento == 0)
                        {
                            // precio facturado - precio capturado * unidades facturadas 

                            lRegOrigen.DescuentoAplicar = 0;

                        }
                        else
                        {
                            // precio facturado * unidades facturadas * descuento
                            lRegOrigen.DescuentoAplicar = Math.Round(lRegOrigen.Precio * lRegOrigen.Cantidad * (lRegOrigen.Descuento / 100));
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
                        _RegProveedores.Add(lRegCliente);
                    }
                }
                lreader.Close();
            }

            return _RegProveedores;



        }


        public decimal mSaldoClienteComercial(long lIdCliente)
        {

            decimal lregresa = 0.0M;
            SqlConnection lconexion = new SqlConnection();

            lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {
                string lsqls = "select sum(abca.ctotalcargo) - sum(abca.ctotalabono) as saldo, abca.CLIMITECREDITOCLIENTE, abca.CBANVENTACREDITO " +
                " from " +
                " ( " +
                "     select " +
                "     case when CNATURALEZA = 0 then ctotal else 0 end ctotalcargo, " +
                "     case when CNATURALEZA = 1 then ctotal else 0 end ctotalabono, " +
                "     c.CLIMITECREDITOCLIENTE " +
                "     , c.CBANVENTACREDITO " +
                "     from admDocumentos d " +
                "     join admClientes c on d.CIDCLIENTEPROVEEDOR = c.CIDCLIENTEPROVEEDOR " +

                "     where CNATURALEZA != 2 " +

                "     and d.cidclienteproveedor = " + lIdCliente +
                " ) abca " +
                " group by abca.CLIMITECREDITOCLIENTE, abca.CBANVENTACREDITO";

                SqlCommand lsql = new SqlCommand(lsqls, lconexion);
                SqlDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        lregresa = long.Parse(lreader[0].ToString());
                    }
                }
                lreader.Close();
            }

            return lregresa;


        }

        public List<RegProveedor> mCargarClientesComercial()
        {
            List<RegConcepto> _RegFacturas = new List<RegConcepto>();
            //RegCliente lRegCliente = new RegCliente();
            SqlConnection lconexion = new SqlConnection();

            List<RegProveedor> _regclientes = new List<RegProveedor>();
            lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {

                SqlCommand lsql = new SqlCommand("select cidclienteproveedor, ccodigocliente,crazonsocial,cbanventacredito, CLIMITECREDITOCLIENTE from admclientes where ctipocliente <=2 ", lconexion);
                SqlDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _regclientes.Clear();

                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegProveedor lRegCliente = new RegProveedor();
                        lRegCliente.Id = long.Parse(lreader[0].ToString());
                        lRegCliente.Codigo = lreader[1].ToString();
                        lRegCliente.RazonSocial = lreader[2].ToString();
                        lRegCliente.BanVentaCredito = int.Parse(lreader[3].ToString());
                        lRegCliente.LimiteCredito = decimal.Parse(lreader[4].ToString());
                        //lRegCliente.Tipocfd = lreader[2].ToString();
                        _regclientes.Add(lRegCliente);
                    }
                }
                lreader.Close();
            }

            return _regclientes;


        }


        public List<RegProveedor> mCargarProveedoresComercial()
        {
            List<RegConcepto> _RegFacturas = new List<RegConcepto>();
            //RegCliente lRegCliente = new RegCliente();
            SqlConnection lconexion = new SqlConnection();

            List<RegProveedor> _regclientes = new List<RegProveedor>();
            lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {

                SqlCommand lsql =
                    new SqlCommand("select cidclienteproveedor, ccodigocliente,crazonsocial,cbanventacredito, CLIMITECREDITOCLIENTE from admclientes where (ctipocliente >=2 or cidclienteproveedor = 0)", lconexion);
                SqlDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _regclientes.Clear();

                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegProveedor lRegCliente = new RegProveedor();
                        lRegCliente.Id = long.Parse(lreader[0].ToString());
                        lRegCliente.Codigo = lreader[1].ToString();
                        lRegCliente.RazonSocial = lreader[2].ToString();
                        lRegCliente.BanVentaCredito = int.Parse(lreader[3].ToString());
                        lRegCliente.LimiteCredito = decimal.Parse(lreader[4].ToString());
                        //lRegCliente.Tipocfd = lreader[2].ToString();
                        _regclientes.Add(lRegCliente);
                    }
                }
                lreader.Close();
            }

            return _regclientes;


        }


        public List<RegCliente> mCargarSeriesPedidosComercial(long aidmovimiento)
        {
            SqlConnection lconexion = new SqlConnection();

            List<RegCliente> _regseries = new List<RegCliente>();
            lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {
                string sql =
"select sp.cidserie, CNUMEROSERIE, CPEDIMENTO from admMovmientosSeriePedido sp " +
"join admNumerosSerie s on sp.cidserie = s.cidserie " +
"where sp.cidmovimiento = " + aidmovimiento.ToString();
                SqlCommand lsql =
                    new SqlCommand(sql, lconexion);
                SqlDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _regseries.Clear();

                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegCliente lRegSerie = new RegCliente();
                        lRegSerie.Id = long.Parse(lreader[0].ToString());
                        lRegSerie.Codigo = lreader[1].ToString();
                        lRegSerie.RazonSocial = lreader[2].ToString();
                        _regseries.Add(lRegSerie);
                    }
                }
                lreader.Close();
            }

            return _regseries;

        }

        public List<RegProveedor> mCargarAgentesComercial()
        {
            SqlConnection lconexion = new SqlConnection();

            List<RegProveedor> _regagentes = new List<RegProveedor>();
            lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {

                SqlCommand lsql = new SqlCommand("select cidagente, ccodigoagente,cnombreagente from admagentes", lconexion);
                SqlDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _regagentes.Clear();

                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegProveedor lRegAgente = new RegProveedor();
                        lRegAgente.Id = long.Parse(lreader[0].ToString());
                        lRegAgente.Codigo = lreader[1].ToString();
                        lRegAgente.RazonSocial = lreader[2].ToString();
                        _regagentes.Add(lRegAgente);
                    }
                }
                lreader.Close();
            }

            return _regagentes;


        }


        public List<RegProveedor> mCargarProductosComercial()
        {
            SqlConnection lconexion = new SqlConnection();

            List<RegProveedor> _regproductos = new List<RegProveedor>();
            lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {

                SqlCommand lsql = new SqlCommand("select cidproducto, ccodigoproducto,cnombreproducto from admproductos", lconexion);
                SqlDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _regproductos.Clear();

                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegProveedor lRegProducto = new RegProveedor();
                        lRegProducto.Id = long.Parse(lreader[0].ToString());
                        lRegProducto.Codigo = lreader[1].ToString();
                        lRegProducto.RazonSocial = lreader[2].ToString();
                        //lRegCliente.Tipocfd = lreader[2].ToString();
                        _regproductos.Add(lRegProducto);
                    }
                }
                lreader.Close();
            }

            return _regproductos;


        }
        public List<RegProveedor> mCargarAlmacenesComercial()
        {



            List<RegConcepto> _RegFacturas = new List<RegConcepto>();
            //RegCliente lRegCliente = new RegCliente();
            SqlConnection lconexion = new SqlConnection();

            List<RegProveedor> _regclientes = new List<RegProveedor>();
            lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {

                SqlCommand lsql = new SqlCommand("select cidalmacen, ccodigoalmacen,cnombrealmacen from admalmacenes ", lconexion);
                SqlDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _regclientes.Clear();

                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegProveedor lRegCliente = new RegProveedor();
                        lRegCliente.Id = long.Parse(lreader[0].ToString());
                        lRegCliente.Codigo = lreader[1].ToString();
                        lRegCliente.RazonSocial = lreader[2].ToString();
                        _regclientes.Add(lRegCliente);
                    }
                }
                lreader.Close();
            }

            return _regclientes;


        }



        public List<RegAlmacen> mCargarAlmacenesComercialv2()
        {



            SqlConnection lconexion = new SqlConnection();

            List<RegAlmacen> _regalmacenes = new List<RegAlmacen>();
            lconexion = miconexion.mAbrirConexionComercial(false);
            if (lconexion != null)
            {

                SqlCommand lsql = new SqlCommand("select cidalmacen, ccodigoalmacen,cnombrealmacen from admalmacenes ", lconexion);
                SqlDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _regalmacenes.Clear();

                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegAlmacen lRegAlmacen = new RegAlmacen();
                        lRegAlmacen.Id = long.Parse(lreader[0].ToString());
                        lRegAlmacen.Codigo = lreader[1].ToString();
                        lRegAlmacen.Nombre = lreader[2].ToString();
                        _regalmacenes.Add(lRegAlmacen);
                    }
                }
                lreader.Close();
            }

            return _regalmacenes;


        }
        public RegProveedor mBuscarCliente(string aCliente, int aTipo, int aTipoCliente)
        {
            OleDbConnection lconexion = new OleDbConnection();
            RegProveedor lReg = new RegProveedor();
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
                    lReg.DiasCredito = int.Parse(lreader[4].ToString());
                }
                lreader.Close();
            }
            return lReg;


        }

        public List<RegEmpresas> mCargarEmpresasAccess(out string amensaje)
        {

            OleDbConnection lconexion = new OleDbConnection();

            lconexion = miconexion.mAbrirConexionAccess(out amensaje);

            List<RegEmpresas> _RegEmpresas = new List<RegEmpresas>();
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
                            RegEmpresas lRegEmpresas = new RegEmpresas();
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

            lconexion = miconexion.mAbrirConexionAccess(out amensaje);

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
                            lRePuntodeVenta.cNombre = lreader[0].ToString();
                            _RegPUntosVenta.Add(lRePuntodeVenta);
                        }
                    }
                    lreader.Close();

                }
                catch (Exception eeeee)
                {
                    amensaje = eeeee.Message;
                }

            }



            return _RegPUntosVenta;




        }


        public List<RegEmpresa> mCargarEmpresas(out string amensaje)
        {

            OleDbConnection lconexion = new OleDbConnection();

            lconexion = miconexion.mAbrirRutaGlobal(out amensaje);

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
            string Cadenaconexion = "data source =" + aServidor + ";initial catalog =" + aBd + ";user id = " + ausu + "; password = " + apwd + ";";

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


        public string mBuscarDocto(string aFolio, int aTipo, Boolean aRevisar)
        {
            OleDbCommand lcmd = new OleDbCommand();
            OleDbDataReader lreader;
            string lRespuesta = "";
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

                    lcmd.CommandText += ", case " +
                                        " when  v.condicion = '' then '0' " +
                                        " when  v.condicion = 'Contado' then '0'" +
                                        " when  isnull(v.condicion,0) = '0' then '0'" +
                                        " else left(v.condicion, isnull(charindex(' DIAS CREDITO',v.condicion,1),0)) " +
                                        " end as condpago, v.agente ";

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
                lRespuesta = mLlenarDocto(lreader, aTipo, aFolio, "Mercado");
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

        private string mProcesaItem(ref int aInicio, string sLine)
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
            List<RegDocto> misdoctos = new List<RegDocto>();
            _RegDoctos.Clear();
            string lrutacarpeta = @GetSettingValueFromAppConfigForDLL("RutaCarpeta");
            //lrutacarpeta = @lrutacarpeta;
            foreach (string txtName in Directory.GetFiles(lrutacarpeta, "*.txt"))
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
                                string x = mProcesaItem(ref linicio, sLine);
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

        public string mBuscarDoctoAccess(Boolean aRevisar)
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
                lreader = null;
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

        private Boolean mBuscarGeneradoADM(string aFolio, int aTipo)
        {
            OleDbConnection lconexion = new OleDbConnection();
            string amensaje = "";
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

            return lrespuesta;
        }


        public RegDocto mBuscarDoctoComercial(string aFolio, string aSerie, string Concepto)
        {

            _RegDoctos.Clear();
            if (miconexion._conexion1 != null)
            {
                if (miconexion._conexion1.State == ConnectionState.Closed)
                    miconexion.mAbrirConexionComercial(false);
            }
            else
                miconexion.mAbrirConexionComercial(false);
            RegDocto ldocto = new RegDocto();


            SqlCommand lsql = new SqlCommand();
            SqlDataReader lreader;

            // miconexion.mAbrirConexionDestino();

            lsql.CommandText = "select d.ciddocumento,d.cfecha,cl.crazonsocial,cl.ccodigocliente, m.cidmovimiento, pl.ccodigoproducto, m.cunidades," +
" pl.cnombreproducto , m.cidalmacen, a.cnombrealmacen, a.ccodigoalmacen," +
" m.cprecio,m.cneto, m.cimpuesto1, m.ctotal, m.cidmovimiento " +
" from admDocumentos d " +
" join admConceptos c on c.cidconceptodocumento = d.cidconceptodocumento " +
" join admClientes cl on cl.cidclienteproveedor = d.cidclienteproveedor " +
" join admmovimientos m on d.ciddocumento = m.ciddocumento " +
" join admProductos pl on m.cidproducto = pl.CIDPRODUCTO" +
" join admAlmacenes a on a.CIDalmacen = m.CIDalmacen " +
                " where cFolio = '" + aFolio + "' and d.cSerieDocumento = '" + aSerie + "'" +
                " and c.ccodigoconcepto = '" + Concepto + "'";
            lsql.Connection = miconexion._conexion1;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            string lregresa = "";
            string x = "";
            long ldoc = 0;
            if (lreader.HasRows)
            {
                while (lreader.Read())
                {
                    try
                    {
                        if (ldoc == 0)
                        {
                            ldocto.cIdDocto = long.Parse(lreader[0].ToString());
                            ldocto.cFecha = DateTime.Parse(lreader[1].ToString());
                            ldocto.cRazonSocial = lreader[2].ToString();
                            ldocto.cCodigoConcepto = Concepto;
                            ldocto.cCodigoCliente = lreader[3].ToString();
                            ldocto.cFolio = long.Parse(aFolio);
                            ldoc = 1;

                        }
                        RegMovto m = new RegMovto();
                        m.cCodigoAlmacen = lreader["ccodigoalmacen"].ToString();

                        m.cCodigoProducto = lreader["ccodigoproducto"].ToString();  //nombre del 
                        m.cNombreProducto = lreader["cnombreproducto"].ToString();


                        m.cNombreAlmacen = lreader["cnombrealmacen"].ToString();

                        m.cUnidades = decimal.Parse(lreader[6].ToString());

                        m.cneto = decimal.Parse(lreader["cneto"].ToString());
                        m.cImpuesto = decimal.Parse(lreader["cimpuesto1"].ToString());
                        //m.cSubtotal = decimal.Parse(lreader["csubtotal"].ToString());
                        m.cTotal = decimal.Parse(lreader["ctotal"].ToString());

                        m.cPrecio = decimal.Parse(lreader["cprecio"].ToString());
                        m.cIdMovto = long.Parse(lreader["cidmovimiento"].ToString());
                        ldocto._RegMovtos.Add(m);
                    }
                    catch (Exception ee)
                    {
                        //                    lreader.Close();
                    }
                }


            }
            lreader.Close();
            miconexion.mCerrarConexionOrigenComercial();
            //_RegDoctos.Add(ldocto);
            return ldocto;


        }


        public RegDocto mBuscarDoctoComercialProduccion(string aFolio, string Concepto, int porcentaje)
        {

            _RegDoctos.Clear();
            if (miconexion._conexion1 != null)
            {
                if (miconexion._conexion1.State == ConnectionState.Closed)
                    miconexion.mAbrirConexionComercial(false);
            }
            else
                miconexion.mAbrirConexionComercial(false);
            RegDocto ldocto = new RegDocto();


            SqlCommand lsql = new SqlCommand();
            SqlDataReader lreader;

            // miconexion.mAbrirConexionDestino();

            lsql.CommandText = "select d.ciddocumento,d.cfecha,cl.crazonsocial, m.cidmovimiento, pl.ccodigoproducto ccodigopaquete, cp.CCANTIDADPRODUCTO " +
", paq.cnombreproducto cnombrepaquete, m.cidalmacen, a.cnombrealmacen, a.ccodigoalmacen,p2.ccodigoproducto, p2.cnombreproducto " +
",m.cunidades " +
", ROW_NUMBER() over(partition by pl.CCODIGOPRODUCTO order by pl.CCODIGOPRODUCTO) orden" +
", paq.ccodigoproducto ccodigopaquete, m.cprecio " +
" from admDocumentos d " +
" join admConceptos c on c.cidconceptodocumento = d.cidconceptodocumento " +
" join admClientes cl on cl.cidclienteproveedor = d.cidclienteproveedor " +
" join admmovimientos m on d.ciddocumento = m.ciddocumento " +
//" join admProductos pl on m.cidproducto = pl.CIDPRODUCTO and pl.CCONTROLEXISTENCIA = 16" +
" join admProductos pl on m.cidproducto = pl.CIDPRODUCTO" +
" join admProductos paq on paq.CCODIGOPRODUCTO = substring(pl.ccodigoproducto,2,100) " +
" join admComponentesPaquete cp on cp.cidpaquete = paq.CIDPRODUCTO " +
" join admProductos p2 on cp.CIDPRODUCTO = p2.CIDPRODUCTO " +
" join admAlmacenes a on a.CIDalmacen = m.CIDalmacen " +
                "where cFolio = '" + aFolio + "' and c.ccodigoconcepto = '" + Concepto + "'" +
                " and d.ctextoextra3 = ''";
            lsql.Connection = miconexion._conexion1;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            string lregresa = "";
            string x = "";
            long ldoc = 0;
            if (lreader.HasRows)
            {
                while (lreader.Read())
                {
                    try
                    {
                        if (ldoc == 0)
                        {
                            ldocto.cIdDocto = long.Parse(lreader[0].ToString());
                            ldocto.cFecha = DateTime.Parse(lreader[1].ToString());
                            ldocto.cRazonSocial = lreader[2].ToString();
                            ldocto.cCodigoConcepto = "19";
                            ldocto.cCodigoCliente = "PROV1";
                            ldocto.cFolio = -1;
                            ldoc = 1;

                        }
                        RegMovto m = new RegMovto();
                        m.cCodigoAlmacen = lreader["ccodigoalmacen"].ToString();

                        m.cCodigoProducto = lreader["ccodigopaquete"].ToString();  //nombre del 
                        m.cNombreProducto = lreader["cnombrepaquete"].ToString();

                        m.ctextoextra1 = lreader["ccodigoproducto"].ToString();
                        m.ctextoextra2 = lreader["cnombreproducto"].ToString();


                        m.cNombreAlmacen = lreader["cnombrealmacen"].ToString();

                        m.cUnidades = decimal.Parse(lreader[5].ToString());

                        m.cMargenUtilidad = decimal.Parse(lreader["cunidades"].ToString());
                        m.cneto = decimal.Parse(lreader["orden"].ToString());
                        m.cIdMovtoOrigen = int.Parse(lreader["cidmovimiento"].ToString());
                        if (porcentaje > 0)
                            m.cUnidades += decimal.Parse(lreader[5].ToString()) * porcentaje / 100;

                        m.ctextoextra3 = lreader["ccodigopaquete"].ToString();
                        m.cimporteextra2 = decimal.Parse(lreader["ccantidadproducto"].ToString());
                        m.cPrecio = decimal.Parse(lreader["cprecio"].ToString());

                        ldocto._RegMovtos.Add(m);
                    }
                    catch (Exception ee)
                    {
                        //                    lreader.Close();
                    }
                }


            }
            lreader.Close();
            //miconexion.mCerrarConexionOrigenComercial();
            _RegDoctos.Add(ldocto);
            return ldocto;


        }

        protected Boolean mBuscarADM(string aFolio, int aTipo)
        {
            bool lrespuesta = false;
            string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();


            miconexion.mAbrirConexionDestino();
            string lcadena = "select cfolio from mgw10008 m8 join mgw10006 m6 on m6.cidconce01 = m8.cidconce01 where m8.cfolio = " + aFolio + " and m6.ccodigoc01 = '" + lCodigoConcepto + "'";
            OleDbCommand lsql = new OleDbCommand(lcadena, miconexion._conexion);
            OleDbDataReader lreader;
            //long lIdDocumento = 0;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lrespuesta = true;
            }
            lreader.Close();
            miconexion.mCerrarConexionDestino();
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
            string lcadena = "";

            //List<string> lvar = new List<string>();

            lvar.Clear();
            int lcuantos = _RegDoctos.Count;
            int lindice = 1;

            if (_RegDoctos.Count == 0)
            {
                lvar.Add("No existe documentos con los filtros seleccionados");
                return lvar;
            }


            foreach (RegDocto _reg in _RegDoctos)
            {
                _RegDoctoOrigen = null;
                _RegDoctoOrigen = new RegDocto();
                _RegDoctoOrigen = _reg;
                string lCodigoConcepto;
                //if (opcion != 5)
                //    lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
                //else
                lCodigoConcepto = _reg.cCodigoConcepto;

                //lrespuesta = _RegDoctoOrigen.sMensaje;
                //if (_RegDoctoOrigen.sMensaje == string.Empty)
                //{
                lrespuesta = mGrabarAdm(_reg.cFolio.ToString(), _RegDoctoOrigen.cFolio, opcion, tipo);

                //}
                //mActualizarBarra((double)lindice / lcuantos);
                //lporcentaje = 0.0D;
                //lporcentaje = (double)lindice / lcuantos;
                //Notificar();
                Notificar((double)(lindice * 100) / lcuantos);



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
                    lrutaorigen += "\\" + _RegDoctoOrigen.cNombreArchivo;
                    string lrutadestino = GetSettingValueFromAppConfigForDLL("RutaCarpetaBackup");
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

        private void mRegresarPrincipales(string lCodigoConcepto, ref long lidconce, ref long tipocfd, ref string cserie, ref int cnaturaleza)
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
            long lret, lidconce = 0, tipocfd = 0;
            string cserie = "";

            int naturaleza = 0;
            mRegresarPrincipales(lCodigoConcepto, ref lidconce, ref tipocfd, ref cserie, ref naturaleza);
            string lresp = mValidarExisteDoc(lidconce, cserie, aFolio);
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


            try
            {
                lret = fGuardaDocumento();
            }
            catch (Exception eee)
            {
                string wwww = eee.Message;
            }
            if (lret != 0)
            {
                //fError(lret, serror, 255);
                _controlfp(0x9001F, 0xFFFFF);
                //miconexion.mCerrarConexionOrigen(1);
                return lret.ToString() + " Documento ya Existe";


            }
            return "";



        }

        private string mGrabarCliente()
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
            long lret = 0;
            mLeerDireccion();

            RegDireccion lRegDireccion = new RegDireccion();
            // la direccion del cliente pasarla a la direccion de la factura
            lRegDireccion = _RegDoctoOrigen._RegDireccion;
            if (lRegDireccion.cNombreCalle != null)
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
            string cad2 = "";
            if (cserie.Trim() != "")
                cad2 = "select max(cfolio)+1  from mgw10008 m8 join mgw10006 m6 on m8.cidconce01 = m6.cidconce01 and m6.ccodigoc01 = '" + aConcepto + "'";
            else
                cad2 = "select max(cfolio)+1 from mgw10008 m8 join mgw10007 m7 on m8.ciddocum02 = m7.ciddocum01 and m7.ciddocum01 = " + aIdDocumentoModelo;

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


        public string mGrabarFresko(long afolionuevo, int opcion, bool incluyetimbrado, int tipo)
        {
            //miconexion.mAbrirConexionDestino(1);
            string lCodigoConcepto;
            //lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            lCodigoConcepto = _RegDoctoOrigen.cCodigoConcepto;

            string lresp1 = mGrabarEncabezado(afolionuevo, lCodigoConcepto, "0");
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
            lresp = mGrabarMovimientos(lIdDocumento, opcion, tipo);

            if (_RegDoctoOrigen.cCodigoConcepto == "1")
            {
                mGrabarRemision(lIdDocumento);
                lCodigoConcepto = "3";
            }

            string lrespuestas = mGrabarExtrasObservaciones(lIdDocumento);



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


        public string mGrabarTraslado(long afolionuevo, int opcion, bool incluyetimbrado, int tipo)
        {

            string lrespuestas = "";
            //miconexion.mAbrirConexionDestino(1);
            string lCodigoConcepto;
            //lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            lCodigoConcepto = _RegDoctoOrigen.cCodigoConcepto;

            string lresp1 = mGrabarEncabezado(afolionuevo, lCodigoConcepto, "0");
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

            //mActualizaDocumentoTraslado(lIdDocumento);

            //string lresp = mGrabarDireccion(lIdDocumento);
            DateTime lFechaDocto;
            lFechaDocto = _RegDoctoOrigen.cFecha;

            string lfechadocto = "";


            lfechadocto = String.Format("{0:MM/dd/yyyy}", lFechaDocto); ;  // "8 08 008 2008"   year



            string lresp = mGrabarMovimientosTraslado(lIdDocumento, opcion, tipo, lfechadocto);


            mActualizaDocumento(lIdDocumento, 0, 0);

            //string lrespuestas = mGrabarExtrasObservaciones(lIdDocumento);



            int lret = fAfectaDocto_Param(lCodigoConcepto, cserie, afolionuevo, true);


            string lcadenaA = "update mgw10008 set cpendiente= ctotal, cunidade01=ctotalun01 where ciddocum01 = " + lIdDocumento.ToString();
            OleDbCommand lsqlA = new OleDbCommand(lcadenaA, miconexion._conexion);
            int xxx = 0;
            xxx = lsqlA.ExecuteNonQuery();

            //mImprimir(lIdDocumento);
            long lresp10 = 0;

            /*int lNumeroMov = 100;
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
            */


            /*  try
              {
                  _controlfp(0x9001F, 0xFFFFF);
              }
              catch (Exception eee)
              {
                  lrespuestas = eee.Message;
              }*/
            // barra.Asignar(100);
            return lrespuestas;
        }


        public string mGrabarAdmNew(long afolionuevo, int opcion, bool incluyetimbrado, int tipo)
        {
            //miconexion.mAbrirConexionDestino(1);
            string lCodigoConcepto;
            //lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            lCodigoConcepto = _RegDoctoOrigen.cCodigoConcepto;

            string lresp1 = mGrabarEncabezado(afolionuevo, lCodigoConcepto, "0");
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
            lresp = mGrabarMovimientosFresko(lIdDocumento, opcion, tipo);

            if (_RegDoctoOrigen.cCodigoConcepto == "1")
            {
                mGrabarRemision(lIdDocumento);
                lCodigoConcepto = "3";
            }

            string lrespuestas = mGrabarExtrasObservaciones(lIdDocumento);



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


        private string mGrabarTraspaso(long aIdDocumentoSalida, long aIdDocumentoeEntrada, double aFoliot)
        {
            long x;
            //    miconexion.mAbrirConexionDestino();
            //string cad = "select cidmovim01, cunidades from mgw10010 where ciddocum01 = " + aIdDocumento + " order by cidmovim01 ";
            //cad = "select cidmovimiento cidmovim01, cunidades from admmovimientos where ciddocumento = " + aIdDocumento + " order by cidmovimiento ";
            string cad = "select cidmovimiento cidmovim01, cunidades from admmovimientos where ciddocumento in ( " + aIdDocumentoSalida + "," + aIdDocumentoeEntrada + ") ORDER BY CNUMEROMOVIMIENTO ASC, CIDDOCUMENTODE DESC ";

            /*
            SqlCommand lsql = new SqlCommand();
            lsql.CommandText = cad;
            lsql.Connection = miconexion._conexion1;
            int lret = lsql.ExecuteNonQuery();
            */



            SqlCommand lsql = new SqlCommand(cad, miconexion._conexion1);
            SqlDataReader lreader;

            //lreader = lsql.ExecuteReader(CommandBehavior.CloseConnection);


            DataTable dt = new DataTable();
            dt.Load(lsql.ExecuteReader());


            x = 1;
            long idmov1 = 0;
            long idmov2 = 0;
            decimal lunidades = 0;
            SqlCommand lsql4 = new SqlCommand();
            string lcadena2 = "";

            foreach (DataRow row in dt.Rows)
            {
                if (x % 2 == 0) // movto 2
                {

                    idmov2 = long.Parse(row[0].ToString());
                    lunidades += decimal.Parse(row[1].ToString());
                    lcadena2 = "update mgw10010 set ciddocum02 =34, ciddocum01 = 0, cnumerom01=0,cafectae01 = 1, cafectad01 = 0, cmovtooc01 = 1, cidmovto01 = " + idmov1 + " where cidmovim01 = " + idmov2;
                    lcadena2 = "update admmovimientos set cporcentajeimpuesto1=0, ciddocumentode =34, ciddocumento = 0, cnumeromovimiento=0,CAFECTAEXISTENCIA = 1, CAFECTADOINVENTARIO = 0, cmovtooculto = 1, CIDMOVTOOWNER = " + idmov1 + " where cidmovimiento = " + idmov2;

                    lcadena2 = "update admmovimientos set ciddocumentode =34, ciddocumento = 0,ccostocapturado = 0, cneto=0,ctotal=0, cnumeromovimiento=0, cmovtooculto = 1, CIDMOVTOOWNER = " + idmov1 + " where cidmovimiento = " + idmov2;


                    lsql4.CommandText = lcadena2;
                    lsql4.Connection = miconexion._conexion1;
                    lsql4.ExecuteNonQuery();

                }
                else
                {

                    idmov1 = long.Parse(row[0].ToString());

                    lcadena2 = "update mgw10010 set ciddocum02 = 34, cafectae01 = 2, cafectad01 = 0, cmovtooc01 = 0  where cidmovim01 = " + idmov1;
                    lcadena2 = "update admmovimientos set cporcentajeimpuesto1=0, ciddocumentode = 34, cafectaexistencia = 2, cafectadoinventario = 0, cmovtooculto = 0  where cidmovimiento = " + idmov1;
                    lcadena2 = "update admmovimientos set ciddocumentode = 34, cmovtooculto = 0  where cidmovimiento = " + idmov1;


                    lsql4.CommandText = lcadena2;
                    lsql4.Connection = miconexion._conexion1;
                    lsql4.ExecuteNonQuery();

                }
                x++;

            }


            //lcadena2 = "update mgw10008 set ciddocum02 = 3, cidconce01= 3, ctotalun01 = " + lunidades + ", cunidade01 = " + lunidades + " where ciddocum01 = " + aIdDocumento;

            //lcadena2 = "update admdocumentos set cidclienteproveedor = 0, crazonsocial = '', cusacliente = 0, cdestinatario = '', cbanobservaciones = 0, ciddocumentode = 34, cidconceptodocumento= 36, ctotalunidades = " + lunidades + ", cunidadespendientes = " + lunidades + " where ciddocumento = " + aIdDocumento;

            lcadena2 = "update admdocumentos set ciddocumentode = 34, cidconceptodocumento= 36, ctotalunidades = " + lunidades + ", cunidadespendientes = " + lunidades + " where ciddocumento = " + aIdDocumentoSalida;


            //lcadena2 = "update mgw10008 set ciddocum02 = 3, cidconce01= 3 where ciddocum01 = " + aIdDocumento;
            // cambiar el concepto al documento
            lsql4.CommandText = lcadena2;
            lsql4.Connection = miconexion._conexion1;
            lsql4.ExecuteNonQuery();
            //  miconexion.mCerrarConexionDestino();

            long lret = fBuscarDocumentoComercial("34", "", aFoliot.ToString());
            if (lret == 0)
            {
                fBorraDocumentoComercial();
            }



            return "";
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
            long idmov1 = 0;
            long idmov2 = 0;
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

            lcadena2 = "update mgw10008 set ciddocum02 = 3, cidconce01= 3, ctotalun01 = " + lunidades + ", cunidade01 = " + lunidades + " where ciddocum01 = " + aIdDocumento;
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

        private string mGrabarExtras(long lIdDocumento, int opcion, double afolionuevo)
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;

            string lresp = "";
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
                OleDbCommand lsqlA = new OleDbCommand(lcadenaA, miconexion._conexion);
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
            long lIdDocumento = mBuscarIdDocumento(lCodigoConcepto, 0, cserie, long.Parse(afolionuevo.ToString().Trim()));


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

            lresp = mGrabarMovimientos(lIdDocumento, opcion, 0);


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
                mActualizaDocumento(lIdDocumento, opcion, afolionuevo);

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
                    decimal limpuestos = decimal.Parse(_RegDoctoOrigen.cImpuestos.ToString());
                    limpuestos = decimal.Round(limpuestos, 4);
                    string lcadena1 = "update mgw10008 set cneto = " + _RegDoctoOrigen.cNeto.ToString() + ", cimpuesto1 = " + limpuestos + ",ctotal = " + ltotal.ToString() + ",cpendiente = " + ltotal.ToString() + ",ctipocam01 = " + _RegDoctoOrigen.cTipoCambio + " where ciddocum01 = " + lIdDocumento;

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

        private string mRegresaProductoporDescription(string descripcion)
        {
            OleDbCommand lsql = new OleDbCommand();
            lsql.Connection = miconexion._conexion;
            lsql.CommandText = "select ccodigop01 from mgw10005 where cNombreP01 = '" + descripcion + "'";
            OleDbDataReader ldr = lsql.ExecuteReader();
            string lcodigo = "";
            if (ldr.HasRows == false)
            {
                // insertar nuevo producto
                ldr.Close();
                lsql.CommandText = "select cprecio1+1 from mgw10005 where cidprodu01 = 0";
                ldr = lsql.ExecuteReader();
                if (ldr.HasRows == true)
                {
                    ldr.Read();
                    lcodigo = "P" + ldr[0].ToString().PadLeft(29, '0');
                    lsql.CommandText = "update mgw10005 set cprecio1 = cprecio1+1 where cidprodu01 = 0";
                    ldr.Close();
                    lsql.ExecuteNonQuery();
                }
                else
                    ldr.Close();

            }
            else
            {
                ldr.Read();
                lcodigo = ldr[0].ToString();
                ldr.Close();
            }
            return lcodigo;
        }



        private string mGrabarMovimientosTraslado(long lIdDocumento, int opcion, int tipo, string lfecha)
        {
            long lret = 0;
            long lNumeroMov = 100;
            foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
            {
                lret = fInsertarMovimiento();
                lret = fSetDatoMovimiento("cIdDocumento", lIdDocumento.ToString());
                lret = fSetDatoMovimiento("cNumeroMovimiento", lNumeroMov.ToString());

                // buscar el producto por description
                // x.cCodigoProducto = mRegresaProductoporDescription(x.cNombreProducto);
                //x.cNombreProducto

                lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
                //lret = fSetDatoMovimiento("cObserva01", x.ctextoextra3);
                if (lret != 0 || x.cCodigoProducto == "")
                {
                    // buscar la desc

                    fInsertaProducto();
                    lret = fSetDatoProducto("CCODIGOPRODUCTO", x.cCodigoProducto);
                    lret = fSetDatoProducto("CNOMBREPRODUCTO", x.cNombreProducto);
                    lret = fSetDatoProducto("CTIPOPRODUCTO", "3");
                    lret = fSetDatoProducto("CMETODOCOSTEO", "1");
                    lret = fSetDatoProducto("CCONTROLEXISTENCIA", "1");
                    lret = fSetDatoProducto("CCLAVESAT", x._RegProducto.CodigoSAT);
                    x.ctextoextra1 = "";
                    //lret = fSetDatoDocumento("COBSERVACIONES", x.cObservaciones);
                    //lret = fSetDatoProducto("CIMPUESTO1", x.cPorcent01.ToString());

                    lret = fSetDatoProducto("CIMPUESTO1", "0");
                    lret = fSetDatoProducto("CIMPUESTO2", x.cImpuesto.ToString());
                    lret = fSetDatoProducto("CFECHAALTA", lfecha);
                    lret = fSetDatoProducto("CSTATUSPRODUCTO", "1");






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
                            cmdunidad.CommandText = "insert into mgw10026 values (" +
                                lidunidad + ",'" + x.cUnidad.ToUpper() + "','" + x.cUnidad.ToUpper() + "','" + x.cUnidad.ToUpper() + "','" + x._RegProducto.CodigoMedidaPesoSAT + "','')";
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
                    lret = fGuardaProducto();


                    StringBuilder aValor1 = new StringBuilder(12);
                    int lret2 = fLeeDatoProducto("cIdProducto", aValor1, 12);
                    int lidproducto = int.Parse(aValor1.ToString());


                    OleDbCommand lsql = new OleDbCommand();
                    if (x.traslado.materialpeligroso != "")
                    {
                        string lcadenax = "insert into mgw10046 values (505,2," + lidproducto.ToString() + ",2,'" + x.traslado.materialpeligroso + "')";
                        lsql.CommandText = lcadenax;
                        lsql.Connection = miconexion._conexion;
                        lsql.ExecuteNonQuery();
                    }

                    if (x.traslado.cvematerialpeligroso != "")
                    {
                        string lcadenax = "insert into mgw10046 values (505,2," + lidproducto.ToString() + ",3,'" + x.traslado.cvematerialpeligroso + "')";
                        lsql.CommandText = lcadenax;
                        lsql.ExecuteNonQuery();
                        lsql.Connection = miconexion._conexion;
                    }


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



                if (tipo == 1)
                {
                    lret = fGuardaMovimiento();
                    if (lret != 0 && lret != 135021)
                    {
                        MessageBox.Show(lret.ToString());
                        StringBuilder sMensaje1 = new StringBuilder(255);
                        if (lret != 0)
                        {
                            int lret1 = (int)lret;
                            fError(lret1, sMensaje1, 255);
                            MessageBox.Show(sMensaje1.ToString());
                            //fProcesaError(doc, doc.cIdDocto, "El documento con cliente " + doc.cCodigoCliente.Trim() + " y folio " + doc.cFolio.ToString() + " presenta el sig. problema " + sMensaje1.ToString(), ref lret2);
                            //fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                            return "0";
                        }
                    }

                    StringBuilder aValor1 = new StringBuilder(12);

                    aValor1.Length = 0;
                    int lret2 = fLeeDatoMovimiento("CIDMOVIMIENTO", aValor1, 12);
                    int lidmovimiento1 = int.Parse(aValor1.ToString());

                    x.cIdMovto = lidmovimiento1;

                    // guardar addenda 

                    OleDbConnection lconexion = new OleDbConnection();
                    //OleDbDataReader lreader;
                    string lcadena = "";

                    lconexion = miconexion._conexion;
                    if (lconexion != null)
                    {

                        // //lcadena = "select cidvalor01,cvalorcl01 from mgw10020 where ccodigov01 = '" + codigo + "' and cidclasi01 =" + anumClasif.ToString();
                        //lcadena = "UPDATE mgw10046 set valor = '" + x.traslado. + "' where idaddenda = 336 and tipocat = 2 and idcat = " + lprod.Id.ToString() + " and numCampo = 1";
                        OleDbCommand lsql = new OleDbCommand(lcadena, lconexion);
                        //int lcuantos = lsql.ExecuteNonQuery();

                        lcadena = "insert into mgw10046 values (505,5," + lidmovimiento1.ToString() + ",5,'" + x.traslado.PesoEnKg + "')";
                        lsql.CommandText = lcadena;
                        lsql.ExecuteNonQuery();
                        //}
                        /*lcadena = "UPDATE mgw10046 set valor = '" + aValor2 + "' where idaddenda = 336 and tipocat = 2 and idcat = " + lprod.Id.ToString() + " and numCampo = 2";
                        lsql.CommandText = lcadena;
                        lcuantos = lsql.ExecuteNonQuery();
                        */
                        //if (lcuantos == 0)
                        // {
                        lcadena = "insert into mgw10046 values (505,5," + lidmovimiento1.ToString() + ",6,'" + x.traslado.ValorMercancia + "')";
                        lsql.CommandText = lcadena;
                        lsql.ExecuteNonQuery();
                        //}
                        /*lcadena = "UPDATE mgw10046 set valor = '" + aValor3 + "' where idaddenda = 336 and tipocat = 2 and idcat = " + lprod.Id.ToString() + " and numCampo = 3";
                        lsql.CommandText = lcadena;
                        lcuantos = lsql.ExecuteNonQuery();

                        if (lcuantos == 0)
                        {*/
                        lcadena = "insert into mgw10046 values (505,5," + lidmovimiento1.ToString() + ",7,'" + x.traslado.Moneda + "')";
                        lsql.CommandText = lcadena;
                        lsql.ExecuteNonQuery();
                        /*}
                        lcadena = "UPDATE mgw10046 set valor = '" + aValor4 + "' where idaddenda = 336 and tipocat = 2 and idcat = " + lprod.Id.ToString() + " and numCampo = 4";
                        lsql.CommandText = lcadena;
                        lcuantos = lsql.ExecuteNonQuery();

                        if (lcuantos == 0)
                        {*/
                        lcadena = "insert into mgw10046 values (505,5," + lidmovimiento1.ToString() + ",8,'')";
                        lsql.CommandText = lcadena;
                        lsql.ExecuteNonQuery();
                        //}

                        lcadena = "insert into mgw10046 values (505,5," + lidmovimiento1.ToString() + ",9,'')";
                        lsql.CommandText = lcadena;
                        lsql.ExecuteNonQuery();
                        lcadena = "insert into mgw10046 values (505,5," + lidmovimiento1.ToString() + ",10,'" + x.traslado.Pedimento + "')";
                        lsql.CommandText = lcadena;
                        lsql.ExecuteNonQuery();
                        lcadena = "insert into mgw10046 values (505,5," + lidmovimiento1.ToString() + ",11,'')";
                        lsql.CommandText = lcadena;
                        lsql.ExecuteNonQuery();
                        lcadena = "insert into mgw10046 values (505,5," + lidmovimiento1.ToString() + ",12,'0.0')";
                        lsql.CommandText = lcadena;
                        lsql.ExecuteNonQuery();
                    }



                    //miconexion.mCerrarConexionDestino();
                }

                //doc._RegMovtos[indicemov].cIdMovto = lidmovimiento;
                //_RegDoctos[indicedoc]._RegMovtos[indicemov++].cIdMovto = lidmovimiento;




                lNumeroMov += 100;

            }
            /*
            OleDbConnection lconexiondoc = new OleDbConnection();
            //OleDbDataReader lreader;
            
            lconexiondoc = miconexion._conexion;
            OleDbCommand lsqldoc = new OleDbCommand();

            string lcadenadoc = "insert into mgw10046 values (505,3," + lIdDocumento.ToString() + ",1,'No')";
            lsqldoc.Connection = lconexiondoc;
            lsqldoc.CommandText = lcadenadoc;
            lsqldoc .ExecuteNonQuery();

            lcadenadoc = "insert into mgw10046 values (505,3," + lIdDocumento.ToString() + ",2,'')";
            lsqldoc.Connection = lconexiondoc;
            lsqldoc.CommandText = lcadenadoc;
            lsqldoc.ExecuteNonQuery();


            lcadenadoc = "insert into mgw10046 values (505,3," + lIdDocumento.ToString() + ",2,'')";
            lsqldoc.Connection = lconexiondoc;
            lsqldoc.CommandText = lcadenadoc;
            lsqldoc.ExecuteNonQuery();

            lcadenadoc = "insert into mgw10046 values (505,3," + lIdDocumento.ToString() + ",2,'')";
            lsqldoc.Connection = lconexiondoc;
            lsqldoc.CommandText = lcadenadoc;
            lsqldoc.ExecuteNonQuery();

            lcadenadoc = "insert into mgw10046 values (505,3," + lIdDocumento.ToString() + ",2,'')";
            lsqldoc.Connection = lconexiondoc;
            lsqldoc.CommandText = lcadenadoc;
            lsqldoc.ExecuteNonQuery();

            lcadenadoc = "insert into mgw10046 values (505,3," + lIdDocumento.ToString() + ",2,'')";
            lsqldoc.Connection = lconexiondoc;
            lsqldoc.CommandText = lcadenadoc;
            lsqldoc.ExecuteNonQuery();

            lcadenadoc = "insert into mgw10046 values (505,3," + lIdDocumento.ToString() + ",2,'')";
            lsqldoc.Connection = lconexiondoc;
            lsqldoc.CommandText = lcadenadoc;
            lsqldoc.ExecuteNonQuery();

            lcadenadoc = "insert into mgw10046 values (505,3," + lIdDocumento.ToString() + ",2,'')";
            lsqldoc.Connection = lconexiondoc;
            lsqldoc.CommandText = lcadenadoc;
            lsqldoc.ExecuteNonQuery();

            lcadenadoc = "insert into mgw10046 values (505,3," + lIdDocumento.ToString() + ",3,'Autotransporte')";
            lsqldoc.Connection = lconexiondoc;
            lsqldoc.CommandText = lcadenadoc;
            lsqldoc.ExecuteNonQuery();

            lcadenadoc = "insert into mgw10046 values (505,3," + lIdDocumento.ToString() + ",4,'')";
            lsqldoc.Connection = lconexiondoc;
            lsqldoc.CommandText = lcadenadoc;
            lsqldoc.ExecuteNonQuery();

            lcadenadoc = "insert into mgw10046 values (505,3," + lIdDocumento.ToString() + ",6,'CL001')";
            lsqldoc.Connection = lconexiondoc;
            lsqldoc.CommandText = lcadenadoc;
            lsqldoc.ExecuteNonQuery();

            lcadenadoc = "insert into mgw10046 values (505,3," + lIdDocumento.ToString() + ",8,'')";
            lsqldoc.Connection = lconexiondoc;
            lsqldoc.CommandText = lcadenadoc;
            lsqldoc.ExecuteNonQuery();
            */
            return "";
        }


        private string mGrabarMovimientosFresko(long lIdDocumento, int opcion, int tipo)
        {
            long lret = 0;
            long lNumeroMov = 100;
            foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
            {
                lret = fInsertarMovimiento();
                lret = fSetDatoMovimiento("cIdDocumento", lIdDocumento.ToString());
                lret = fSetDatoMovimiento("cNumeroMovimiento", lNumeroMov.ToString());

                // buscar el producto por description
                x.cCodigoProducto = mRegresaProductoporDescription(x.cNombreProducto);
                //x.cNombreProducto

                lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
                //lret = fSetDatoMovimiento("cObserva01", x.ctextoextra3);
                if (lret != 0 || x.cCodigoProducto == "")
                {
                    // buscar la desc

                    fInsertaProducto();
                    lret = fSetDatoProducto("CCODIGOPRODUCTO", x.cCodigoProducto);
                    lret = fSetDatoProducto("CNOMBREPRODUCTO", x.cNombreProducto);
                    lret = fSetDatoProducto("CTIPOPRODUCTO", "1");
                    lret = fSetDatoProducto("CMETODOCOSTEO", "1");
                    lret = fSetDatoProducto("CCONTROLEXISTENCIA", "1");
                    x.ctextoextra1 = "";
                    //lret = fSetDatoDocumento("COBSERVACIONES", x.cObservaciones);
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
                            cmdunidad.CommandText = "insert into mgw10026 values (" + lidunidad + ",'" + x.cUnidad.ToUpper() + "','" + x.cUnidad.ToUpper() + "','" + x.cUnidad.ToUpper() + "','','')";
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
                    lret = fGuardaProducto();
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
            string lcadena = "";
            decimal lresta = aExistencia - x.cUnidades;
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
            decimal lidclien = 0;
            //OleDbParameter lparametroIdDocumento = new OleDbParameter("@p1", _RegDoctoOrigen.cIdDocto);
            lsql.CommandText = "select m30.cidalmacen, m30.centrada01-m30.csalidas01 as ini, m30.centrada02-m30.csalidas02 as enero, m30.centrada03-m30.csalidas03 as febrero, m30.centrada04-m30.csalidas04 as marzo, m30.centrada05-m30.csalidas05 as abril, m30.centrada06-m30.csalidas06 as mayo, m30.centrada07-m30.csalidas07 as junio, m30.centrada08-m30.csalidas08 as julio, m30.centrada09-m30.csalidas09 as agosto, m30.centrada10-m30.csalidas10 as septiembre,m30.centrada11-m30.csalidas11 as octubre,m30.centrada12-m30.csalidas12 as noviembre,m30.centrada13-m30.csalidas13 as diciembre " +
            " from mgw10031 m31 join mgw10030 m30 on m30.cidejerc01 = m31.cidejerc01 " +
            " join mgw10005 m5 on m5.cidprodu01 = m30.cidprodu01 " +
            " join mgw10003 m3 on m3.cidalmacen = m30.cidalmacen " +
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

        protected virtual void mLeerDireccion()
        {


            //_RegDoctoOrigen._RegDireccion; 
        }

        private Boolean mGrabarInterfaz(string aFolio, int aTipo)
        {
            OleDbConnection lconexion = new OleDbConnection();
            string amensaje = "";
            lconexion = miconexion.mAbrirRutaGlobal(out amensaje);
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

            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            long lidclien;
            //OleDbParameter lparametroIdDocumento = new OleDbParameter("@p1", _RegDoctoOrigen.cIdDocto);
            lsql.CommandText = "select max(cidclien01) + 1 as cidclien01 from mgw10002";
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            _RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                lreader.Read();
                lidclien = long.Parse(lreader["cidclien01"].ToString());
            }
            else
                lidclien = 1;
            lreader.Close();


            //OleDbConnection lconexion = new OleDbConnection();
            // lconexion = miconexion.mAbrirConexionDestino ();
            bool lrespuesta = false;
            string lfecha = _RegDoctoOrigen.cFecha.ToString();
            DateTime ldate = DateTime.Parse(lfecha);
            lfecha = ldate.ToString("dd/MM/yyyy");

            //lconexion = miconexion.mAbrirConexionDestino();
            string lcadena = "insert into mgw10002 (cidclien01, ccodigoc01,crazonso01,cfechaalta,crfc,cidmoneda, clistapr01, ctipocli01,cestatus) values (" +
                lidclien +
                ",'" + _RegDoctoOrigen.cCodigoCliente + "','" + _RegDoctoOrigen.cRazonSocial + "'," +
                "ctod('" + lfecha + "'),'" +
                _RegDoctoOrigen.cRFC + "'" +
                ",1,1,1,1)";
            OleDbCommand lsql1 = new OleDbCommand(lcadena, miconexion._conexion);
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



        protected virtual bool mActualizaDocumento(long liddocum, int aopcion, double afolionuevo)
        {
            //miconexion.mAbrirConexionDestino();
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
                long cnatural = 0;
                long cescfd = 0;
                //string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
                //"ConceptoDocumento"
                string lCodigoConcepto = _RegDoctoOrigen.cCodigoConcepto;
                lsql.CommandText = "select cidconce01, ciddocum01,cseriepo01, cverfacele, cnatural01, cescfd from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
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
                    cnatural = long.Parse(lreader["cnatural01"].ToString());
                    cescfd = long.Parse(lreader["cescfd"].ToString());
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


                string lcadena1 = "update mgw10008 set cunidade01 = ctotalun01, cescfd = " + cescfd.ToString() + ", cnatural01= " + cnatural.ToString() + " where ciddocum01 = " + liddocum;

                OleDbCommand lsql3 = new OleDbCommand(lcadena1, miconexion._conexion);

                lsql3.ExecuteNonQuery();


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
                //miconexion.mCerrarConexionDestino();
            }
            //.mCerrarConexionDestino ();



            return lrespuesta;

        }
        private bool mActualizaDocumentoTraslado(long liddocum)
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
            int cnatural = 0;
            string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("Concepto").ToString().Trim();
            lsql.CommandText = "select cescfd,cnatural01  from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                lreader.Read();
                cescfd = int.Parse(lreader["cescfd"].ToString());
                cnatural = int.Parse(lreader["cnatural01"].ToString());
            }

            lreader.Close();
            if (cescfd == 0)
                return true;


            //OleDbParameter lparametroIdDocumento = new OleDbParameter("@p1", _RegDoctoOrigen.cIdDocto);
            string lcadena = "update mgw10008 set cescfd = 1, cnatural01= " + cnatural.ToString() + " where ciddocum01 = " + liddocum;

            OleDbCommand lsql1 = new OleDbCommand(lcadena, miconexion._conexion);

            lsql1.ExecuteNonQuery();


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


            /* lcadena = "insert into mgw10045 (cidfoldig,ciddoctode,cidcptodoc,ciddocto,cserie,cfolio,cestado,centregado, cfechaemi,cestrad,ctotal) " +
                                  " values (" + cidfoldig + "," + ciddocum01 + "," + cidconce + "," + liddocum + ",'" + cserie.Trim() + "'," + x + ",1, 0, ctod('" + lfecha + "'),3," + ctotal + ")";
             //lcadena = "insert into mgw10045 (cidfoldig,ciddoctode,cidcptodoc,ciddocto,cserie,cfolio,cestado, cfechaemi,cestrad) " +
             //                 "values (8,4,3001,11,'B',444,1,ctod('" + lfecha + "'),3)";
             OleDbCommand lsql2 = new OleDbCommand(lcadena, miconexion._conexion);
             lsql1.CommandText = "SET NULL OFF";
             lsql1.ExecuteNonQuery();

             lsql2.ExecuteNonQuery();*/
            lrespuesta = true;


            miconexion.mCerrarConexionDestino();


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
                lsql.CommandText = "select top 1 cidfoldig, cfolio, cserie from mgw10045 where ciddocto = 0 and cidcptodoc = " + cidconce + " order by cidfoldig asc";

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
                x.Actualizar(error);
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
                " and dtos(m8.cfecha) <= '" + sfechaf +
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
            double aFolio = 0;
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
            fSiguienteFolioComercial(concepto, ref aSerie, ref aFolio);
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
                fAltaMovimientoSeriesCapas_ParamComercial(movto.cIdMovto.ToString().Trim(), movto.cUnidades.ToString().Trim(), "1", "", movto._RegCapa.Pedimento, movto._RegCapa.NoAduana.ToString().Trim(), lfecha, "", "", "");
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

    /*    public void mValidaClienteProveedor(RegDocto adocto, int grabacliente = 1)
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

        */

        public bool mValidaProducto(RegMovto amovto, ref string lidunidad, int ConCapas = 1, int sat33 = 0)
        {
            //string lidunidad="";

            SqlCommand m = new SqlCommand();



            //return 1011;
            if (amovto._RegProducto.CodigoMedidaPesoSAT != null)

                m.CommandText = "SELECT CIDUNIDAD FROM admUnidadesMedidaPeso where CCLAVEINT ='" + amovto._RegProducto.CodigoMedidaPesoSAT + "'";

            else
                m.CommandText = "SELECT CIDUNIDADBASE FROM admProductos where Ccodigoproducto ='" + amovto.cCodigoProducto.Trim() + "'";

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


                StringBuilder aValor1 = new StringBuilder(12);
                int lret2 = fLeeDatoProductoComercial("cIdProducto", aValor1, 12);
                int lidproducto = int.Parse(aValor1.ToString());


                m.CommandText = "insert into admDatosAddenda values (367,2," + lidproducto.ToString() + ",4,'" + amovto._RegProducto.ComercioExterior + "')";
                m.Connection = miconexion._conexion1;
                m.ExecuteNonQuery();



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

        private int mGrabaEncabezadoComercial(RegDocto doc, int incluyedireccion, ref int aIdDocumento, ref long aFolio1, ref string aSerie, int conComercioExterior = 0)
        {
            int lret2 = 0;
            int lerrordocto = 0;
            StringBuilder sMensaje1 = new StringBuilder(512);
            string aCodigoConcepto = "";
            string ltextoextra1cliente = "";

            if (conComercioExterior == 1)
            {
                SqlCommand lsql = new SqlCommand();
                lsql.CommandText = "select ctextoextra1 from admClientes where ccodigocliente = '" + doc.cCodigoCliente + "'";
                lsql.Connection = miconexion._conexion1;

                SqlDataReader l;
                l = lsql.ExecuteReader();
                if (l.HasRows)
                {
                    l.Read();
                    ltextoextra1cliente = l["ctextoextra1"].ToString().Trim();
                }
                l.Close();
            }



            double aFolio = 0;
            if (doc.cFolio == 0)
            {
                try
                {
                    // int z = fSiguienteFolioComercial(doc.cCodigoConcepto, ref  aSerie, ref  aFolio);
                    aFolio1 = long.Parse(aFolio.ToString());
                    aFolio = -1;
                }
                catch (Exception ii)
                {
                }
            }
            else
            {
                aFolio = doc.cFolio;
            }




            if (aFolio == 0)
            {
                aFolio = 1;
                aFolio1 = long.Parse(aFolio.ToString());
            }



            fInsertarDocumentoComercial();

            lret2 = fSetDatoDocumentoComercial("cCodigoConcepto", doc.cCodigoConcepto);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                //fProcesaError(doc, doc.cIdDocto, "El documento con cliente " + doc.cCodigoCliente.Trim() + " y folio " + doc.cFolio.ToString() + " presenta el sig. problema " + sMensaje1.ToString(), ref lret2);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
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
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    return 0;
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
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }
            lret2 = fSetDatoDocumentoComercial("cRFC", doc.cRFC);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }





            //lret2 = fSetDatoDocumentoComercial("cIdMoneda", "2");
            //if (lret2 != 0)
            //    fErrorComercial(lret2, sMensaje1, 512);

            //lret2 = fSetDatoDocumentoComercial("cTipoCambio", doc.cTipoCambio.ToString().Trim());
            //if (lret2 != 0)
            //    fErrorComercial(lret2, sMensaje1, 512);

            //DateTime lFechaVencimiento = DateTime.Today;
            string lfechavenc = String.Format("{0:MM/dd/yyyy}", DateTime.Today);
            lfechavenc = String.Format("{0:MM/dd/yyyy}", doc.cFecha);
            lret2 = fSetDatoDocumentoComercial("cFecha", lfechavenc);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }

            if (aFolio != -1)
            {
                lret2 = fSetDatoDocumentoComercial("cFolio", aFolio.ToString());
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    return 0;
                }
            }
            lret2 = fSetDatoDocumentoComercial("cFechaVencimiento", lfechavenc);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }

            lret2 = fSetDatoDocumentoComercial("cCodigoCliente", doc.cCodigoCliente);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }

            RegCliente lc = new RegCliente();

            lc = mBuscarClienteComercial(doc.cCodigoCliente);



            lret2 = fSetDatoDocumentoComercial("CMETODOPAG", doc.cFormaPago);
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }

            if (doc.cMetodoPago == "PPD")
            {
                lret2 = fSetDatoDocumentoComercial("CCANTPARCI", "2");
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    return 0;
                }
            }

            if (doc.cUsoCFDI != "")
            {
                lret2 = fSetDatoDocumentoComercial("CCODCONCBA", doc.cUsoCFDI);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    return 0;
                }
            }


            lret2 = fGuardaDocumentoComercial();
            if (lret2 != 0)
            {
                fErrorComercial(lret2, sMensaje1, 512);
                fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                return 0;
            }

            StringBuilder aValor = new StringBuilder(12);
            lret2 = fLeeDatoDocumentoComercial("CIDDOCUMENTO", aValor, 12);
            int liddocumento = int.Parse(aValor.ToString());


            lret2 = fLeeDatoDocumentoComercial("CFOLIO", aValor, 12);
            long llfolio = Convert.ToInt32(decimal.Parse(aValor.ToString()));

            lret2 = fLeeDatoDocumentoComercial("CSERIEDOCUMENTO", aValor, 12);
            string lSerie = aValor.ToString();
            aSerie = lSerie;

            doc.cFolio = llfolio;
            doc.cSerie = lSerie;

            doc.cIdDocto = liddocumento;
            /*if (incluyedireccion == 1)
                lret2 = mGrabaDireccionComercial(doc);*/
            lret2 = mgrabamoneda(liddocumento, doc.cMoneda, doc.cTipoCambio);

            if (conComercioExterior == 1)
                lret2 = mgrabacomercioexterior(liddocumento, ltextoextra1cliente, doc.cTipoCambio);

            return 1;
        }

        private int mgrabacomercioexterior(int aiddocumento, string atexto, decimal aTC)
        {
            // also comercio exterior
            SqlCommand lsql = new SqlCommand();
            lsql.Connection = miconexion._conexion1;

            lsql.CommandText = "insert into admDatosAddenda values (367,3," + aiddocumento.ToString() + ",2,'Exportación')";
            int lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (367,3," + aiddocumento.ToString() + ",3,'IMPORTACION O EXPORTACION DEFINITIVA')";
            lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (367,3," + aiddocumento.ToString() + ",4,'No Funge como certificado de origen')";
            lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (367,3," + aiddocumento.ToString() + ",7,'" + atexto + "')";
            lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (367,3," + aiddocumento.ToString() + ",8,'No tiene subdivisión')";
            lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (367,3," + aiddocumento.ToString() + ",10," + aTC.ToString().Trim() + ")";
            lret1 = lsql.ExecuteNonQuery();

            /*            insert into admDatosAddenda values (367,3,7,2,'Exportación')
            insert into admDatosAddenda values (367,3,7,3,'IMPORTACION O EXPORTACION DEFINITIVA')
            insert into admDatosAddenda values (367,3,7,4,'No Funge como certificado de origen')
            insert into admDatosAddenda values (367,3,7,7,'DAP - ENTREGADA EN LUGAR')
            insert into admDatosAddenda values (367,3,7,8,'No tiene subdivisión')
            insert into admDatosAddenda values (367,3,7,10,'18.58')*/
            //--lsql.Connection = miconexion._conexion1;
            return 1;

        }


        private int mgrabaorigen(RegDocto docto, long liddocumento)
        {
            SqlCommand lsql = new SqlCommand();
            lsql.Connection = miconexion._conexion1;
            lsql.CommandText = "update admdocumentos set ctotalunidades=" + docto.cTotalUnidades.ToString()
                + ",cunidadespendientes=" + docto.cTotalUnidades.ToString() + ",ciddocumentoorigen=" + docto.cIdDocto.ToString()
                + " where ciddocumento = " + liddocumento.ToString();

            int lret1 = lsql.ExecuteNonQuery();
            return 1;

        }

        private int mgrabalugarexpedicionyAddenda(RegDocto docto)
        {
            // also comercio exterior
            SqlCommand lsql = new SqlCommand();
            lsql.Connection = miconexion._conexion1;
            //lsql.CommandText = "update admdocumentos set clugarexpedicion = " + lugarexpedicion + " where ciddocumento =" + aiddocumento.ToString(); 
            //int lret1 = lsql.ExecuteNonQuery();

            int lret1 = 0;
            int lndex = 1;

            if (docto.cTextoExtra1 == "F")
                lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + "," + "1".ToString() + ",'FACTURA')";

            if (docto.cTextoExtra1 == "NC")
                lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + "," + "1".ToString() + ",'NOTA DE CREDITO')";

            if (docto.cTextoExtra1 == "ND")
                lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + "," + "1".ToString() + ",'NOTA DE CARGO')";

            lret1 = lsql.ExecuteNonQuery();

            /*Folio Unico de Factura FUF 
                Fecha de la Factura 
                Fecha Limite de Pago 
                Cuenta de Orden del PM  
                Nombre del Banco 
                Sucursal del Banco  
                Numero de Cuenta del Proveedor 
                Numero de Cuenta CLABE del Proveedor 
                Referencia del Banco    
                Contacto del Proveedor
                */
            lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + ",2,'" + docto.addendiux.FolioUnicodeFacturaFUF + "')";
            DateTime xx;
            string z;
            string x;
            try
            {
                x = docto.addendiux.FechadelaFactura;
                //xx = DateTime.Parse(docto.addendiux.FechadelaFactura);
                //z = x.Substring(0, 2) + "-" + x.Substring(3, 2) + "-" + x.Substring(6, 4);
                //20210602
                //z = x.Substring(6, 2) + "-" + x.Substring(4, 2) + "-" + x.Substring(2, 2);
                z = x;
            }
            catch (Exception eee)
            { z = docto.addendiux.FechadelaFactura; }



            lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + ",2,'" + docto.addendiux.FolioUnicodeFacturaFUF + "')";
            lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + ",3,'" + z + "')";
            lret1 = lsql.ExecuteNonQuery();

            try
            {
                x = docto.addendiux.FechaLimitedePago;

                //2021/06/02

                //xx = DateTime.Parse(docto.addendiux.FechaLimitedePago);
                //z = x.Substring(0, 2) + "-" + x.Substring(3, 2) + "-" + x.Substring(6, 4);
                //z = x.Substring(6, 2) + "-" + x.Substring(4, 2) + "-" + x.Substring(2, 2);
                z = x;
            }
            catch (Exception eee)
            { z = docto.addendiux.FechaLimitedePago; }

            lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + ",4,'" + z + "')";
            lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + ",5,'" + docto.addendiux.CuentadeOrdendelPM + "')";
            lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + ",6,'" + docto.addendiux.NombredelBanco + "')";
            lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + ",7,'" + docto.addendiux.SucursaldelBanco + "')";
            lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + ",8,'" + docto.addendiux.NumerodeCuentadelProveedor + "')";
            lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + ",9,'" + docto.addendiux.NumerodeCuentaCLABEdelProveedor + "')";
            lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + ",10,'" + docto.addendiux.ReferenciadelBanco + "')";
            lret1 = lsql.ExecuteNonQuery();
            lsql.CommandText = "insert into admDatosAddenda values (484,3," + docto.cIdDocto.ToString() + ",11,'" + docto.addendiux.ContactodelProveedor + "')";
            lret1 = lsql.ExecuteNonQuery();


            //Num Linea   Folio Unico Concepto Cantidad    Unidad Precio Unitario Importe Linea Importe Orig Importe Modif Monto Ajuste IVA Total Monto Letra

            lndex = 1;
            foreach (cAddendaMovimiento y in docto.addendiux.lista)
            {
                long lidmov = docto.addendiux.lista[lndex - 1].idmovim;
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",12,'" + y.NumLinea + "')";
                lret1 = lsql.ExecuteNonQuery();
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",13,'" + y.FolioUnico + "')";
                lret1 = lsql.ExecuteNonQuery();
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",14,'" + y.Concepto + "')";
                lret1 = lsql.ExecuteNonQuery();
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",15,'" + y.Cantidad + "')";
                lret1 = lsql.ExecuteNonQuery();
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",16,'" + y.Unidad + "')";
                lret1 = lsql.ExecuteNonQuery();
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",17,'" + y.PrecioUnitario + "')";
                lret1 = lsql.ExecuteNonQuery();
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",18,'" + y.ImporteLinea + "')";
                lret1 = lsql.ExecuteNonQuery();
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",19,'" + y.ImporteOrig + "')";
                lret1 = lsql.ExecuteNonQuery();
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",20,'" + y.ImporteModif + "')";
                lret1 = lsql.ExecuteNonQuery();
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",21,'" + y.MontoAjuste + "')";
                lret1 = lsql.ExecuteNonQuery();
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",22,'" + y.IVA + "')";
                lret1 = lsql.ExecuteNonQuery();
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",23,'" + y.Total + "')";
                lret1 = lsql.ExecuteNonQuery();
                lsql.CommandText = "insert into admDatosAddenda values (484,5," + lidmov.ToString() + ",24,'" + y.MontoLetra + "')";
                lret1 = lsql.ExecuteNonQuery();










                lndex++;


            }


            /*lndex++;
            foreach (string x in docto._Addendas)
            {
                DateTime xx;
                string z;
                try
                {
                    xx = DateTime.Parse(x);
                    z = x.Substring(0, 2) + "-" + x.Substring(3, 2) + "-" + x.Substring(6, 4);
                                        }
                catch (Exception eee)
                { z = x; }
              lsql.CommandText = "insert into admDatosAddenda values (375,3," + docto.cIdDocto.ToString() + "," + lndex.ToString()+ ",'" + z + "')";
              lret1 = lsql.ExecuteNonQuery();
                lndex++;
            }*/

            if (docto.cTipoRelacion != "" && docto.cUUID != "")
            {
                string tiporelacion = docto.cTipoRelacion.Substring(0, 2);
                string comando = "update admFoliosDigitales set ccadpedi = '" + docto.cUUID + "', cusuban02 = '" + tiporelacion + "'  where ciddocto =" + docto.cIdDocto;
                SqlCommand lsql4 = new SqlCommand(comando, miconexion._conexion1);
                lsql4.ExecuteNonQuery();
            }


            return 1;

        }

        private int mGrabarMovimientosComercial(RegDocto doc, int indicedoc, ref decimal ltotaunidadesdocto, int concomercioexterior = 0, int traspaso = 0)
        {
            int lret2 = 0;
            StringBuilder sMensaje1 = new StringBuilder(512);
            int lerrordocto = 0;
            //decimal ltotaunidadesdocto = 0;
            int lerrormovto = 0;
            int indicemov = 0;
            int indicemov1 = 0;
            int lcontrolexistencia = 0;




            foreach (RegMovto movto in doc._RegMovtos)
            {

                if (lerrormovto != 0)
                    continue;


                if (movto.cError != "")
                {
                    if (doc.cCodigoConcepto == "35")
                        fProcesaError(doc, movto, "Mov", movto.cError, 0);
                    continue;

                }




                fInsertarMovimientoComercial();
                string lidunidad = "";
                //mValidaProducto(movto, ref lidunidad, 0, 1);
                lret2 = fSetDatoMovimientoComercial("cIdDocumento", doc.cIdDocto.ToString().Trim());
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    //fProcesaError("El producto " + movto.cCodigoProducto + " " + sMensaje1.ToString(), ref lerrormovto);
                    fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                    return 0;
                }

                int ltipoproducto;
                StringBuilder aValor1 = new StringBuilder(12);
                try
                {
                    lret2 = fSetDatoMovimientoComercial("cCodigoProducto", movto.cCodigoProducto.Trim());
                }
                catch (Exception eee)
                {
                    if (eee.Message.Contains("Intento"))
                    {
                        fCancelaCambiosMovimientoComercial();
                        fInsertarMovimientoComercial();
                        lret2 = fSetDatoMovimientoComercial("cIdDocumento", doc.cIdDocto.ToString().Trim());

                        lret2 = fSetDatoMovimientoComercial("cCodigoAlmacen", movto.cCodigoAlmacen);

                        lret2 = fSetDatoMovimientoComercial("cCodigoProducto", "z");
                        lret2 = fSetDatoMovimientoComercial("cCodigoProducto", movto.cCodigoProducto.Trim());



                    }

                }
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                    return 0;
                }
                else
                {
                    int busca = fBuscaProductoComercial(movto.cCodigoProducto.Trim());

                    if (busca != 0)
                    {
                        lret2 = fLeeDatoProductoComercial("cTipoProducto", aValor1, 12);
                        int lidproducto = int.Parse(aValor1.ToString());

                    }
                    else
                        if (traspaso == 1)
                    {
                        lret2 = fLeeDatoProductoComercial("cControlExistencia", aValor1, 12);
                        lcontrolexistencia = int.Parse(aValor1.ToString());
                    }

                }
                lret2 = fSetDatoMovimientoComercial("cCodigoAlmacen", movto.cCodigoAlmacen);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                    return 0;
                }
                //int lRet3 = fSetDatoMovimientoComercial("cUnidadesCapturadas", movto.cUnidades.ToString().Trim());


                lret2 = fSetDatoMovimientoComercial("CIDUNIDAD", lidunidad);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                    return 0;
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
                            fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
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
                    if (movto.idseries.Count > 0)
                    {
                        lret2 = fGuardaMovimientoComercial();
                        if (lret2 != 0)
                        {
                            fErrorComercial(lret2, sMensaje1, 512);
                            // MessageBox.Show("Error: " + sMensaje);
                            lret2 = fGuardaMovimientoComercial();
                            if (lret2 != 0)
                            {
                                fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                                continue;
                            }
                        }
                        StringBuilder aValor = new StringBuilder(12);
                        aValor.Length = 0;
                        lret2 = fLeeDatoMovimientoComercial("CIDMOVIMIENTO", aValor, 12);
                        int lidmovimiento = int.Parse(aValor.ToString());

                        SqlCommand lsql1 = new SqlCommand();
                        lsql1.CommandText = "insert into admmovimientosserie" +
                            " select " + lidmovimiento.ToString() + ", cidserie, getdate() from admMovmientosSeriePedido " +
                            " where cidmovimiento =" + movto.cIdMovto;

                        lsql1.Connection = miconexion._conexion1;
                        int cuantos = lsql1.ExecuteNonQuery();

                        lsql1.CommandText = "update admmovimientos set " +
                                            " cunidades = " + cuantos +
                                            " ,CUNIDADESCAPTURADAS = " + cuantos.ToString() +
                                            " ,cprecio = " + movto.cPrecio +
                                            " ,CPRECIOCAPTURADO = " + movto.cPrecio +
                                            " ,CCOSTOESPECIFICO = (select sum(ccosto) " +
                                            " from admMovmientosSeriePedido ms " +
                                            " join admNumerosSerie s on ms.cidserie = s.CIDSERIE) " +
                                            " , cneto = " + movto.cneto.ToString() +
                                            " , cimpuesto1 = " + movto.cImpuesto.ToString() +
                                            " , ctotal = " + movto.cTotal.ToString() +
                                            " , cunidadespendientes = " + cuantos.ToString() +
                                            " where cidmovimiento = " + lidmovimiento.ToString();

                        cuantos = lsql1.ExecuteNonQuery();

                        lsql1.CommandText = "update admdocumentos set " +
                                            " ctotalunidades = ctotalunidades  + " + cuantos +
                                            " ,Cpendiente = cpendiente + " + movto.cTotal.ToString() +
                                            " , cneto = cneto + " + movto.cneto.ToString() +
                                            " , cimpuesto1 = cimpuesto1 +" + movto.cImpuesto.ToString() +
                                            " , ctotal = ctotal + " + movto.cTotal.ToString() +
                                            " where ciddocumento = " + doc.cIdDocto.ToString();

                        cuantos = lsql1.ExecuteNonQuery();


                        lsql1.CommandText = "update admnumerosserie set cestado = 7 where cidserie in " +
                            "(select  cidserie from admMovmientosSeriePedido " +
                            " where cidmovimiento =" + movto.cIdMovto + ")";


                        cuantos = lsql1.ExecuteNonQuery();


                    }
                    else
                    {
                        if (traspaso == 1 && (lcontrolexistencia == 16 || lcontrolexistencia == 9 || lcontrolexistencia == 4))
                        {
                            movto.cUnidades = 0;
                            ltotaunidadesdocto = 1;
                        }
                        lRet = fSetDatoMovimientoComercial("cUnidadesCapturadas", movto.cUnidades.ToString());
                        if (lRet != 0)
                        {
                            fErrorComercial(lRet, sMensaje1, 512);
                            fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                            fBorraMovimientoComercial();
                            continue;


                            return 0;
                        }


                        lRet = fSetDatoMovimientoComercial("cImporteExtra1", movto.cimporteextra1.ToString());
                        if (lRet != 0)
                        {
                            fErrorComercial(lRet, sMensaje1, 512);
                            fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                            fBorraMovimientoComercial();
                            continue;


                            return 0;
                        }
                        lRet = fSetDatoMovimientoComercial("cImporteExtra2", movto.cimporteextra2.ToString());
                        if (lRet != 0)
                        {
                            fErrorComercial(lRet, sMensaje1, 512);
                            fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                            fBorraMovimientoComercial();
                            continue;


                            return 0;
                        }

                        lRet = fSetDatoMovimientoComercial("cTextoExtra3", movto.ctextoextra3.ToString());
                        if (lRet != 0)
                        {
                            fErrorComercial(lRet, sMensaje1, 512);
                            fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                            fBorraMovimientoComercial();
                            continue;


                            return 0;
                        }

                        if (traspaso == 1 && lcontrolexistencia == 1)
                            ltotaunidadesdocto += movto.cUnidades;





                        if (traspaso == 0)
                            if (doc.cCodigoConcepto == "34" || doc.cCodigoConcepto == "340" || doc.cCodigoConcepto == "35")
                            {
                                lRet = fSetDatoMovimientoComercial("cCostoCapturado", movto.cPrecio.ToString().Trim());
                                if (lRet != 0)
                                {
                                    fErrorComercial(lRet, sMensaje1, 512);
                                    fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                                    return 0;
                                }
                            }

                        lRet = fSetDatoMovimientoComercial("cPrecioCapturado", movto.cPrecio.ToString().Trim());
                        if (lRet != 0)
                        {
                            fErrorComercial(lRet, sMensaje1, 512);
                            fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                            return 0;
                        }
                        /*if (doc.cReferencia != "32")
                        lRet = fSetDatoMovimientoComercial("cImpuesto1", movto.cImpuesto.ToString().Trim());
                        if (lRet != 0)
                        {
                            fErrorComercial(lRet, sMensaje1, 512);
                            fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                            return 0; 
                        }*/


                        lRet = fGuardaMovimientoComercial();
                        if (lRet != 0)
                        {
                            fErrorComercial(lRet, sMensaje1, 512);
                            lret2 = fGuardaMovimientoComercial();
                            if (lret2 != 0)
                            {
                                fErrorComercial(lRet, sMensaje1, 512);
                                fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                                continue;
                                //return 0;
                            }

                        }
                    }
                }
                /*else
                {
                    if (traspaso == 1 ) { 
                        SqlCommand lsqlnummov = new SqlCommand();

                        lsqlnummov.CommandText = "update admmovimientos set cnumero admDatosAddenda values (367,5," + aValor1.ToString() + ",2,'1.')";
                        lsqlnummov.Connection = miconexion._conexion1;
                        int lret = lsqlnummov.ExecuteNonQuery();
                }*/

                // StringBuilder aValor1 = new StringBuilder(12);
                aValor1.Length = 0;
                lret2 = fLeeDatoMovimientoComercial("CIDMOVIMIENTO", aValor1, 12);
                int lidmovimiento1 = int.Parse(aValor1.ToString());
                movto.cIdMovto = lidmovimiento1;
                if (doc.addendiux.lista.Count > 0)
                    doc.addendiux.lista[indicemov1++].idmovim = lidmovimiento1;

                SqlCommand lsql = new SqlCommand();

                if (movto._RegProducto.UnidadMicroplaneComercioExterior == "KGS")
                {
                    lsql.CommandText = "insert into admDatosAddenda values (367,5," + aValor1.ToString() + ",2,'1.')";

                    lsql.CommandText = "insert into admDatosAddenda values (367,5," + aValor1.ToString() + ",2,'" + movto.cUnidades.ToString() + "')";
                    lsql.Connection = miconexion._conexion1;
                    int lret = lsql.ExecuteNonQuery();
                }
                if (movto.cIdMovto != 0)
                {
                    lsql.CommandText = "update admMovimientos set cidmovtoorigen =" + movto.cIdMovto.ToString() + ",cunidadesorigen=" + movto.cUnidades +
                        "where cidmovimiento =" + lidmovimiento1.ToString();
                    lsql.Connection = miconexion._conexion1;
                    int lret = lsql.ExecuteNonQuery();
                }
                //return lret;


                //int lret1 = lsql.ExecuteNonQuery();

                if (movto.cIdMovtoOrigen > 0)
                {
                    lsql.CommandText = "update admMovimientos set cimporteextra1 = cimporteextra1 - " + movto.cUnidades +
                        " where cidmovimiento =" + movto.cIdMovtoOrigen.ToString();
                    lsql.Connection = miconexion._conexion1;
                    int lret = lsql.ExecuteNonQuery();
                }

            }
            if (concomercioexterior == 2)
            {
                lret2 = mgrabalugarexpedicionyAddenda(doc);
            }
            return 1;
        }

        public string mGrabarDoctosComercial(int incluyetimbrado, ref long lultimoFolio, int incluyedireccion, int concomercioexterior, int grabarcliente)
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

            int zzzzz = fAbreEmpresa(rutadestino);






            int indicedoc = 0;
            int lret2;
            int lcuantos = _RegDoctos.Count;
            int ltotales = lcuantos;
            int lindice = 1;
            int liddocumento = 0;
            decimal ltotalunidadesdocto = 0;

            foreach (RegDocto doc in _RegDoctos)
            {

                long aFolio = 0;
                string aSerie = "";
                int lRetorno = mGrabaEncabezadoComercial(doc, incluyedireccion, ref liddocumento, ref aFolio, ref aSerie, concomercioexterior, grabarcliente);
                if (lRetorno == 1)
                {
                    if (doc.cFolio == 0)
                    {
                        doc.cFolio = aFolio;
                        doc.cSerie = aSerie;
                    }
                    lRetorno = mGrabarMovimientosComercial(doc, indicedoc, ref ltotalunidadesdocto);
                    indicedoc++;

                }

                Notificar((double)(lindice++ * 100) / lcuantos);

                if (lRetorno == 1)
                {
                    mGrabarUnidadesDocto(doc.cIdDocto, ltotalunidadesdocto);


                    if (incluyetimbrado == 1)
                    {

                        string lpass = "";
                        lpass = GetSettingValueFromAppConfigForDLL("Pass").ToString().Trim();

                        lultimoFolio = doc.cFolio;
                        double dfolio = 0.0;
                        dfolio = Convert.ToDouble(doc.cFolio);

                        int lresp20 = fEmitirDocumentoComercial(doc.cCodigoConcepto, doc.cSerie, dfolio, lpass, "");
                        if (lresp20 != 0)
                        {
                            fErrorComercial(lresp20, sMensaje1, 512);
                            fProcesaError(doc, null, "Doc", sMensaje1.ToString(), 0);


                        }
                        else
                        {
                            //lresp20 = fEntregEnDiscoXMLComercial(doc.cCodigoConcepto, doc.cSerie, doc.cFolio, 0, @"C:\Compac\Empresas\Reportes\Formatos Digitales\reportes_Servidor\COMERCIAL\Factura.rdl");
                            lresp20 = fEntregEnDiscoXMLComercial(doc.cCodigoConcepto, doc.cSerie, doc.cFolio, 0, "");
                            if (lresp20 != 0)
                            {
                                fErrorComercial(lresp20, sMensaje1, 512);
                                fProcesaError(doc, null, "Doc", sMensaje1.ToString(), 0);
                            }
                            if (doc.cNombreArchivo != "")
                            {
                                string archivoorigencompleto = @GetSettingValueFromAppConfigForDLL("RutaOrigen").ToString().Trim() + @"\" + doc.cNombreArchivo;
                                string archivodestinocompleto = @GetSettingValueFromAppConfigForDLL("RutaBien").ToString().Trim() + @"\" + doc.cNombreArchivo;
                                try
                                {
                                    System.IO.File.Move(archivoorigencompleto, archivodestinocompleto);
                                }
                                catch (Exception aaaa)
                                { }


                            }
                        }
                    }


                    //lexitosos++;
                }
                //else
                //  fBorraDocumentoComercial();
                //Notificar((double)(lindice++ * 100) / lcuantos);

            }










            fCierraEmpresa();

            return "";

        }

        public void mCerrarSdkComercial()
        {
            try
            {

                miconexion.mCerrarConexionOrigenComercial();
                fTerminaSDK();
            }
            catch (Exception eeeeee)
            {
            }

        }



        public string mGrabarDoctosComercialborrar(List<RegDocto> Doctos, ref int lexitosos, ref int ltotales, int incluyedireccion = 1)
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
                    fSiguienteFolioComercial(doc.cCodigoConcepto, ref aSerie, ref aFolio);
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
                        fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                        continue;
                    }
                }
                lret2 = fSetDatoDocumentoComercial("cRazonSocial", doc._RegCliente.RazonSocial);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    continue;
                }
                lret2 = fSetDatoDocumentoComercial("cRFC", doc.cRazonSocial);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    continue;
                }

                lret2 = fSetDatoDocumentoComercial("cCodigoConcepto", doc.cCodigoConcepto);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
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
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    continue;
                }

                lret2 = fSetDatoDocumentoComercial("cFolio", aFolio.ToString());
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    continue;
                }
                lret2 = fSetDatoDocumentoComercial("cFechaVencimiento", lfechavenc);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    continue;
                }

                lret2 = fSetDatoDocumentoComercial("cCodigoCliente", doc.cCodigoCliente);
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
                    continue;
                }
                lret2 = fGuardaDocumentoComercial();
                if (lret2 != 0)
                {
                    fErrorComercial(lret2, sMensaje1, 512);
                    fProcesaError(doc, null, "Doc", sMensaje1.ToString());
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
                    mValidaProducto(movto, ref lidunidad);
                    lret2 = fSetDatoMovimientoComercial("cIdDocumento", liddocumento.ToString().Trim());
                    if (lret2 != 0)
                    {
                        fErrorComercial(lret2, sMensaje1, 512);
                        fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                        continue;
                    }
                    lret2 = fSetDatoMovimientoComercial("cCodigoProducto", movto.cCodigoProducto);
                    if (lret2 != 0)
                    {
                        fErrorComercial(lret2, sMensaje1, 512);
                        fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                        continue;
                    }
                    lret2 = fSetDatoMovimientoComercial("cCodigoAlmacen", movto.cCodigoAlmacen);
                    if (lret2 != 0)
                    {
                        fErrorComercial(lret2, sMensaje1, 512);
                        fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
                        continue;
                    }
                    //int lRet3 = fSetDatoMovimientoComercial("cUnidadesCapturadas", movto.cUnidades.ToString().Trim());

                    lret2 = fSetDatoMovimientoComercial("CIDUNIDAD", "1");
                    if (lret2 != 0)
                    {
                        fErrorComercial(lret2, sMensaje1, 512);
                        fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
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
                            fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
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
                    StringBuilder sIdproducto = new StringBuilder(12);
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
                        fProcesaError(doc, movto, "Mov", sMensaje1.ToString());
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
        /* private void fProcesaError(RegDocto doc,  long id, string error, ref int lerrormovto)
         {
             if (id !=0)
                 mBorrarDocto(doc);
             Notificar(error);
             lerrormovto = 1;

         }*/


        private void fProcesaError(RegDocto doc, RegMovto movto, string tipo, string sMensaje1, int aBorrar = 1)
        {
            string error = "";
            string archivoorigencompleto = "";
            string archivodestinocompleto = "";
            if (doc.cIdDocto != 0 && aBorrar == 1)
                mBorrarDocto(doc);

            if (tipo == "Doc")
                if (doc.cNombreArchivo == "")
                {
                    if (doc.cCodigoCliente != "")
                        error = "El documento con cliente " + doc.cCodigoCliente.Trim() + " y folio " + doc.cFolio.ToString() + " presenta el sig. problema " + sMensaje1.ToString();
                    else
                        error = "El documento con  folio " + doc.cFolio.ToString() + " presenta el sig. problema " + sMensaje1.ToString();

                }
                else
                {
                    error = "El archivo " + doc.cNombreArchivo.Trim() + " que genero el documento con folio " + doc.cFolio.ToString() + " presenta el sig. problema " + sMensaje1.ToString();
                    archivoorigencompleto = @GetSettingValueFromAppConfigForDLL("RutaOrigen").ToString().Trim() + @"\" + doc.cNombreArchivo;
                    archivodestinocompleto = @GetSettingValueFromAppConfigForDLL("RutaMal").ToString().Trim() + @"\" + doc.cNombreArchivo;
                    try
                    {
                        System.IO.File.Move(archivoorigencompleto, archivodestinocompleto);
                    }
                    catch (Exception eeee)
                    {

                    }
                }
            else
                if (doc.cNombreArchivo == "")
                error = "El producto " + movto.cCodigoProducto + " del documento con cliente " + doc.cCodigoCliente.Trim() + " y folio " + doc.cFolio.ToString() + " presenta el sig. problema " + sMensaje1.ToString();
            else
            {
                error = "El producto " + movto.cCodigoProducto + "del archivo " + doc.cNombreArchivo.Trim() + " que genero el documento con folio " + doc.cFolio.ToString() + " presenta el sig. problema " + sMensaje1.ToString();
                archivoorigencompleto = @GetSettingValueFromAppConfigForDLL("RutaOrigen").ToString().Trim() + @"\" + doc.cNombreArchivo;
                archivodestinocompleto = @GetSettingValueFromAppConfigForDLL("RutaMal").ToString().Trim() + @"\" + doc.cNombreArchivo;
                try
                {
                    System.IO.File.Move(archivoorigencompleto, archivodestinocompleto);
                }
                catch (Exception eeee)
                {

                }

            }
            Notificar(error);

        }
        private string empresacomercial;

        public void mAsignaEmpresaComercial(RegConexion empresa)
        {
            miconexion.mAbrirConexionComercial(empresa, true);
            empresacomercial = empresa.database;
            StringBuilder sMensaje1 = new StringBuilder(512);

            string rutadestino = empresacomercial;

            /*int lcierraconexion = 0;
            if (miconexion._conexion1 == null)
            {
                miconexion.mAbrirConexionComercial(true);
                lcierraconexion = 1;
            }*/

            int lResultado = fSetNombrePAQ("CONTPAQ I Comercial");
            if (lResultado != 0)
            {
                fErrorComercial(lResultado, sMensaje1, 512);
            }
            /*            
                            if (incluyetimbrado == 1)
                            {
                                int lresp10 = fInicializaLicenseInfoComercial(0);
                                if (lresp10 != 0)
                                {
                                    fErrorComercial(lresp10, sMensaje1, 512);
                                }
                            }*/

            sdkcomercial = true;
            // }

            int zzzzz = fAbreEmpresa(rutadestino);
        }

        private void mBorrarDocto(RegDocto doc)
        {
            string lconcepto = GetSettingValueFromAppConfigForDLL("Concepto");
            long lret = fBuscarDocumentoComercial(lconcepto, doc.cSerie, doc.cFolio.ToString());
            if (lret == 0)
            {
                fBorraDocumentoComercial();
            }
        }

        public void mBorrarDocto(string Concepto, string Serie, string Folio)
        {
            StringBuilder sMensaje1 = new StringBuilder(512);



            long lret = fBuscarDocumentoComercial(Concepto, Serie, Folio);
            if (lret == 0)
            {
                int lresul = fBorraDocumentoComercial();
            }
            /* if(lcierraconexion == 1)
                 miconexion.mCerrarConexionOrigenComercial();*/


        }



        private void mBorrarDocto(long id)
        {

            SqlCommand lsql = new SqlCommand();
            lsql.CommandText = "  delete from admDocumentos WHERE CIDDOCUMENTO = " + id.ToString().Trim();
            lsql.Connection = miconexion._conexion1;
            int lret = lsql.ExecuteNonQuery();

        }

        /*
        private void fProcesaError(RegDocto doc)
        {
            Notificar(error);
            lerrormovto = 1;

        }*/

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
            if (aMoneda != "Peso Mexicano")
            {
                lsql.CommandText = "update admDocumentos set cidmoneda = 2 ,ctipocambio = " + aTC.ToString().Trim() + " where ciddocumento = " + aiddocumento.ToString().Trim();
                lsql.Connection = miconexion._conexion1;
                int lret = lsql.ExecuteNonQuery();
            }
            else
            {
                lsql.CommandText = "update admDocumentos set cidmoneda = 1 ,ctipocambio = 1 where ciddocumento = " + aiddocumento.ToString().Trim();
                lsql.Connection = miconexion._conexion1;
                int lret = lsql.ExecuteNonQuery();
            }




            return 1;
        }

        private int mGrabaDireccionComercial(RegDocto doc)
        {
            fInsertaDireccionComercial();
            int lret = fSetDatoDireccionComercial("CIDCATALOGO", doc.cIdDocto.ToString().Trim());
            lret = fSetDatoDireccionComercial("CTIPOCATALOGO", "3");
            lret = fSetDatoDireccionComercial("CTIPODIRECCION", "1");
            lret = fSetDatoDireccionComercial("CNOMBRECALLE", doc._RegDireccion.cNombreCalle);
            lret = fSetDatoDireccionComercial("CNUMEROEXTERIOR", doc._RegDireccion.cNumeroExterior);
            lret = fSetDatoDireccionComercial("CCOLONIA", doc._RegDireccion.cColonia);
            lret = fSetDatoDireccionComercial("CCODIGOPOSTAL", doc._RegDireccion.cCodigoPostal);
            //lret = fSetDatoDireccionComercial("CTELEFONO1", );
            lret = fSetDatoDireccionComercial("CEMAIL", doc._RegDireccion.cEmail);
            lret = fSetDatoDireccionComercial("CPAIS", "Mexico");
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
            " CPEDIMENTO,CADUANA,CFECHAPEDIMENTO,CTIPOCAMBIO,CEXISTENCIA,CCOSTO,CTIMESTAMP,CNUMADUANA) " +
            " values (" +
            lultimacapa + "," + movto._RegCapa._Almacen.Id.ToString() + "," + movto._RegProducto.Id.ToString().Trim() + ",'" + lfechahoy + "'," + "2,''," + "'18991230'," + "'18991230'," + "'" + movto._RegCapa.Pedimento + "',"
            + "'" + laduana + "','" + lfechaped + "'," + movto._RegCapa.tc.ToString().Trim() + "," + movto.cUnidades + "," + movto.cPrecio + ",'" + lfechatime + "'," + movto._RegCapa.NoAduana + ")";
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
            string lcadena = "";
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


        public void mGrabarDatosTraslado(List<RegProducto> lista, string aValor1, string aValor2, string aValor3, string aValor4)
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


        public List<RegProducto> mMostrarProductos(string aCodigo1, string aCodigo2, string aCodigo3, string aCodigo4, string aCodigo5, string aCodigo6)
        {
            OleDbConnection lconexion = new OleDbConnection();
            //OleDbDataReader lreader;
            List<RegProducto> lprods = new List<RegProducto>();
            string lcadena = "";



            lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {
                //lcadena = "select cidvalor01,cvalorcl01 from mgw10020 where ccodigov01 = '" + codigo + "' and cidclasi01 =" + anumClasif.ToString();
                lcadena = "select m5.* from mgw10005 m5 ";

                if (aCodigo1 != "")
                    lcadena += "join mgw10020 m20a " +
                    " on  m5.cidvalor01 = m20a.cidvalor01 " +
                    " and m20a.ccodigov01 = '" + aCodigo1 + "'";

                if (aCodigo2 != "")
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
            miconexion.mAbrirConexionComercial(false);
            RegProducto lprod = new RegProducto();

            string lquery = "select ccodigoproducto from mgw10005 where ccodigop01 = '" + codigo + "'";

            SqlCommand lsql = new SqlCommand();
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

        public void mSetDoctos(List<RegDocto> lista)
        {
            _RegDoctos.Clear();
            _RegDoctos = lista;
        }

        public RegCliente mBuscarClienteComercialId(int Id)
        {
            RegCliente lcliente = new RegCliente();


            Boolean lcerrar = false;
            if (miconexion._conexion1 == null)
            {
                if (cadenaconexion != "")
                    miconexion.mAbrirConexionComercial(cadenaconexion, false);
                else
                {
                    miconexion.mAbrirConexionComercial(false);
                    lcerrar = true;
                }
            }
            //lconexion = miconexion._conexion;

            string lquery = "";

            SqlCommand lsql = new SqlCommand();
            SqlDataReader lreader;

            // miconexion.mAbrirConexionDestino();

            lsql.CommandText = "select cidclienteproveedor,ccodigocliente,crazonsocial, cmetodopag,ctextoextra1 from admClientes where cidclienteproveedor = '" + Id + "'";
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
                    lcliente.RazonSocial = lreader[2].ToString();
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



        public RegCliente mBuscarClienteComercial(string codigo)
        {
            RegCliente lcliente = new RegCliente();

            if (codigo == "")
                return lcliente;

            Boolean lcerrar = false;
            if (miconexion._conexion1 == null)
            {
                if (cadenaconexion != "")
                    miconexion.mAbrirConexionComercial(cadenaconexion, false);
                else
                {
                    miconexion.mAbrirConexionComercial(false);
                    lcerrar = true;
                }
            }
            //lconexion = miconexion._conexion;

            string lquery = "";

            SqlCommand lsql = new SqlCommand();
            SqlDataReader lreader;

            // miconexion.mAbrirConexionDestino();

            lsql.CommandText = "select cidclienteproveedor,ccodigocliente,crazonsocial, cmetodopag,ctextoextra1 from admClientes where ccodigocliente = '" + codigo + "'";
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
                    lcliente.RazonSocial = lreader[2].ToString();
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


        public RegCliente mBuscarNumeroSerieComercial(string NumeroSerie, string aCodigoProducto, string aCodigoAlmacen)
        {
            RegCliente lcliente = new RegCliente();

            if (NumeroSerie == "")
                return lcliente;

            Boolean lcerrar = false;
            if (miconexion._conexion1 == null)
            {
                if (cadenaconexion != "")
                    miconexion.mAbrirConexionComercial(cadenaconexion, false);
                else
                {
                    miconexion.mAbrirConexionComercial(false);
                    lcerrar = true;
                }

            }
            else
                if (miconexion._conexion1.State == ConnectionState.Closed)
            {
                miconexion._conexion1.Open();
                lcerrar = true;
            }
            //lconexion = miconexion._conexion;

            string lquery = "";

            SqlCommand lsql = new SqlCommand();
            SqlDataReader lreader;

            // miconexion.mAbrirConexionDestino();

            lsql.CommandText = "select cidserie, CNUMEROSERIE, CPEDIMENTO, caduana, cfechapedimento from admNumerosSerie n " +
            "JOIN admproductos p on n.cidproducto = p.cidproducto and p.ccodigoproducto = '" + aCodigoProducto + "'" +
            "join admalmacenes a on a.cidalmacen = n.cidalmacen and a.ccodigoalmacen = '" + aCodigoAlmacen + "'" +
            "where cnumeroserie = '" + NumeroSerie + "'" +
            "and n.cestado = 1 " +
            " AND not exists " +
            " (select 1 from admMovmientosSeriePedido ms where ms.cidserie = n.CIDSERIE)";


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
                    lcliente.RazonSocial = lreader[2].ToString();
                    //lcliente.MetodoPago = lreader[3].ToString();


                }
                catch (Exception ee)
                {
                    //                    lreader.Close();
                }

            }
            if (lcerrar == true)
            {

                miconexion.mCerrarConexionOrigenComercial();
            }
            lreader.Close();
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
            string lregresa = "";
            string x = "";

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
