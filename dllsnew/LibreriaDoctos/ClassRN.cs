using System;
using System.Collections.Generic;
using System.Text;
using Interfaces ;
using System.Data;
//using BarradeProgreso;


namespace LibreriaDoctos
{
    public class ClassRN
    {
        public string productos = "";
        public string almacenes = "";
        public ClassBD lbd = new ClassBD();


       
        public List<RegConcepto> mCargarConceptosPedidosComercial()
        {
            return lbd.mCargarConceptosComercial(2, 0);
        }

        public List<RegConcepto> mCargarConceptosDevolucionComercial()
        {
            return lbd.mCargarConceptosComercial(5, 0);
        }
        public string mGrabarDoctosComercial(int incluyetimbrado, ref long lultimoFolio, int incluyedireccion, int conComercioExterior = 0, int grabarcliente = 1, int traspaso = 0)
        {
            return lbd.mGrabarDoctosComercial(incluyetimbrado, ref lultimoFolio, incluyedireccion, conComercioExterior, grabarcliente, traspaso);
        }

        public List<RegConcepto> mCargarConceptosCargosComercial()
        {
            return lbd.mCargarConceptosComercial(13, 0, 0);
        }

        
        public string mLlenarinfoAutorizaciones(int liddocumento, string concepto, string doctode)
        {
            return lbd.mLlenarinfoAutorizaciones(liddocumento, concepto,doctode);
        }

        public string mLLenarInfoPedidosFacturas(string archivo)
        {
            return lbd.mLLenarInfoPedidosFacturas(archivo);
        }

        public string mLLenarInfoMontesori(string archivo)
        {
            return lbd.mLLenarInfoMontesori(archivo);
        }

        public string mLLenarInfoAdrianaTraspaso(string archivo)
        {
            return lbd.mLLenarInfoAdrianaTraspaso(archivo);
        }

            public string mLLenarInfoAddendas(string archivo)
        {
            return lbd.mLLenarInfoAddendas(archivo);
        }

        public int mEjecutarComando(string comando, int aclientes, int lporcodigo)
        {
            return lbd.mEjecutarComando(comando, aclientes, lporcodigo);
        }

        public int mEjecutarComando2(string comando, int aclientes, int lporcodigo, string empresa)
        {
            return lbd.mEjecutarComando2(comando, aclientes, lporcodigo, empresa);
        }

        public int mEjecutarComando3(string comando, int aclientes, int lporcodigo, string empresaorigen, string empresadestino)
        {
            return lbd.mEjecutarComando3(comando, aclientes, lporcodigo, empresaorigen, empresadestino);
        }

        public int mProcesarInventarios(string comandoDoctos, string comandoMovtos, int aclientes, int lporcodigo, string empresaorigen, string empresadestino, string comandofolios = "")
        {
            return lbd.mProcesarInventarios(comandoDoctos, comandoMovtos, aclientes, lporcodigo, empresaorigen, empresadestino, comandofolios);
        }
        public int mValidaSQLConexion(string server, string bd, string user, string psw)
        {
            return lbd.mValidaSQLConexion(server, bd, user, psw);
        }

        public void mCerrarSdkComercial()
        {
            lbd.mCerrarSdkComercial();
        }

        public void mBorrarDocto(string Concepto, string Serie, string Folio)
        {
            lbd.mBorrarDocto(Concepto, Serie, Folio);
        }

        public void mAsignaEmpresaComercial(RegConexion empresa)
    {
            lbd.mAsignaEmpresaComercial(empresa);
    }

        public string mLlenarinfoXML(string archivo)
        {
            return lbd.mLlenarinfoXML(archivo);
        }

        public void mLlenarinfoMicroplane(int afolioinicial, int afoliofinal)
        {
            lbd.mLlenarinfoMicroplane(afolioinicial, afoliofinal);
        }

        public List<RegConcepto> mCargarConceptosFacturacfdiComercial()
        {
            return lbd.mCargarConceptosComercial(4, 0, 1);
        }

        public List<RegConcepto> mCargarConceptosDevolucioncfdiComercial()
        {
            return lbd.mCargarConceptosComercial(5, 0, 1);
        }

        public List<RegConcepto> mCargarConceptosPagocfdiComercial()
        {
            return lbd.mCargarConceptosComercial(9, 0, 1);
        }

        public List<RegConcepto> mCargarConceptosNCcfdiComercial()
        {
            return lbd.mCargarConceptosComercial(7, 0, 1);
        }

        public int mLlenarInfoAmcoPedidos(string archivo)
        {
            return lbd.mLlenarinfoAmcoPedidos(archivo);
        }


        public int mLlenarInfoFresko(string archivo)
        {
            return lbd.mLlenarinfoFresko(archivo);
        }

        public string mLlenarTraslado(string archivo)
        {
            return lbd.mLlenarTraslado(archivo);
        }


        public void mLLenarInfoFacturacionMasiva(string archivo)
        {
            lbd.mLlenarinfoFacturacionMasiva(archivo);
        }

        public void mLlenarinfo(string archivo, string Observaciones777, string Observaciones888, string txtObservaciones999, string Referencia, string ObservacionesMov, string refmovto777, string textoextra1777, string refmovto888, string textoextra1888, string refmovto999, string textoextra1999)
        {
            lbd.mLlenarinfo(archivo, Observaciones777, Observaciones888, txtObservaciones999, Referencia, ObservacionesMov,refmovto777, textoextra1777, refmovto888, textoextra1888, refmovto999, textoextra1999);
        }

        public List<string> mGrabarDoctos(bool incluyetimbrado, int tipo)
        {
            return lbd.mGrabarDoctos( incluyetimbrado, tipo);
        }

        public List<string> mGrabarDoctosFresko(bool incluyetimbrado, int tipo)
        {
            return lbd.mGrabarDoctosFresko(incluyetimbrado, tipo);
        }

        public List<string> mGrabarDoctosTraslado(bool incluyetimbrado, int tipo)
        {
            return lbd.mGrabarDoctosTraslado(incluyetimbrado, tipo);
        }

        public void mTraerInformacionPrimerReporte(ref DataSet PorConcepto, DateTime ini, DateTime fin)
        {
                lbd.mTraerInformacionPrimerReporte(ref PorConcepto, ini, fin);
        }

        public string mGrabarTablaAdicional()
        {
            return lbd.mGrabarTablaAdicional();
        }

        public string mGrabarAbono(string lConcepto, int lDocumentoModelo)
        {
            return lbd.mGrabarAbono(lConcepto,lDocumentoModelo);
        }

        public List<RegOrigen> mCargarDocumentos(int aDocumentoModelo, int aFolio, string aSerie)
        {
            return lbd.mCargarDocumentos(aDocumentoModelo,aFolio, aSerie);
        }

        public List<RegOrigen> mCargarDocumentosComercialDoctoDeCliente(int aDocumentoModelo, long aIdCliente)
        {
            return lbd.mCargarDocumentosComercialDoctoDeCliente(aDocumentoModelo, aIdCliente);
        }

        public List<RegDocto> mCargarDocumentosComercialReferencia(string aConcepto, string aReferencia, ref DataTable dt, ref DataTable dt2, int pt = 0)
        {
            return lbd.mCargarDocumentosComercialReferencia(aConcepto, aReferencia, ref dt, ref dt2, pt);
        }


        public void mInicializar(string aRutaOrigen, string aRutaDestino)
        {
            //Properties.Settings.Default.RutaEmpresaSamira = aRutaOrigen; 
        }
        //public ClassBD lbd = new ClassBD();
        //public ClassBD lbd;
        public Boolean mBuscar(long aFolio, string aConcepto, string aSerie, int aTipo)
        {
            return lbd.mBuscar(aFolio, aConcepto, aSerie, aTipo);
        }


        public RegDocto mBuscarDoctoComercialProduccion(string aFolio, string Concepto, int porcentaje)
        {
            return lbd.mBuscarDoctoComercialProduccion(aFolio, Concepto,porcentaje);
        }

        public List<RegCliente> mCargarSeriesPedidosComercial(long lidmovimiento)
        {
            return lbd.mCargarSeriesPedidosComercial(lidmovimiento);
        }
        public RegDocto mBuscarDoctoComercial(string aFolio, string aSerie, string Concepto)
        {
            return lbd.mBuscarDoctoComercial(aFolio, aSerie, Concepto);

        }

        public string mBuscarDocto(string aFolio, int aTipo, bool aRevisar)
        {
           return lbd.mBuscarDocto(aFolio,  aTipo, aRevisar );
        }

        public   virtual  string mBuscarDoctoFlex(string aFolio, int aTipo, bool aRevisar)
        {
            return lbd.mBuscarDoctoAccess(aRevisar);
        }

        public virtual string mBuscarDoctos(long aFolioinicial, long afoliofinal , int aTipo, bool aRevisar)
        {
            return lbd.mBuscarDoctos(aFolioinicial, afoliofinal , aTipo, aRevisar);
        }

        public Boolean mValidarConexionIntell(string aRuta)
        {
            return lbd.mValidarConexionIntell(aRuta);
        }

        public Boolean mValidarConexionIntell(string aServidor, string aBd, string ausu, string apwd)
        {
            return lbd.mValidarConexionIntell(aServidor, aBd, ausu, apwd);
        }

        public string mGrabarAdm(string afolioant, double afolionuevo, int opcion, int tipo)
        {
            return lbd.mGrabarAdm(afolioant, afolionuevo , opcion, tipo);
        }

        public List<string> mGrabarAdms(int opcion, int tipo)
        {
            lbd.primerdocto = new RegDocto ();
            return lbd.mGrabarAdms(opcion, tipo);
        }

        public string mBuscarDescripcionProducto(string descripcion)
        {
            return lbd.mBuscarDescripcionProducto(descripcion);
        }

        public RegProducto mBuscarProductoComercial(string codigo)
        {
            return lbd.mBuscarProductoComercial(codigo);
        }

        public RegProducto mBuscarProducto(string codigo)
        {
            return lbd.mBuscarProducto(codigo);
        }

        public string mGrabarmGrabarComplemento(string codigoini, string codigofin)
        {
            return lbd.mGrabarmGrabarComplemento(codigoini, codigofin);
        }

        public RegCliente mBuscarClienteComercial(string codigo)
        {
            return lbd.mBuscarClienteComercial(codigo);
        }

        public RegCliente mBuscarSerieComercial(string Serie, string aCodigoProducto, string aCodigoAlmacen)
        {
            return lbd.mBuscarNumeroSerieComercial(Serie, aCodigoProducto, aCodigoAlmacen);
        }

        public void mSetDoctos(List<RegDocto> lista)
        {
            lbd.mSetDoctos(lista);
        }

        public RegCliente mBuscarClienteComercialId(int Id)
        {
            return lbd.mBuscarClienteComercialId(Id);
        }

        public RegProducto mBuscarClasificacion(string codigo, int anumClasif, int aTipoCatalogo)
        {
            return lbd.mBuscarClasificacion(codigo, anumClasif, aTipoCatalogo);
        }

        public List<RegProducto> mMostrarProductos(string aCodigo1, string aCodigo2,string aCodigo3,string aCodigo4,string aCodigo5,string aCodigo6)
        {
            return lbd.mMostrarProductos( aCodigo1, aCodigo2,aCodigo3,aCodigo4,aCodigo5,aCodigo6);
        }

        public void mGrabarComplemento(List<RegProducto> lista, string aValor1, string aValor2, string aValor3, string aValor4)
    {
        lbd.mGrabarComplemento(lista, aValor1, aValor2, aValor3, aValor4);
    }

        /*
        public string mGrabarDoctosComercial(List<RegDocto> Doctos, string usu, string pass)
        {
            return lbd.mGrabarDoctosComercial44(Doctos, usu, pass);
        }
         * */

        public string mGrabarDoctosComercialBorrar(List<RegDocto> Doctos,  ref int lexitosos, ref int ltotales, int condireccion)
        {
            return lbd.mGrabarDoctosComercialborrar(Doctos, ref lexitosos, ref ltotales, condireccion);
        }

        

        public string mGrabarDoctosComercial(int incluyetimbrado, ref long lultimoFolio, int incluyedireccion, int conComercioExterior = 0)
        {
            return lbd.mGrabarDoctosComercial(incluyetimbrado, ref lultimoFolio, incluyedireccion, conComercioExterior);
        }

        public string mGrabarDestinos( )
        {
            string lregresa = "";
            lregresa =lbd.mGrabarDestinos();
            productos = lbd.productos;
            almacenes = lbd.almacenes ;
            return lregresa;
        }

        public List<RegConcepto> mCargarConceptosFactura()
        {
            return lbd.mCargarConceptos(4,0,0);
        }

        public List<RegConcepto> mCargarConceptosCartaPorte()
        {
            return lbd.mCargarConceptos(4, 0, 0,1);
        }


      

        public List<RegConcepto> mCargarConceptosFacturacfdi()
        {
            return lbd.mCargarConceptos(4, 0,1);
        }

        public List<RegConcepto> mCargarConceptosFacturaComercial()
        {
            return lbd.mCargarConceptosComercial(4, 0);
        }

        public List<RegConcepto> mCargarConceptosEntradaComercial()
        {
            return lbd.mCargarConceptosComercial(32, 0);
        }

        public List<RegConcepto> mCargarConceptosFacturaComercial(int a)
        {
            return lbd.mCargarConceptosComercial(4, 0,a);
        }

        public List<RegConcepto> mCargarConceptosCompraComercial()
        {
            return lbd.mCargarConceptosComercial(19, 0);
        }

        public RegAlmacen mBuscarAlmacenAsumidoComercial(string aCodigoConcepto)
        {
            return lbd.mBuscarAlmacenAsumidoComercial(aCodigoConcepto);
        }

        public List<RegProveedor> mCargarClientes()
        {
            return lbd.mCargarClientes ();
        }

        public decimal mSaldoClienteComercial(long lIdCliente)
        {

            
            return lbd.mSaldoClienteComercial(lIdCliente);


        }


        public List<RegProveedor> mCargarClientesComercial()
        {
            return lbd.mCargarClientesComercial();
        }

        public List<RegProveedor> mCargarProveedoresComercial()
        {
            return lbd.mCargarProveedoresComercial();
        }
        public List<RegProveedor> mCargarAgentesComercial()
        {
            return lbd.mCargarAgentesComercial();
        }

        public List<RegProveedor> mCargarProductosComercial()
        {
            return lbd.mCargarProductosComercial();
        }

        public List<RegProveedor> mCargarAlmacenesComercial()
        {
            return lbd.mCargarAlmacenesComercial();
        }

        public List<RegAlmacen> mCargarAlmacenesComercialv2()
        {
            return lbd.mCargarAlmacenesComercialv2();
        }
        public List<RegConcepto> mCargarConceptosPedido()
        {
            return lbd.mCargarConceptos(2, 0,0);
        }

        public List<RegConcepto> mCargarConceptosDevolucion()
        {
            return lbd.mCargarConceptos(5,0,0);
        }
        public List<RegConcepto> mCargarConceptosNotaCredito()
        {
            return lbd.mCargarConceptos(7, 0,0);
        }
        public List<RegConcepto> mCargarConceptosNotaCargo()
        {
            return lbd.mCargarConceptos(13, 0,0);
        }

        public List<RegConcepto> mCargarConceptosCompraOrigen()
        {
            return lbd.mCargarConceptos(19, 0,0);
        }
        public RegProveedor mBuscarCliente(string aCliente)
        {
            return lbd.mBuscarCliente(aCliente,0,0);
        }
        public RegProveedor mBuscarProveedor(string aProveedor)
        {
            return lbd.mBuscarCliente(aProveedor, 1, 1);
        }

        public  void mSeteaDirectorio(string aRuta)
        {
            lbd.mAsignaRuta( aRuta);
        }

        public void mCargaCom()
        {
            lbd.mCargaCom();
        }
        public void liberarrecursos()
        {
            lbd.liberarrecursos();
        }
        public List<RegEmpresa> mCargarEmpresas(out string mensaje)
        {
            return lbd.mCargarEmpresas(out mensaje);
        }

        public List<RegPuntodeVenta> mCargarPuntoVenta(string aEmpresa, out string mensaje)
        {
            return lbd.mCargarPuntoVenta( aEmpresa, out mensaje);
        }

        public List<RegEmpresas> mCargarEmpresasAccess(out string mensaje)
        {
            return lbd.mCargarEmpresasAccess(out mensaje);
        }

        public virtual string mBuscarDoctosArchivo(string aArchivo)
        {
            return "";
        }
}
}
