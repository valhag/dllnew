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


        public void mCerrarSdkComercial()
        {
            lbd.mCerrarSdkComercial();
        }

        public void mLlenarinfoXML(string archivo)
        {
            lbd.mLlenarinfoXML(archivo);
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

        

        public string mGrabarDoctosComercial(int incluyetimbrado, ref long lultimoFolio)
        {
            return lbd.mGrabarDoctosComercial(incluyetimbrado, ref lultimoFolio);
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

        public List<RegConcepto> mCargarConceptosFacturacfdi()
        {
            return lbd.mCargarConceptos(4, 0,1);
        }

        public List<RegConcepto> mCargarConceptosFacturaComercial()
        {
            return lbd.mCargarConceptosComercial(4, 0);
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
