using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace LibreriaDoctos
{


    public class RegElemento
    {
        public string id;
        public string school_id;
        public string code;
        public string amount;
        public string type;
        public string unit;
        public string date;
    }

    public struct constantes
    {
        public const int kLongFecha = 24;
        public const int kLongSerie = 12;
        public const int kLongCodigo = 31;
        public const int kLongNombre = 61;
        public const int kLongReferencia = 21;
        public const int kLongDescripcion = 61;
        public const int kLongCuenta = 101;
        public const int kLongMensaje = 3001;
        public const int kLongNombreProducto = 256;
        public const int kLongAbreviatura = 4;
        public const int kLongCodValorClasif = 4;

    }
    
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 4)]
    public struct TDocumento
    {
    

        public Double aFolio;
        public int aNumMoneda;
        public Double aTipoCambio;
        public Double aImporte;
        public Double aDescuentoDoc1;
        public Double aDescuentoDoc2;
        public int aSistemaOrigen;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = constantes.kLongCodigo)]
        public String aCodConcepto;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = constantes.kLongSerie)]
        public String aSerie;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = constantes.kLongFecha)]
        public String aFecha;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = constantes.kLongCodigo)]
        public String aCodigoCteProv;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = constantes.kLongCodigo)]
        public String aCodigoAgente;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = constantes.kLongReferencia)]
        public String aReferencia;
        public int aAfecta;
        public int aGasto1;
        public int aGasto2;
        public int aGasto3;
    }
    

    public class tttTDocumento1
    {
    
    
      public double aFolio;
      public int aNumMoneda;
      public double aTipoCambio;
      public double aImporte;
      public double aDescuentoDoc1;
      public double aDescuentoDoc2;
      public int aSistemaOrigen;
      //public fixed char aCodConcepto[31];
      public StringBuilder aCodConcepto = new StringBuilder(31);
      public string aSerie;
      public string aFecha;
      public string aCodigoCteProv;
      public string aCodigoAgente;
      public string aReferencia;
      public int aAfecta;
      public double aGasto1;
      public double aGasto2 ;
      public double aGasto3 ;
      
    }

    public class RegEmpresas
    {
        private string _Empresa;
        public string cEmpresa
        {
            get { return _Empresa; }
            set { _Empresa = value; }
        }
    }

    public class RegPuntodeVenta
    {
        private string _Empresa;
        private string _Nombre;

        public string cEmpresa
        {
            get { return _Empresa; }
            set { _Empresa = value; }
        }
        public string cNombre
        {
            get { return _Nombre; }
            set { _Nombre = value; }
        }
    }


    public class RegDocto
    {
        public List<RegMovto> _RegMovtos = new List<RegMovto>();
        public RegDireccion _RegDireccion = new RegDireccion ();
        public RegCliente _RegCliente = new RegCliente();
        private long _cIdDocto;
        private string _cCodigoCliente="";
        private string _cCodigoConcepto="";
        public string _cNombreConcepto = "";
        private long _cIdConcepto = 0;
        private string _cRFC = "";
        private string _cRazonSocial = "";
        private string _cMoneda = "";
        private string _cCond = "";
        private string _cTextoExtra1="";
        private string _cTextoExtra2 = "";
        private string _cTextoExtra3 = "";
        private string _cNombreArchivo = "";
        private DateTime _cFechaVencimiento;
        private string _cReferencia="";
        public  string _cObservaciones = "";
        private string _cMetodoPago = "";


        private string _cRegimenFiscal = "";
        public string cRegimenFiscal
        {
            get { return _cRegimenFiscal; }
            set { _cRegimenFiscal = value; }
        }

        
        private int _cContado = 0;
        public int cContado
        {
            get { return _cContado; }
            set { _cContado = value; }
        }

        public string cMetodoPago
        {
            get { return _cMetodoPago; }
            set { _cMetodoPago = value; }
        }

        public DateTime cFechaVencimiento
        {
            get { return _cFechaVencimiento; }
            set { _cFechaVencimiento = value; }
        }

        public string cNombreArchivo
        {
            get { return _cNombreArchivo; }
            set { _cNombreArchivo = value; }
        }
        

        public string cTextoExtra1
        {
            get { return _cTextoExtra1; }
            set { _cTextoExtra1 = value; }
        }

        public string cTextoExtra2
        {
            get { return _cTextoExtra2; }
            set { _cTextoExtra2 = value; }
        }

        public string cTextoExtra3
        {
            get { return _cTextoExtra3; }
            set { _cTextoExtra3 = value; }
        }
        private string _sMensaje="";




        public string cReferencia
        {
            get { return _cReferencia; }
            set { _cReferencia = value; }
        }


        public string sMensaje
        {
            get { return _sMensaje; }
            set { _sMensaje = value; }
        }

        public string cCond
        {
            get { return _cCond; }
            set { _cCond = value; }
        }
        private string _cAgente="";

        public string cAgente
        {
            get { return _cAgente; }
            set { _cAgente = value; }
        }



        private double _cNeto=0;

        public double cNeto
        {
            get { return _cNeto; }
            set { _cNeto = value; }
        }
        private double _cImpuestos=0;

        public double cImpuestos
        {
            get { return _cImpuestos; }
            set { _cImpuestos = value; }
        }

        private double _cImpuesto2=0;

        public double cImpuesto2
        {
            get { return _cImpuesto2; }
            set { _cImpuesto2 = value; }
        }

        public string cMoneda
        {
            get { return _cMoneda; }
            set { _cMoneda = value; }
        }
        private decimal _cTipoCambio;

        public decimal cTipoCambio
        {
            get { return _cTipoCambio; }
            set { _cTipoCambio = value; }
        }

        public string cRazonSocial
        {
            get { return _cRazonSocial; }
            set { _cRazonSocial = value; }
        }

        public string cRFC
        {
            get { return _cRFC; }
            set { _cRFC = value; }
        }


        public long cIdConcepto
        {
            get { return _cIdConcepto; }
            set { _cIdConcepto = value; }
        }


        public string cCodigoConcepto
        {
            get { return _cCodigoConcepto; }
            set { _cCodigoConcepto = value; }
        }
        private DateTime _cFecha;

        public DateTime cFecha
        {
            get { return _cFecha; }
            set { _cFecha = value; }
        }
        private long _cFolio=0;

        public long cFolio
        {
            get { return _cFolio; }
            set { _cFolio = value; }
        }


        public long cIdDocto
        {
            get { return _cIdDocto; }
            set { _cIdDocto = value; }
        }
        public string cCodigoCliente
        {
            get { return _cCodigoCliente; }
            set { _cCodigoCliente = value; }
        }

        private string _cSerie="";

        public string cSerie
        {
            get { return _cSerie; }
            set { _cSerie = value; }
        }




    }
    public class RegMovto
    {

        public RegCapa _RegCapa = new RegCapa();
        public RegProducto _RegProducto = new RegProducto();
        public string cObservaciones { get; set; }
        public string cReferencia { get; set; }
        public string ctextoextra1 { get; set; }
        public string ctextoextra2 { get; set; }
        public int procesado { get; set; }
        private string _ctextoextra3=""; 
        private string _cUnidad ="";

        private string _cAlmacenEntrada = "";
        public string cAlmacenEntrada
        {
            get { return _cAlmacenEntrada; }
            set { _cAlmacenEntrada = value; }
        }

        public string ctextoextra3
        {
            get { return _ctextoextra3; }
            set { _ctextoextra3 = value; }
        }

        
        public string cUnidad
        {
            get { return _cUnidad; }
            set { _cUnidad = value; }
        }

        
        private decimal _cMargenUtilidad=0;

        public decimal cMargenUtilidad
        {
            get { return _cMargenUtilidad; }
            set { _cMargenUtilidad = value; }
        }

        private string _cCodigoAlmacen = "";

        public string cCodigoAlmacen
        {
            get { return _cCodigoAlmacen; }
            set { _cCodigoAlmacen = value; }
        }

        private string _cNombreAlmacen = "";

        public string cNombreAlmacen
        {
            get { return _cNombreAlmacen; }
            set { _cNombreAlmacen = value; }
        }

        private long _cIdMovto=0;

        public long cIdMovto
        {
            get { return _cIdMovto; }
            set { _cIdMovto = value; }
        }
        private long _cIdDocto=0;

        public long cIdDocto
        {
            get { return _cIdDocto; }
            set { _cIdDocto = value; }
        }
        private string _cNombreProducto = "";

        public string cNombreProducto
        {
            get { return _cNombreProducto; }
            set { _cNombreProducto = value; }
        }

        private string _cCodigoProducto = "";

        public string cCodigoProducto
        {
            get { return _cCodigoProducto; }
            set { _cCodigoProducto = value; }
        }
        private decimal _cUnidades=0;

        public decimal cUnidades
        {
            get { return _cUnidades; }
            set { _cUnidades = value; }
        }
        
        private decimal _cPrecio=0;

        public decimal cPrecio
        {
            get { return _cPrecio; }
            set { _cPrecio = value; }
        }
        
        private decimal _cSubtotal=0;

        public decimal cSubtotal
        {
            get { return _cSubtotal; }
            set { _cSubtotal = value; }
        }
        private decimal _cTotal=0;

        public decimal cTotal
        {
            get { return _cTotal; }
            set { _cTotal = value; }
        }
        private decimal _cImpuesto;

        public decimal cImpuesto
        {
            get { return _cImpuesto; }
            set { _cImpuesto = value; }
        }

        private decimal _cImpuesto2=0;

        public decimal cImpuesto2
        {
            get { return _cImpuesto2; }
            set { _cImpuesto2 = value; }
        }

        private decimal _cPorcent01=0;
        public decimal cPorcent01
        {
            get { return _cPorcent01; }
            set { _cPorcent01 = value; }
        }

        private decimal _cneto=0;
        public decimal cneto
        {
            get { return _cneto; }
            set { _cneto = value; }
        }



    }

    public class RegCliente
    {
        private long _Id;

        public long Id
        {
            get { return _Id; }
            set { _Id = value; }
        }

        private string _Codigo;

        public string Codigo
        {
            get { return _Codigo; }
            set { _Codigo = value; }
        }
        private string _RazonSocial;

        public string RazonSocial
        {
            get { return _RazonSocial; }
            set { _RazonSocial = value; }
        }
        private string _RFC;

        public string RFC
        {
            get { return _RFC; }
            set { _RFC = value; }
        }
        private int _DiasCredito;

        public int DiasCredito
        {
            get { return _DiasCredito; }
            set { _DiasCredito = value; }
        }


        private string _MetodoPago;

        public string MetodoPago
        {
            get { return _MetodoPago; }
            set { _MetodoPago = value; }
        }


    }


    public class RegProveedor
    {
        private long _Id;

        public long Id
        {
            get { return _Id; }
            set { _Id = value; }
        }

        private string _Codigo;

        public string Codigo
        {
            get { return _Codigo; }
            set { _Codigo = value; }
        }
        private string _RazonSocial;

        public string RazonSocial
        {
            get { return _RazonSocial; }
            set { _RazonSocial = value; }
        }
        private string _RFC;

        public string RFC
        {
            get { return _RFC; }
            set { _RFC = value; }
        }
        private int _DiasCredito;

        public int DiasCredito
        {
            get { return _DiasCredito; }
            set { _DiasCredito = value; }
        }
        


    }

    public class RegCapa
    {

        private string _Pedimento ="";
        public RegAlmacen _Almacen = new RegAlmacen();

        public string Pedimento
        {
            get { return _Pedimento; }
            set { _Pedimento = value; }
        }

        private string _NoAduana;

        public string NoAduana
        {
            get { return _NoAduana; }
            set { _NoAduana = value; }
        }

        private decimal _Unidades;
        public decimal Unidades
        {
            get { return _Unidades; }
            set { _Unidades = value; }
        }
        private decimal _tc;
        public decimal tc
        {
            get { return _tc; }
            set { _tc = value; }
        }
        private DateTime _FechaFabricacion;
        public DateTime FechaFabricacion
        {
            get { return _FechaFabricacion; }
            set { _FechaFabricacion = value; }
        }
    }

    public class RegProducto
    {

        private double _Precio;


        private string _noIdentificacion;

        public string noIdentificacion
        {
            get { return _noIdentificacion; }
            set { _noIdentificacion = value; }
        }
        public double Precio
        {
            get { return _Precio; }
            set { _Precio = value; }
        }

        private long _Id;

        public long Id
        {
            get { return _Id; }
            set { _Id = value; }
        }

        private string _Codigo;

        public string Codigo
        {
            get { return _Codigo; }
            set { _Codigo = value; }
        }
        private string _Nombre;

        public string Nombre
        {
            get { return _Nombre; }
            set { _Nombre = value; }
        }
        private decimal _ImporteExtra1;
        public decimal ImporteExtra1
        {
            get { return _ImporteExtra1; }
            set { _ImporteExtra1 = value; }
        }
    }

    public class RegAlmacen
    {


        private long _Id;

        public long Id
        {
            get { return _Id; }
            set { _Id = value; }
        }

        private string _Codigo;

        public string Codigo
        {
            get { return _Codigo; }
            set { _Codigo = value; }
        }
        private string _Nombre;

        public string Nombre
        {
            get { return _Nombre; }
            set { _Nombre = value; }
        }
    }

    public class RegConcepto
    {
        private string _Codigo;

        public string Codigo
        {
            get { return _Codigo; }
            set { _Codigo = value; }
        }
        private string _Nombre;

        public string Nombre
        {
            get { return _Nombre; }
            set { _Nombre = value; }
        }
        private string _sTipocfd;

        public string Tipocfd
        {
            get { return _sTipocfd; }
            set { _sTipocfd = value; }
        }
    }

    public class RegEmpresa
    {
        public string _Nombre;

        public string Nombre
        {
            get { return _Nombre; }
            set { _Nombre = value; }
        }
        private string _Ruta;

        public string Ruta
        {
            get { return _Ruta; }
            set { _Ruta = value; }
        }
    }

    public class RegDireccion
    {
        private string _cEmail = "";
        private string _cEmail2 = "";

        public string cEmail
        {
            get { return _cEmail; }
            set { _cEmail = value; }
        }
        public string cEmail2
        {
            get { return _cEmail2; }
            set { _cEmail2 = value; }
        }
        private string _cNombreCalle;

        public string cNombreCalle
        {
            get { return _cNombreCalle; }
            set { _cNombreCalle = value; }
        }
        private string _cNumeroExterior;

        public string cNumeroExterior
        {
            get { return _cNumeroExterior; }
            set { _cNumeroExterior = value; }
        }
        private string _cNumeroInterior;

        public string cNumeroInterior
        {
            get { return _cNumeroInterior; }
            set { _cNumeroInterior = value; }
        }
        private string _cColonia;

        public string cColonia
        {
            get { return _cColonia; }
            set { _cColonia = value; }
        }
        private string _cCodigoPostal;

        public string cCodigoPostal
        {
            get { return _cCodigoPostal; }
            set { _cCodigoPostal = value; }
        }
        private string _cEstado;

        public string cEstado
        {
            get { return _cEstado; }
            set { _cEstado = value; }
        }
        private string _cPais;

        public string cPais
        {
            get { return _cPais; }
            set { _cPais = value; }
        }
        private string _cCiudad;

        public string cCiudad
        {
            get { return _cCiudad; }
            set { _cCiudad = value; }
        }
    }
    

    public class RegOrigen
    {

        private int _cidproducto;
        private string _Folio;
        private string _Fecha;
        private string _CodigoProducto;
        private string _NombreProducto;
        private decimal _Cantidad;
        private decimal _Precio;
        private decimal _Precio2;
        private decimal _IEPS;
        private decimal _IEPS2;
        private decimal _Descuento;
        private long _cIdClien01;
        private decimal _cTotal;
        private string _RazonSocial;
        private string _CodigoCliente;
        private string _RFC;
        private decimal _TotalMov;
        //private decimal _TotalMov2;
        private decimal _DescuentoAplicar;
        
        

        #region decl
        public int cidproducto
        {
            get { return _cidproducto; }
            set { _cidproducto = value; }
        }



        

        public string Folio
        {
            get { return _Folio; }
            set { _Folio = value; }
        }


        public string RFC
        {
            get { return _RFC; }
            set { _RFC = value; }
        }
        
        public string Fecha
        {
            get { return _Fecha; }
            set { _Fecha = value; }
        }
        

        public string CodigoProducto
        {
            get { return _CodigoProducto; }
            set { _CodigoProducto = value; }
        }

        

        
        public string NombreProducto
        {
            get { return _NombreProducto; }
            set { _NombreProducto = value; }
        }
        

        public decimal Cantidad
        {
            get { return _Cantidad; }
            set { _Cantidad = value; }
        }

        

        public decimal Precio
        {
            get { return _Precio; }
            set { _Precio = value; }
        }

        public decimal Precio2
        {
            get { return _Precio2; }
            set { _Precio2 = value; }
        }
        

        public decimal IEPS
        {
            get { return _IEPS; }
            set { _IEPS = value; }
        }
        public decimal IEPS2
        {
            get { return _IEPS2; }
            set { _IEPS2 = value; }
        }

        public decimal Descuento
        {
            get { return _Descuento; }
            set { _Descuento = value; }
        }

        

        public long cIdClien01
        {
            get { return _cIdClien01; }
            set { _cIdClien01 = value; }
        }

        public decimal TotalMov
        {
            get { return _TotalMov; }
            set { _TotalMov = value; }
        }

        public decimal cTotal
        {
            get { return _cTotal; }
            set { _cTotal = value; }
        }

        
        public string RazonSocial
        {
            get { return _RazonSocial; }
            set { _RazonSocial = value; }
        }
        

        public string CodigoCliente
        {
            get { return _CodigoCliente; }
            set { _CodigoCliente = value; }
        }
        //public decimal TotalMov2
        //{
        //    get { return _TotalMov2; }
        //    set { _TotalMov2 = value; }
        //}
        public decimal DescuentoAplicar
        {
            get { return _DescuentoAplicar; }
            set { _DescuentoAplicar = value; }
        }

        #endregion decl
    }

}
