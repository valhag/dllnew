using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using LibreriaDoctos;
using System.Runtime.InteropServices;
using Microsoft.Win32;


namespace Controles
{

    
    /*
    Public Declare Function MCC_CreaComponente Lib "CAC100" (ByVal pNombreComponente As String, ByVal pPadre As String) As Long
Public Declare Function MCC_PonPropiedad Lib "CAC100" (ByVal pPtrComponente As Long, ByVal pNombrePropiedad As String, ByVal pValorPropiedad As String, ByVal pPtrCACForma As String) As Long
Public Declare Sub MCC_EjecutaMetodo Lib "CAC100" (ByVal pPtrComponente As Long, ByVal pNombreMetodo As String, ByVal pIdControl As Long, ByVal pPtrTabla As String, ByVal pPtrInformacion As String)
Public Declare Function MCC_PidePropiedad Lib "CAC100" (ByVal pPtrComponente As Long, ByVal pNombrePropiedad As String, ByVal pValorPropiedad As String) As Long
Public Declare Function MCC_DestruyeComponente Lib "CAC100" (ByVal pPtrComponente As Long) As Long
    */

    public partial class codigocatalogo : UserControl
    {

        public event EventHandler PressSearchButton;

        public string llaveregistry = "SOFTWARE\\Computación en Acción, SA CV\\AdminPAQ";
        public string llaveregistrycomercial = "SOFTWARE\\Computación en Acción, SA CV\\CONTPAQ I COMERCIAL";

        [DllImport("KERNEL32.DLL")]
        static extern int SetCurrentDirectory(string pPtrDirActual);
        //[DllImport("CAC100.DLL")]        static extern int MCC_CreaComponente(string pNombreComponente, string pPadre);
        [DllImport("CAC100.DLL")]
        static extern int MCC_PonPropiedad(int pPtrComponente, string pNombrePropiedad, string pValorPropiedad, string pPtrCACForma);

//        ByVal pPtrComponente As Long, ByVal pNombrePropiedad As String, ByVal pValorPropiedad As String, ByVal pPtrCACForma As String
        [DllImport("CAC100.DLL")]
        static extern int MCC_CreaComponente(string pNombreComponente, string pPadre);
        [DllImport("CAC100.DLL")]
        static extern void MCC_EjecutaMetodo(int pPtrComponente, string pNombreMetodo , int pIdControl, string pPtrTabla , string pPtrInformacion);
        [DllImport("CAC100.DLL")]
        static extern int MCC_PidePropiedad(int pPtrComponente, string pNombrePropiedad, StringBuilder pValorPropiedad);

        [DllImport("CAC100.DLL")]
        static extern int MCC_DestruyeComponente(int pPtrComponente);

        ClassRN lrn = new ClassRN();


        public string lrutaempresa;
        private int tipocatalogo;
        private int noClasificacion;

        public codigocatalogo()
        {
            InitializeComponent();
        }

        public string mGetCodigo()
        {
            return textBox1.Text;
        }


        public void cambiardlladmin()
        {
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry);
            Object obc = hklp.GetValue("DIRECTORIOBASE");
            string lruta1 = obc.ToString();
            string lruta2 = @lruta1;
            SetCurrentDirectory(obc.ToString());
        }
        public void mSeteartipo(int aTipo, int anoClasificacion)
        {
            // 1 agentes
            // 2 clientes
            // 3 proveddores
            // 4 productos
            // almacenes
            // 6 clasificacion de productos
            tipocatalogo = aTipo;
            noClasificacion = anoClasificacion;
        }

        public  void mSetLabelText (string text)
        {
            label1.Text = text;

        }

        private void codigocatalogo_Load(object sender, EventArgs e)
        {
            lrn.mSeteaDirectorio(Directory.GetCurrentDirectory());

            button1.Click += new EventHandler(OnButtonClick);

        }

        public void OnButtonClick(object sender, EventArgs e)
        {
            // Delegate the event to the caller
            if (PressSearchButton != null)
                PressSearchButton(this.button1, e);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                RegProducto lreg = new RegProducto();
                lreg = lrn.mBuscarClasificacion(textBox1.Text,noClasificacion, tipocatalogo);
                if (lreg.Id != 0)
                    textBox2.Text = lreg.Nombre;
                else
                    MessageBox.Show ("Clasificacion no existe");

            }
            else
                textBox2.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {



            int lPtrComponente = 0;
            string lNombreComponente = "TDlgSeleccion" + Convert.ToChar(0);
            string lTituloForma = "";
            string lNombreTabla = lrutaempresa;
            string lColumna1 = "";
            string lColumna2 = "";
            string lCampo = "";
            string sLimite = "";
            if (noClasificacion == 0)
                switch (tipocatalogo)
                {
                    case 4:
                        lTituloForma = "Catálogo de Productos" + Convert.ToChar(0);
                        lNombreTabla = lNombreTabla + @"\MGW10005";
                        lCampo = "cIdProducto" + Convert.ToChar(0);
                        lColumna1 = "1,150,Código,cCodigoProducto,,iCodigoTipo" + Convert.ToChar(0);
                        lColumna2 = "2,500,Nombre,cNombreProducto,,iNombreTipo" + Convert.ToChar(0);
                        break;
                }
            else
            {
                lNombreTabla = lNombreTabla + @"\MGW10020";
                lCampo = "cIdValorClasificacion" + Convert.ToChar(0);
                lColumna1 = "1,150,Clasificación,cCodigoValorClasificacion,,icCodigoClasificacion" + Convert.ToChar(0);
                lColumna2 = "2,500,Descripción,cValorClasificacion,," + Convert.ToChar(0);
                /*switch (tipocatalogo)
                {
/*
                    case 4:
                        switch (noClasificacion)
                        {
                            case 1:
                                lTituloForma = "Clasificacion 1 de Productos" + Convert.ToChar(0);
                                sLimite = "25";
                                break;
                            case 2:
                                lTituloForma = "Clasificacion 2 de Productos" + Convert.ToChar(0);
                                sLimite = "26";
                                break;
                            case 3:
                                lTituloForma = "Clasificacion 3 de Productos" + Convert.ToChar(0);
                                sLimite = "27";
                                break;
                            case 4:
                                lTituloForma = "Clasificacion 4 de Productos" + Convert.ToChar(0);
                                sLimite = "28";
                                break;
                            case 5:
                                lTituloForma = "Clasificacion 5 de Productos" + Convert.ToChar(0);
                                sLimite = "29";
                                break;
                            case 6:
                                lTituloForma = "Clasificacion 6 de Productos" + Convert.ToChar(0);
                                sLimite = "30";
                                break;

                        }
                        break;
                }
                    */

                sLimite = noClasificacion.ToString();





                // necesitoo cambiarme al directorio del sdk
                cambiardlladmin();

                lPtrComponente = MCC_CreaComponente(lNombreComponente, null);

                if (lPtrComponente != 0)
                {
                    MCC_PonPropiedad(lPtrComponente, "TituloForma", lTituloForma, null);
                    MCC_PonPropiedad(lPtrComponente, "Tabla", lNombreTabla, null);
                    MCC_PonPropiedad(lPtrComponente, "Campo", lCampo, null);
                    MCC_PonPropiedad(lPtrComponente, "CampoRegreso", "1", null);
                    MCC_PonPropiedad(lPtrComponente, "NumeroColumnas", "2", null);
                    MCC_PonPropiedad(lPtrComponente, "Columna", lColumna1, null);
                    MCC_PonPropiedad(lPtrComponente, "Columna", lColumna2, null);
                    MCC_PonPropiedad(lPtrComponente, "Alias", "", null);

                    if (noClasificacion != 0)
                    {
                        MCC_PonPropiedad(lPtrComponente, "CampoRango", "cIdClasificacion", null);
                        MCC_PonPropiedad(lPtrComponente, "LimiteInferior", sLimite, null);
                        MCC_PonPropiedad(lPtrComponente, "LimiteSuperior", sLimite, null);
                    }

                    MCC_EjecutaMetodo(lPtrComponente, "Ejecuta", 0, null, null);
                    StringBuilder lBuffer = new StringBuilder(256);
                    MCC_PidePropiedad(lPtrComponente, "ValorColumna", lBuffer);
                    MCC_DestruyeComponente(lPtrComponente);
                    textBox1.Text = lBuffer.ToString();

                    if (textBox1.Text != "")
                    {
                        //RegProducto lreg = lrn.mBuscarProducto(textBox1.Text);
                        RegProducto lreg1 = lrn.mBuscarClasificacion(textBox1.Text, int.Parse(sLimite), tipocatalogo);
                        textBox2.Text = lreg1.Nombre;
                    }
                    else
                    { textBox2.Text = ""; }

                    // MessageBox.Show(lBuffer.ToString());
                }
            }
        }
        public void mClean()
        {
            textBox1.Text = "";
            textBox2.Text = "";

        }
        
    }
}
