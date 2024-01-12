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

    public partial class codigocatalogocomercial : UserControl
    {

                public event EventHandler PressSearchButton;

        /*       public string llaveregistry = "SOFTWARE\\Computación en Acción, SA CV\\AdminPAQ";
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
               static extern int MCC_DestruyeComponente(int pPtrComponente);*/

        public ClassRN lrn = new ClassRN();

        public ClassConexion con = new ClassConexion();

        public event EventHandler TextLeave;


        public string lrutaempresa;
        private int tipocatalogo;
        private int noClasificacion;

        public RegProveedor lRegClienteProveedor;

        public codigocatalogocomercial()
        {
            InitializeComponent();
            textBox1.Leave += new EventHandler(OnTextLeave);
        }

        public void OnTextLeave(object sender, EventArgs e)
        {
            // Delegate the event to the caller
            if (TextLeave != null)
                TextLeave(this.textBox1, e);
        }

        public string mGetCodigo()
        {
            return textBox1.Text;
        }

        public string mGetNombre()
        {
            return textBox2.Text;
        }

        public void mSetDescripcion(string ltexto)
        {
            textBox2.Text = ltexto;
        }

        public void mSetCodigo(string ltexto)
        {
            textBox1.Text = ltexto;
        }

        public void mSeteartipo(int aTipo, int anoClasificacion=0)
        {
            // 1 agentes
            // 2 clientes
            // 3 proveddores
            // 4 productos
            // 5 almacenes
            // 6 clasificacion de productos
            tipocatalogo = aTipo;
            noClasificacion = anoClasificacion;
        }

        public void mSetLibreria(ClassRN alrn)
        {
            
            lrn = alrn;
        }

        public  void mSetLabelText (string text)
        {
            label1.Text = text;

        }

        /*public void mSetConexion(RegConexion CadenaConexion)
        {

            SqlConnection  micon = lrn.lbd.miconexion.mAbrirConexionComercial(CadenaConexion, false);




        }*/

        public void mSetFocus()
        {
            textBox1.Focus();
        }


        private void codigocatalogo_Load(object sender, EventArgs e)
        {
        //    lrn.mSeteaDirectorio(Directory.GetCurrentDirectory());

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
                //string sEmpresa = EmpresasComercial1.
                //x.mValidarCatalogoComercial(1, textBox1.Text, );
                
                RegCliente lreg = new RegCliente();
                //lreg = lrn.mBuscarClasificacion(textBox1.Text,noClasificacion, tipocatalogo);
               lreg = this.lrn.mBuscarClienteComercial(textBox1.Text);
                if (lreg.Id != 0)
                    textBox2.Text = lreg.RazonSocial;
                else
                    MessageBox.Show ("Cliente no existe");

            }
            else
                textBox2.Text = "";
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<RegProveedor> lista = new List<RegProveedor>();

            //List<RegAlmacen> lista = new List<RegProveedor>();

            if (tipocatalogo == 2)
                lista = lrn.mCargarClientesComercial();
            if (tipocatalogo == 1)
                lista = lrn.mCargarAgentesComercial();
            if (tipocatalogo == 5)
                lista = lrn.mCargarAlmacenesComercial();
            if (tipocatalogo == 4)
                lista = lrn.mCargarProductosComercial();

            RegProveedor lregresa = new RegProveedor();

            Form1 x = new Form1(lista, ref lregresa);

            x.ShowDialog(out lregresa);
            //MessageBox.Show(lregresa.RazonSocial);
            textBox1.Text = lregresa.Codigo;
            textBox2.Text = lregresa.RazonSocial;
            lRegClienteProveedor = lregresa;




        }


        public void mClean()
        {
            textBox1.Text = "";
            textBox2.Text = "";

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        
    }
}
