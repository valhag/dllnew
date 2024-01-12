using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Data.SqlClient;


namespace Controles
{
    public partial class EmpresasComercial : UserControl
    {
        public ComboBox micombo = new ComboBox();
        public event EventHandler SelectedItem;
        
        string Cadenaconexion = "";
        public string aliasbdd = "";

        public string llaveregistry = "SOFTWARE\\Computación en Acción, SA CV\\CONTPAQ I COMERCIAL";
        [DllImport("KERNEL32.DLL")]static extern int SetCurrentDirectory(string pPtrDirActual);
        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fSetNombrePAQ(string aSistema);
        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fError(int aNumError, string aMensaje, int aLen);
        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fTerminaSDK();
        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fPosPrimerEmpresa(ref int aIdEmpresa, ref string aNombreEmpresa, ref string aDirectorioEmpresa);

        [DllImport("MGWSERVICIOS.DLL")]
        static extern int fAbreEmpresa (string aDirectorioEmpresa );

        [DllImport("MGWSERVICIOS.DLL")]
        static extern void fCierraEmpresa();


        public EmpresasComercial(string aCadena)
        {
            Cadenaconexion = aCadena;
        }

        public EmpresasComercial()
        {
            InitializeComponent();

            comboBox1.SelectedIndexChanged += new EventHandler(OnSelectedItem);
            micombo = comboBox1;
            
        }

        public void OnTextLeave(object sender, EventArgs e)
        {
            MessageBox.Show("Uno");
        }

        public void SetTitulo(string x)
        {
            groupBox2.Text = x;
        }

        public void OnSelectedItem(object sender, EventArgs e)
        {
            // Delegate the event to the caller
            if (SelectedItem != null)
                SelectedItem(this.comboBox1, e);
        }

        public void Populate(string aCadena)
        {
            Cadenaconexion = aCadena;
            DataTable Empresas = null;
            mTraerEmpresas(ref Empresas);
            if (Empresas != null)
            {
                mllenaList(Empresas);
            }
            else
            {
                MessageBox.Show("Es necesario que configure correctamente los datos de la configuracion de la conexion a sqlserver");
            }
        }

        public void PopulateC(string aCadena)
        {
            Cadenaconexion = aCadena;
            DataTable Empresas = null;
            mTraerEmpresasC(ref Empresas);
            if (Empresas != null)
            {
                mllenaListC(Empresas);
            }
            else
            {
                MessageBox.Show("Es necesario que configure correctamente los datos de la configuracion de la conexion a sqlserver");
            }
        }
        private void mllenaList(DataTable Empresas)
        {
            if (comboBox1.Items.Count == 0)
            {
                comboBox1.Items.Clear();
                comboBox1.DataSource = Empresas;
                comboBox1.DisplayMember = "cnombreempresa";
                comboBox1.ValueMember = "crutadatos";
            }

        }

        private void mllenaListC(DataTable Empresas)
        {
            if (comboBox1.Items.Count == 0)
            {
                comboBox1.Items.Clear();
                comboBox1.DataSource = Empresas;
                comboBox1.DisplayMember = "nombre";
                comboBox1.ValueMember = "aliasbdd";
            }

        }
        private void mTraerEmpresasC(ref DataTable Empresas)
        {
            SqlConnection DbConnection = new SqlConnection(Cadenaconexion);


            SqlCommand mySqlCommand = new SqlCommand("select nombre,aliasbdd from ListaEmpresas where nombre != '(Predeterminada)'", DbConnection);
            DataSet ds = new DataSet();
            //mySqlCommand.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
            mySqlDataAdapter.SelectCommand = mySqlCommand;

            try
            {
                mySqlDataAdapter.Fill(ds);
                Empresas = ds.Tables[0];

            }
            catch (Exception ee)
            {

            }
        }
        private void mTraerEmpresas(ref DataTable Empresas)
        {
            SqlConnection DbConnection = new SqlConnection(Cadenaconexion);


            SqlCommand mySqlCommand = new SqlCommand("select cnombreempresa,crutadatos from Empresas where cnombreempresa != '(Predeterminada)'", DbConnection);
            DataSet ds = new DataSet();
            //mySqlCommand.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
            mySqlDataAdapter.SelectCommand = mySqlCommand;

            try
            {
                mySqlDataAdapter.Fill(ds);
                Empresas = ds.Tables[0];

            }
            catch (Exception ee)
            {

            }
        }

        public class ComboboxItem
        {
            public string Text { get; set; }
            public string Value { get; set; }
            public override string ToString() { return Text; }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            ComboBox cmb = (ComboBox)sender;
            if (cmb.SelectedIndex != -1)
            {
                int selectedIndex = cmb.SelectedIndex;


                DataRowView selectedCar = (DataRowView)cmb.SelectedItem;
                aliasbdd = selectedCar.Row[1].ToString();
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void EmpresasComercial_Load(object sender, EventArgs e)
        {
                    }

        private void comboBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // buscar las empresas
            // buscar directorio base
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry);
            Object obc = hklp.GetValue("DIRECTORIOBASE");
            string lruta1 = obc.ToString();
            string lruta2 = @lruta1;
            SetCurrentDirectory(obc.ToString());
            MessageBox.Show(obc.ToString());

            return ;
            string sMensaje = "";
            int lResultado = fSetNombrePAQ("CONTPAQ I Comercial");
            if (lResultado != 0)
            {
                fError(lResultado, sMensaje, 512);
                MessageBox.Show("Error: " + sMensaje);
            }

            

            string aNombreEmpresa ="0000000";
            string aDirectorioEmpresa="0000000000";

            //StringBuilder aNombreEmpresa = new StringBuilder(30);
            //StringBuilder aDirectorioEmpresa = new StringBuilder(30);

            int aIdEmpresa = 0;
            lResultado = fPosPrimerEmpresa(ref aIdEmpresa, ref aNombreEmpresa, ref aDirectorioEmpresa);

            string lDirectorioEmpresa = @"C:\Compac\Empresas\adBIOS2";
            fAbreEmpresa(lDirectorioEmpresa);
            fCierraEmpresa();
            

            fTerminaSDK();

        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;
            if (cmb.SelectedIndex != -1)
            {
                int selectedIndex = cmb.SelectedIndex;


                DataRowView selectedCar = (DataRowView)cmb.SelectedItem;
                aliasbdd = selectedCar.Row[1].ToString();
            }
        }

       
    }
}
