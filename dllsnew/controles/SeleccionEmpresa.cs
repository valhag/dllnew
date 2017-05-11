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

namespace InterfazArchivoAdmin
{


    public partial class SeleccionEmpresa : UserControl
    
    {
        public event UserControlClickHandler InnerButtonClick;
        public delegate void UserControlClickHandler(object sender, EventArgs e);

        public event EventHandler SelectedItem;
        

        ClassRN lrn = new ClassRN();


        public string lrutaempresa;

        public SeleccionEmpresa()
        {
            InitializeComponent();
        }

        private int mcargarEmpresa()
        {
            if (DateTime.Today > DateTime.Parse("2013/12/01"))
            {
                //MessageBox.Show ("La configuracion de adminpaq no es correcta");
                //return 1;
            }

            //List<RegEmpresa> x = new List<RegEmpresa>();
            //x = lrn.mCargarEmpresas();
            string mensaje = "";
            this.comboBox1.Items.Clear();
            this.comboBox1.DataSource = lrn.mCargarEmpresas(out mensaje);
            //MessageBox.Show (mensaje);

            //if (mensaje == "")
            //{
            comboBox1.DisplayMember = "Nombre";
            comboBox1.ValueMember = "Ruta";
            comboBox1.Update();
            try
            {
                comboBox1.SelectedIndex = -1;
            }
            catch (Exception ee)
            {
            }
            if (comboBox1.Items.Count == 0)
                return 1;

            comboBox1.SelectedIndex = 0;
            return 0;
            //}
            //else
            //   MessageBox.Show (mensaje);
        }
        private void SeleccionEmpresa_Load(object sender, EventArgs e)
        {
            lrn.mSeteaDirectorio(Directory.GetCurrentDirectory());
            

            if (mcargarEmpresa() == 0)
            {
                string lempresa = "";
                //string lempresa = Properties.Settings.Default.Empresa.Trim();
                //comboBox1.Text = lempresa;

                int i = 0;
                int z = 0;
                foreach (RegEmpresa item in comboBox1.Items)
                {
                    if (item.Nombre.Trim() == lempresa)
                        z = i;
                    else
                        i++;
                }

                comboBox1.SelectedIndex = z;
                comboBox1.SelectedIndexChanged += new EventHandler(OnSelectedItem);
                
                return;
            }
        }

        public void OnSelectedItem(object sender, EventArgs e)
        {
            // Delegate the event to the caller
            if (SelectedItem != null)
                SelectedItem(this.comboBox1, e);
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            lrutaempresa = comboBox1.SelectedValue.ToString().Trim();
            if (this.InnerButtonClick != null)
            {
                this.InnerButtonClick(sender, e);
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
    }
}
