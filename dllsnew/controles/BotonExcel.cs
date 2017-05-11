using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
// testing

namespace controles
{
    public partial class BotonExcel : UserControl
    {
        public int tipo;
        public BotonExcel()
        {
            InitializeComponent();
        }
        public string mRegresarNombre()
        {
            return textBox1.Text;
        }
        public void mSetearEtiqueta(string aetiqueta)
        {
            this.label1.Text = aetiqueta;
        }
        public void mGeneraNombre(int atipo)
        {
            string lnombre = Directory.GetCurrentDirectory();
            if (atipo==1) // bitacora
                this.textBox1.Text = lnombre + "\\bitacora.csv";
        }

        public void mGeneraNombre(int atipo, string lNombreArchivo)
        {
            string lnombre = Directory.GetCurrentDirectory();
            if (atipo == 1) // bitacora
                this.textBox1.Text = lnombre + "\\" + lNombreArchivo + ".csv";
            if (atipo == 2) // excel
                this.textBox1.Text = lnombre + "\\" + lNombreArchivo + ".xls";

        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fbDialog;
            fbDialog = new System.Windows.Forms.OpenFileDialog();
            switch(tipo)
            {
                case (0):
                    fbDialog.DefaultExt = "xlsx";
                    fbDialog.Filter = "Excel documents (*.xlsx)|*.xlsx";
                    break;
                case (1):
                    fbDialog.DefaultExt = "txt";
                    fbDialog.Filter = "txt files (*.txt)|*.txt";

                    //fbDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    break;
            }
            if (fbDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                   textBox1.Text = fbDialog.FileName; 
            }
        }

        public void mAsignaTipo(int aTipo)
        {
            tipo = aTipo;
        }
        private void BotonExcel_Load(object sender, EventArgs e)
        {

        }
    }
}
