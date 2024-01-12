using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using LibreriaDoctos;

namespace Controles
{
    public partial class Form1 : Form
    {
        public string mDataSource;
        public List<RegProveedor> lista = new List<RegProveedor>();
        public RegProveedor lregresa = new RegProveedor();

        public Form1()
        {
            InitializeComponent();
        }

        public DialogResult ShowDialog(out RegProveedor result)
        {
            DialogResult dialogResult = base.ShowDialog();

            result = lregresa;
            return dialogResult;
        }

        public Form1(List<RegProveedor> alista, ref RegProveedor aregresa)
        {
            InitializeComponent();
            lista = alista;

        }
        
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            dataGridView1.DataSource = lista;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Visible = false;
            dataGridView1.Columns[3].Visible = false;
            dataGridView1.Columns[4].Visible = false;
            dataGridView1.Columns[2].Width = 500;


        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13 || e.KeyChar == (char)10)
            {
                var result = lista.Where(x => x.RazonSocial.Contains(textBox2.Text)).ToList();
                
                dataGridView1.DataSource = result;


            }

            
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            lregresa = (RegProveedor)dataGridView1.CurrentRow.DataBoundItem;

            this.Close();
            //MessageBox.Show(dataGridView1.CurrentRow.Cells[0].Value.ToString());
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab)
            {
                
            }
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            var result = lista.Where(x => x.RazonSocial.Contains(textBox2.Text)).ToList();

            dataGridView1.DataSource = result;

        }

        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {
            
                
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                //MessageBox.Show(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                lregresa = (RegProveedor)dataGridView1.CurrentRow.DataBoundItem; ;
                this.Close();

            }

        }
    }
}
