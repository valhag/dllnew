using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Controles
{

    
    public partial class TBDecimales : UserControl
    {
        public event EventHandler TextLeave;
        public event EventHandler TextEnter;
        public event EventHandler TextKeyDown;
        //public delegate void ButtonClickedEventHandler(object sender, EventArgs e);
        //public event ButtonClickedEventHandler OnUserControlButtonClicked;

        public TBDecimales()
        {
            InitializeComponent();
            textBox1.Leave += new EventHandler(OnTextLeave);
            textBox1.KeyDown += new KeyEventHandler(OnTextKeyDown);
        }

        public string mRegresarDecimal()
        {
            return textBox1.Text;
        }

        public void mHabilitar()
        {
            textBox1.ReadOnly = true ;
        }

        public void mSetearDecimal(string aNumero)
        {
            textBox1.Text = aNumero;
        }

        private void OnTextKeyDown(object sender, KeyEventArgs e)
        {
            // Delegate the event to the caller
            if (TextKeyDown != null)
                TextKeyDown(this, e);
        }

        private void OnTextLeave(object sender, EventArgs e)
        {
            // Delegate the event to the caller
            if (TextLeave != null)
                TextLeave(this, e);
        }

        

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 8 || e.KeyChar == 13)
            {
                

                e.Handled = false;
                if (e.KeyChar == 13)
                    SendKeys.Send("{tab}");
                return;
            }
            bool IsDec = false;
            int nroDec = 0;

            for (int i = 0; i < textBox1.Text.Length; i++)
            {
                if (textBox1.Text[i] == '.')
                    IsDec = true;

                if (IsDec && nroDec++ >= 2)
                {
                    e.Handled = true;
                    return;
                }
            }
            if (e.KeyChar >= 48 && e.KeyChar <= 57)
                e.Handled = false;
            else if (e.KeyChar == 46)
                e.Handled = (IsDec) ? true : false;
            else
                e.Handled = true; 
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void TBDecimales_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (this.TextLeave != null)
                this.TextLeave(this, e); 
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
           
        }
    }
}
