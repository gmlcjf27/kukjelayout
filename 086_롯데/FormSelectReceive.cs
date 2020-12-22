using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace _086_롯데
{
    public partial class FormSelectReceive : Form
    {
       
        private int iSelect = 0;

        public FormSelectReceive()
        {
            InitializeComponent();
        }

        public int GetSelected
        {
            get
            {
                return iSelect;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            ButtonClick(1);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ButtonClick(2);
        }
        //VVIP-서울
        private void button3_Click(object sender, EventArgs e)
        {
            ButtonClick(3);
        }
        //VVIP-서울외
        private void button4_Click(object sender, EventArgs e)
        {
            ButtonClick(4);
        }

        private void ButtonClick(int value)
        {
            iSelect = value;
            this.DialogResult = DialogResult.OK;
        }

        
    }
}