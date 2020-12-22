using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace NPI_DOWN
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
        private void ButtonClick(int value)
        {
            iSelect = value;
            this.DialogResult = DialogResult.OK;
        }
        //BC_NIP
		private void button1_Click(object sender, EventArgs e) 
        {
			ButtonClick(1);
		}
        //국민_NPI
        private void button2_Click(object sender, EventArgs e)
        {
            ButtonClick(2);
        }        
        //신한_NPI
        private void button3_Click(object sender, EventArgs e)
        {
            ButtonClick(3);
        }
        //하나_NPI
        private void button4_Click(object sender, EventArgs e)
        {
            ButtonClick(4);
        }
        //롯데_NPI
        private void button5_Click(object sender, EventArgs e)
        {
            ButtonClick(5);
        }
        //현대
        private void button6_Click(object sender, EventArgs e)
        {
            ButtonClick(6);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ButtonClick(7);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ButtonClick(8);
        }
        //카카오뱅크
        private void button9_Click(object sender, EventArgs e)
        {
            ButtonClick(9);
        }
        //인터파크
        private void button11_Click(object sender, EventArgs e)
        {
            ButtonClick(10);
        }
        //취소
        private void button10_Click(object sender, EventArgs e)
        {
            ButtonClick(0);
        }
	}
}