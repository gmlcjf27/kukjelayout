using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Day_Close_Down
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

        //신한일일
        private void button1_Click(object sender, EventArgs e)
        {
            ButtonClick(1);
        }
        //삼성일일
        private void button2_Click(object sender, EventArgs e)
        {
            ButtonClick(2);
        }
        //하나일일
        private void button3_Click(object sender, EventArgs e)
        {
            ButtonClick(3);
        }
        //롯데일일
        private void button5_Click(object sender, EventArgs e)
        {
            ButtonClick(5);
        }
        //농협
        private void button6_Click(object sender, EventArgs e)
        {
            ButtonClick(6);
        }
        //비씨동의서
        private void button7_Click(object sender, EventArgs e)
        {
            ButtonClick(7);
        }
        //국민
        private void button8_Click(object sender, EventArgs e)
        {
            ButtonClick(8);
        }
        
        //취소
        private void button99_Click(object sender, EventArgs e)
        {
            ButtonClick(0);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ButtonClick(9);
        }
	}
}