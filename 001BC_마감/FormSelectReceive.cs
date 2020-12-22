using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace _001_BC_Close
{
	public partial class FormSelectReceive : Form
	{
		private int iSelect = 0;

		public FormSelectReceive() {
			InitializeComponent();
		}

		public int GetSelected {
			get {
				return iSelect;
			}
		}
        private void ButtonClick(int value)
        {
            iSelect = value;
            this.DialogResult = DialogResult.OK;
        }

        //결과자료
		private void button1_Click(object sender, EventArgs e) {
			ButtonClick(1);
		}

        //개시자료
        private void button2_Click(object sender, EventArgs e)
        {
            ButtonClick(2);
        }
        //제일은행
        private void button3_Click(object sender, EventArgs e)
        {
            ButtonClick(3);
        }
        //농협월말
        private void button4_Click(object sender, EventArgs e)
        {
            ButtonClick(4);
        }
        //수령지변경
        private void button5_Click(object sender, EventArgs e)
        {
            ButtonClick(5);
        }       
        //동의서입고
        private void button7_Click(object sender, EventArgs e)
        {
            ButtonClick(6);
        }
        //동의서등기번호
        private void button8_Click(object sender, EventArgs e)
        {
            ButtonClick(7);
        }
        //취소
        private void button6_Click(object sender, EventArgs e)
        {
            ButtonClick(0);
        }

        
	}
}