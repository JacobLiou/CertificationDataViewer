using Sunny.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sofar.CertificationDataViewer
{
    public partial class InsertForm : Form
    {
        public InsertForm()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {

            this.DialogResult = DialogResult.OK;
            this.Close();
        }


        public string[] GetInsertedData()
        {
            return new string[] { uiTextBox1.Text, uiTextBox2.Text, uiTextBox3.Text,
                uiTextBox4.Text, uiTextBox5.Text, uiTextBox6.Text, uiTextBox7.Text,
                uiTextBox8.Text, uiTextBox9.Text ,uiTextBox10.Text,uiTextBox11.Text,
                uiTextBox12.Text, uiTextBox13.Text};
        }

        public void GetSelectedData(string[] data)
        {

            for (int i = 0; i < data.Length && i < 13; i++)
            {
                var textBox = this.Controls.Find($"uiTextBox{i + 1}", true).FirstOrDefault() as UITextBox;
                if (textBox != null)
                {
                    textBox.Text = data[i];
                }
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
