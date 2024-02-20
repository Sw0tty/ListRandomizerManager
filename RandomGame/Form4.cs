using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RandomGame
{
    public partial class Form4 : Form
    {
        public Form4(ComboBox.ObjectCollection comboCollection)
        {
            InitializeComponent();
            
            foreach (string item in comboCollection)
            {
                checkedListBox1.Items.Add(item);
            }
        }

        public List<string> CheckedValues()
        {
            return checkedListBox1.CheckedItems.Cast<string>().ToList();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show("Не был выбран ни одная категория для выгрузки.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
