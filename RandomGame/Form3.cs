using System;
using System.Windows.Forms;

namespace RandomGame
{
    public partial class Form3 : Form
    {
        public Form3(ListBox.SelectedObjectCollection selectedCollection, ComboBox.ObjectCollection comboBoxCollection, string selectedItem)
        {
            InitializeComponent();
            foreach (object item in selectedCollection)
            {
                if (textBox1.Text == "")
                {
                    textBox1.AppendText(item.ToString());
                    continue;
                }
                textBox1.AppendText("\r\n" + item.ToString());
            }
            
            foreach (object item in comboBoxCollection)
            {
                if (item.ToString() == selectedItem)
                    continue;
                comboBox1.Items.Add(item.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("Для перемещения необходимо выбрать категорию.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        public string NewCategory()
        {
            return comboBox1.SelectedItem.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand;
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand;
        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
        }
    }
}
