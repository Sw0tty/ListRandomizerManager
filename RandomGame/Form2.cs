﻿using System;
using System.Windows.Forms;


namespace RandomGame
{
    public partial class Form2 : Form
    {
        public string value;

        public Form2(string value)
        {
            InitializeComponent();
            textBox1.Text = value;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Trim(' ') == "")
            {
                MessageBox.Show("Новое значение не может быть пустым", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        public string NewValue()
        {
            return textBox1.Text.Trim(' ');
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
