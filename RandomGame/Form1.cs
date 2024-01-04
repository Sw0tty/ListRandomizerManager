using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace RandomGame
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            StringCollection collection = Properties.Settings.Default.listBoxValues;

            if (collection != null)
            {
                foreach (var item in collection)
                {
                    listBox1.Items.Add(item);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count != 0)
            {
                Random rnd = new Random();
                label1.Text = listBox1.Items[rnd.Next(0, listBox1.Items.Count)].ToString();
            }
            else
            {
                MessageBox.Show("Список пуст", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string gameName = textBox1.Text;
            if (gameName != "")
            {
                bool gameExist = false;

                foreach (var item in listBox1.Items)
                {
                    if (item.ToString() == gameName)
                    {
                        gameExist = true;
                        break;
                    }
                }

                if (gameExist)
                {
                    MessageBox.Show("Данная игра уже присутствует в списке", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    listBox1.Items.Add(gameName);
                    saveListBox();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItems.Count != 0)
            {
                listBox1.Items.Remove(listBox1.SelectedItem);
                saveListBox();
            }
            else
            {
                MessageBox.Show("Ничего не выбрано", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void saveListBox()
        {
            StringCollection coll = new StringCollection();
            coll.AddRange(listBox1.Items.Cast<string>().ToArray());
            Properties.Settings.Default.listBoxValues = coll;
            Properties.Settings.Default.Save();
        }
    }
}
