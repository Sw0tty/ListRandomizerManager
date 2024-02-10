using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using ClosedXML.Excel;
using System.Runtime.InteropServices.ComTypes;
using DocumentFormat.OpenXml.Spreadsheet;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using System.Threading;
using DocumentFormat.OpenXml.Drawing;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace RandomGame
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            LoadJson();

            string lastSelectedCategory = Properties.Settings.Default.lastSelectedCategory;
            //Dictionary<string, List<string>> d = Properties.Settings.Default.dictionaryValues;

            if (lastSelectedCategory != "")
            {
                comboBox1.SelectedItem = lastSelectedCategory;
            }

            //MessageBox.Show(collection.Count.ToString());

            //MessageBox.Show(d.Count.ToString());

            //comboBox1.SelectedItem = "films";


            //MessageBox.Show(AppDomain.CurrentDomain.BaseDirectory);


            if (comboBox1.SelectedItem != null)
            {
                button8.Enabled = true;
            }
        }

        private void BlockGroups()
        {
            groupBox1.Enabled = false;
            groupBox2.Enabled = false;
            groupBox4.Enabled = false;
            groupBox5.Enabled = false;
        }

        private void EnablekGroups()
        {
            groupBox1.Enabled = true;
            groupBox2.Enabled = true;
            groupBox4.Enabled = true;
            groupBox5.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count != 0)
            {
                if (!backgroundWorker1.IsBusy)
                {
                    button1.Enabled = false;
                    backgroundWorker1.RunWorkerAsync();
                }
            }
            else
            {
                MessageBox.Show("Список пуст", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string newValueInCategory = valueBox.Text.Trim(' ');

            if (comboBox1.SelectedItem != null)
            {
                if (newValueInCategory != "")
                {
                    if (JSONWorker.DataFromFile[comboBox1.SelectedItem.ToString()].Contains(newValueInCategory))
                    {
                        MessageBox.Show("В категории уже присутствует данное значение.", "Ошибка");
                    }
                    else
                    {
                        listBox1.Items.Add(newValueInCategory);

                        JSONWorker.AddValue(comboBox1.SelectedItem.ToString(), newValueInCategory);
                        CheckCountList();
                    }
                }
                else
                {
                    MessageBox.Show("Значение не введено.", "Ошибка");
                }
            }
            else
            {
                MessageBox.Show("Выберите категорию.", "Ошибка");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItems.Count != 0)
            {
                object selectedValue = listBox1.SelectedItem;
                JSONWorker.DeleteValue(comboBox1.SelectedItem.ToString(), selectedValue.ToString());
                listBox1.Items.Remove(selectedValue);
                button3.Enabled = false;
                //saveListBox();
                CheckCountList();
            }
            else
            {
                MessageBox.Show("Ничего не выбрано", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void SaveProperties()
        {
            if (comboBox1.SelectedItem != null)
            {
                Properties.Settings.Default.lastSelectedCategory = comboBox1.SelectedItem.ToString();
                Properties.Settings.Default.Save();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var wb = new XLWorkbook();

            foreach (string category in JSONWorker.DataFromFile.Keys)
            {
                var ws = wb.Worksheets.Add(category);
                int valueCaount = 0;

                foreach (string value in JSONWorker.DataFromFile[category])
                {
                    ws.Cell($"A{valueCaount + 1}").Value = value;
                    valueCaount++;
                }
                ws.Columns().AdjustToContents();  
            }

            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
            {
                wb.SaveAs($"{dialog.SelectedPath}\\CollectionValue.xlsx");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            int count = 0;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                var workbook = new XLWorkbook(filePath);
                
                foreach (var ws in workbook.Worksheets)
                {
                    if (!JSONWorker.DataFromFile.ContainsKey(ws.Name))
                    {
                        JSONWorker.AddCategory(listBox1, comboBox1, ws.Name);
                    }

                    for (int i = 1; i < ws.RowCount(); i++)
                    {
                        var row = ws.Row(i);
                        string cellValue = row.Cell(1).Value.ToString();

                        if (cellValue == "")
                        {
                            break;
                        }
                        if (JSONWorker.DataFromFile[ws.Name].Contains(cellValue))
                        {
                            continue;
                        }
                        listBox1.Items.Add(cellValue);
                        JSONWorker.AddValue(ws.Name, cellValue);
                        count++;
                    }
                }
                MessageBox.Show($"Ипортировано {count} записей.");
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            Invoke(new Action(() => BlockGroups() ));
            BackgroundWorker worker = sender as BackgroundWorker;

            for (int i = 1; i < 35; i++)
            {
                Random rnd = new Random();
                worker.ReportProgress(rnd.Next(0, listBox1.Items.Count));
                
                Thread.Sleep(10 * i);
            }
            Invoke(new Action(() => EnablekGroups()));
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            label1.Text = listBox1.Items[e.ProgressPercentage].ToString();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button1.Enabled = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string newCategory = textBox2.Text.Trim(' ');
            if (newCategory != "")
            {
                if (JSONWorker.DataFromFile.ContainsKey(newCategory))
                {
                    MessageBox.Show("Категория уже сущетсвует.");
                }
                else
                {
                    //comboBox1.Items.Add(newCategory);
                    
                    JSONWorker.AddCategory(listBox1, comboBox1, newCategory);
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите удалить данную категорию?", "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                JSONWorker.DeleteCategory(comboBox1.SelectedItem.ToString());
                comboBox1.Items.Remove(comboBox1.SelectedItem);
                listBox1.Items.Clear();
                button8.Enabled = false;
                CheckCountList();
            }
        }

        public void LoadJson()
        {
            if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + JSONWorker.CollectionFileName))
            {
                using (StreamReader r = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + JSONWorker.CollectionFileName))
                {
                    string json = r.ReadToEnd();
                    if (json != "")
                    {
                        List<CategoryValues> items = JsonSerializer.Deserialize<List<CategoryValues>>(json);
                        Dictionary<string, List<string>> newDict = new Dictionary<string, List<string>>();

                        foreach (CategoryValues item in items)
                        {
                            comboBox1.Items.Add(item.CategoryName);

                            newDict[item.CategoryName] = new List<string>();
                            foreach (string value in item.ValuesCat)
                            {
                                newDict[item.CategoryName].Add(value);
                            }
                        }
                        JSONWorker.DataFromFile = newDict;
                    }

                }
            }
            else
            {
                JSONWorker.CreateCollectionFile();
            }            
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            button3.Enabled = false;
            button9.Enabled = false;
            button8.Enabled = true;
            JSONWorker.LoadCategoryData(listBox1, comboBox1);
            CheckCountList();
        }

        private void CheckCountList()
        {
            if (listBox1.Items.Count > 1)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            JSONWorker.SaveCollectionFile();
            SaveProperties();
        }

        private void listBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            button3.Enabled = true;
            button9.Enabled = true;
        }

        public string SelectedItem()
        {
            return listBox1.SelectedItem.ToString();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Form2 settingsForm = new Form2(listBox1.SelectedItem.ToString());
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                string newValue = settingsForm.NewValue();
                JSONWorker.UpdateValue(comboBox1.SelectedItem.ToString(), listBox1.SelectedItem.ToString(), newValue);
                listBox1.Items[listBox1.SelectedIndex] = newValue;
            }
            
        }
    }
}
