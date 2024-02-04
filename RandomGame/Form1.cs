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

            StringCollection collection = Properties.Settings.Default.listBoxValues;
            StringDictionary dict = Properties.Settings.Default.myDict;
            //Dictionary<string, List<string>> d = Properties.Settings.Default.dictionaryValues;


            //MessageBox.Show(collection.Count.ToString());

            //MessageBox.Show(d.Count.ToString());

            comboBox1.SelectedItem = "films";




            /*if (collection != null)
            {
                foreach (var item in collection)
                {
                    listBox1.Items.Add(item);
                }
            }*/
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
            string newValue = textBox1.Text.Trim(' ');


            if (newValue != "")
            {
                if (JSONWorker.DataFromFile.ContainsKey(newValue))
                {
                    MessageBox.Show("Значение уже сущетсвует");
                }
                else
                {
                    listBox1.Items.Add(newValue);

                    JSONWorker.AddValue(comboBox1.SelectedItem.ToString(), newValue);
                }
            }
/*
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
            }*/
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

        private void button4_Click(object sender, EventArgs e)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("My games");
            
/*            // Title
            ws.Cell("A1").Value = "All my games";*/

            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                ws.Cell($"A{i + 1}").Value = listBox1.Items[i].ToString();
            }
            ws.Columns().AdjustToContents();

            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
            {
                wb.SaveAs($"{dialog.SelectedPath}\\MyGames.xlsx");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                var workbook = new XLWorkbook(filePath);
                var ws = workbook.Worksheet(1);
                int count = 0;

                for (int i = 1; i < ws.RowCount(); i++)
                {
                    var row = ws.Row(i);
                    string cellValue = row.Cell(1).Value.ToString();

                    if (cellValue == "")
                    {
                        break;
                    }
                    if (listBox1.Items.Contains(cellValue))
                    {
                        continue;
                    }
                    listBox1.Items.Add(cellValue);
                    count++;
                }
                saveListBox();
                MessageBox.Show($"Добавлено {count} записей.");
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            for (int i = 1; i < 35; i++)
            {
                Random rnd = new Random();
                worker.ReportProgress(rnd.Next(0, listBox1.Items.Count));
                
                Thread.Sleep(10 * i);
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            label1.Text = listBox1.Items[e.ProgressPercentage].ToString();
            if (label1.Text == "PAYDAY 2")
            {
                MessageBox.Show("");
            }
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
                    MessageBox.Show("Категория уже сущетсвует");
                }
                else
                {
                    comboBox1.Items.Add(newCategory);
                    
                    JSONWorker.AddCategory(listBox1, comboBox1, newCategory);
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (textBox2.Text.Trim(' ') != "")
            {
                listBox2.Items.Add(textBox2.Text.Trim(' '));

            }
        }



        private void button8_Click(object sender, EventArgs e)
        {
            LoadJson();
           /* List<CategoryValues> categoryValues = new List<CategoryValues>();

            categoryValues.Add(new CategoryValues() { CategoryName = "games", ValuesCat = new List<string>() { "mine", "some" } });
            categoryValues.Add(new CategoryValues() { CategoryName = "films", ValuesCat = new List<string>() { "mila", "dada" } });

            string json = JsonSerializer.Serialize(categoryValues);

            File.WriteAllText(@"D:\customers.json", json);*/
        }
        public void LoadJson()
        {
            using (StreamReader r = new StreamReader(@"D:\customers.json"))
            {
                string json = r.ReadToEnd();
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
                    //comboBox1.SelectedItem = item.CategoryName;
                }
                JSONWorker.DataFromFile = newDict;
                //JSONData.LoadCategoryData(listBox2, comboBox1);
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            JSONWorker.LoadCategoryData(listBox1, comboBox1);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            List<CategoryValues> categoryValues = new List<CategoryValues>();

            foreach (string category in JSONWorker.DataFromFile.Keys)
            {
                List<string> catValues = new List<string>();
                foreach (string value in JSONWorker.DataFromFile[category])
                {
                    catValues.Add(value);
                }
                categoryValues.Add(new CategoryValues() { CategoryName = category, ValuesCat = catValues });
            }

            
            //new CategoryValues() { CategoryName = "films", ValuesCat = new List<string>() { "mila", "dada" } };

            string json = JsonSerializer.Serialize(categoryValues);

            File.WriteAllText(@"D:\customers.json", json);
        }
    }
}
