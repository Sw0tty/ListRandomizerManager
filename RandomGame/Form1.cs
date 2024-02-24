using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Threading;
using System.Text.Json;


namespace RandomGame
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            LoadJson();

            string lastSelectedCategory = Properties.Settings.Default.lastSelectedCategory;

            if (lastSelectedCategory != "")
            {
                comboBox1.SelectedItem = lastSelectedCategory;
            }

            if (comboBox1.SelectedItem != null)
            {
                button7.Enabled = true;
                button8.Enabled = true;
                groupBox1.Enabled = true;
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
                        valueBox.Text = "";
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
            if (MessageBox.Show("Вы действительно хотите удалить выбранные записи?", "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (listBox1.SelectedItems.Count > 1)
                {
                    List<object> selectedItems = new List<object>();

                    foreach (object item in listBox1.SelectedItems)
                    {
                        selectedItems.Add(item);
                    }
                    foreach (object item in selectedItems)
                    {
                        JSONWorker.DeleteValue(comboBox1.SelectedItem.ToString(), item.ToString());
                        listBox1.Items.Remove(item);
                        button3.Enabled = false;
                        CheckCountList();
                    }
                }
                else
                {
                    object selectedValue = listBox1.SelectedItem;
                    JSONWorker.DeleteValue(comboBox1.SelectedItem.ToString(), selectedValue.ToString());
                    listBox1.Items.Remove(selectedValue);
                    button3.Enabled = false;
                    CheckCountList();
                }
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
            Form4 settingsForm = new Form4(comboBox1.Items);
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                var wb = new XLWorkbook();

                foreach (string category in settingsForm.CheckedValues())
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
                    JSONWorker.AddCategory(listBox1, comboBox1, newCategory);
                    textBox2.Text = "";
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
                button7.Enabled = false;
                groupBox1.Enabled = false;
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
            if (comboBox1.SelectedItem == null)
                groupBox1.Enabled = false;
            else
                groupBox1.Enabled = true;

            button3.Enabled = false;
            button9.Enabled = false;
            button10.Enabled = false;
            button8.Enabled = true;
            button7.Enabled = true;
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
            if (listBox1.SelectedItems.Count == 1)
            {
                button9.Enabled = true;
            }
            else
            {
                button9.Enabled = false;
            }
            button3.Enabled = true;
            button10.Enabled = true;
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
                JSONWorker.SaveCollectionFile();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form2 settingsForm = new Form2(comboBox1.SelectedItem.ToString());
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                string newCategoryName = settingsForm.NewValue();
                string oldCategoryName = comboBox1.SelectedItem.ToString();
                comboBox1.Items.Remove(oldCategoryName);
                comboBox1.Items.Add(newCategoryName);
                JSONWorker.UpdateCategory(oldCategoryName, newCategoryName);
                comboBox1.SelectedItem = newCategoryName;
                JSONWorker.SaveCollectionFile();
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Form3 settingsForm = new Form3(listBox1.SelectedItems, comboBox1.Items, comboBox1.SelectedItem.ToString());
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                string newCategory = settingsForm.NewCategory();
                List<string> valuesInNewCategory = JSONWorker.ReturnCategoryValues(newCategory);
                List<string> problemValues = new List<string>();

                foreach (string value in listBox1.SelectedItems)
                {
                    if (valuesInNewCategory.Contains(value))
                        problemValues.Add(value);
                }

                if (problemValues.Count > 0)
                {
                    MessageBox.Show($"В конечной категории уже присутствуют значения: {string.Join(", ", problemValues)}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    List<string> selectedItems = new List<string>(listBox1.SelectedItems.Cast<string>().ToList());
                    foreach (string value in selectedItems)
                    {
                        JSONWorker.AddValue(newCategory, value);
                        JSONWorker.DeleteValue(comboBox1.SelectedItem.ToString(), value);
                        listBox1.Items.Remove(value);
                    }
                    button3.Enabled = false;
                    button10.Enabled = false;
                    CheckCountList();
                }
            }
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

        private void button3_MouseEnter(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand;
        }

        private void button3_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
        }

        private void button4_MouseEnter(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand;
        }

        private void button4_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
        }

        private void button5_MouseEnter(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand;
        }

        private void button5_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
        }

        private void button6_MouseEnter(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand;
        }

        private void button6_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
        }

        private void button7_MouseEnter(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand;
        }

        private void button7_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
        }

        private void button8_MouseEnter(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand;
        }

        private void button8_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
        }

        private void button9_MouseEnter(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand;
        }

        private void button9_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
        }

        private void button10_MouseEnter(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand;
        }

        private void button10_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
        }
    }
}
