using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.Windows.Forms;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;

namespace RandomGame
{
    public class CategoryValues
    {
        public string CategoryName { get; set; }
        public List<string> ValuesCat { get; set; }
    }

    public static class JSONWorker
    {
        public static string CollectionFileName = "CollectionOfValues.json";
        public static Dictionary<string, List<string>> DataFromFile = new Dictionary<string, List<string>>();

        public static void CreateCollectionFile()
        {
            File.Create(AppDomain.CurrentDomain.BaseDirectory + CollectionFileName).Close();
        }

        public static void SaveCollectionFile()
        {
            List<CategoryValues> categoryValues = new List<CategoryValues>();

            foreach (string category in DataFromFile.Keys)
            {
                List<string> catValues = new List<string>();
                foreach (string value in DataFromFile[category])
                {
                    catValues.Add(value);
                }
                categoryValues.Add(new CategoryValues() { CategoryName = category, ValuesCat = catValues });
            }

            string json = JsonSerializer.Serialize(categoryValues);

            File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + CollectionFileName, json);
        }

        public static void LoadCategoryData(ListBox listBox, ComboBox comboBox)
        {
            listBox.Items.Clear();
            foreach (string item in DataFromFile[comboBox.SelectedItem.ToString()])
            {
                listBox.Items.Add(item);
            }
        }

        public static List<string> ReturnCategoryValues(string category)
        {
            return DataFromFile[category].Cast<string>().ToList();
        }

        public static void AddCategory(ListBox listBox, ComboBox comboBox, string newCategory)
        {
            listBox.Items.Clear();
            comboBox.Items.Add(newCategory);
            DataFromFile[newCategory] = new List<string>();
            comboBox.SelectedItem = newCategory;
        }

        public static void AddValue(string category, string newValue)
        {
            DataFromFile[category].Add(newValue);
        }

        public static void UpdateValue(string category, string oldValue, string newValue)
        {
            foreach (string value in DataFromFile[category])
            {
                if (value == oldValue)
                {
                    DataFromFile[category][DataFromFile[category].IndexOf(oldValue)] = newValue;
                    break;
                }
            }
        }

        public static void UpdateCategory(string oldNameCategory, string newNameCategory)
        {
            List<string> categoryData = new List<string>(DataFromFile[oldNameCategory]);
            DataFromFile.Remove(oldNameCategory);
            DataFromFile.Add(newNameCategory, categoryData);
        }

        public static void DeleteValue(string category, string value)
        {
            DataFromFile[category].Remove(value);
        }

        public static void DeleteCategory(string category)
        {
            DataFromFile.Remove(category);
        }
    }
}
