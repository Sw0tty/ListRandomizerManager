using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.Windows.Forms;


namespace RandomGame
{
    public class CategoryValues
    {
        public string CategoryName { get; set; }
        public List<string> ValuesCat { get; set; }
    }

    public static class JSONWorker
    {
        public static Dictionary<string, List<string>> DataFromFile = new Dictionary<string, List<string>>();

        public static void LoadCategoryData(ListBox listBox, ComboBox comboBox)
        {
            listBox.Items.Clear();
            foreach (string item in DataFromFile[comboBox.SelectedItem.ToString()])
            {
                listBox.Items.Add(item);
            }
        }

        public static void AddCategory(ListBox listBox, ComboBox comboBox, string newCategory)
        {
            listBox.Items.Clear();
            DataFromFile[newCategory] = new List<string>();
            comboBox.SelectedItem = newCategory;
        }

        public static void AddValue(string category, string newValue)
        {
            DataFromFile[category].Add(newValue);
        }
    }
}
