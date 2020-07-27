using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Baosight.ColdRolling.LogicLayer
{
    public class FileHelper
    {
        public static bool createConfigFile(string fileName, List<Dictionary<string, string>> data)
        {
            StreamWriter sw = new StreamWriter(fileName, false, Encoding.UTF8);
            foreach (Dictionary<string, string> dictionary in data)
            {
                string content = "";
                foreach (var item in dictionary)
                {
                    content += string.Format("{0}={1};");
                }
                sw.WriteLine(content);
            }
            sw.Close();
            return true;
        }

        public static List<Dictionary<string, string>> readConfigFile(string fileName)
        {
            StreamReader sr = new StreamReader(fileName, Encoding.UTF8);
            List<Dictionary<string, string>> result = new List<Dictionary<string, string>>();

            string nextLine;
            while ((nextLine = sr.ReadLine()) != null)
            {
                Dictionary<string, string> dictionary = new Dictionary<string, string>();
                string[] items = nextLine.Split(';');
                foreach (string item in items)
                {
                    if (item == "")
                    {
                        continue;
                    }
                    string[] keyandvalue = item.Split('=');
                    string key = keyandvalue[0];
                    string value = keyandvalue[1];
                    dictionary.Add(key, value);
                }
                result.Add(dictionary);
            }
            sr.Close();
            return result;
        }
    }
}
