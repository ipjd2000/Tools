using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace LLDOCX.Options
{
    class ProcessData
    {
        private static string GetSaveDataFile()
        {
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LL-Docx");
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return Path.Combine(folder, "LLDocx.dll");
        }
        public static SettingOption GetData()
        {
            var settionOption = new SettingOption();
            string filePath = GetSaveDataFile();
            if (File.Exists(filePath))
            {
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    BinaryFormatter binaryFormatter = new BinaryFormatter();
                    settionOption = binaryFormatter.Deserialize(fileStream) as SettingOption;
                    fileStream.Close();
                    fileStream.Dispose();
                }
            }
            return settionOption;
        }
        public static void SaveData(SettingOption settingOption)
        {
            string filePath = GetSaveDataFile();
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {
                BinaryFormatter binaryFormatter = new BinaryFormatter();
                binaryFormatter.Serialize(fileStream, settingOption);
                fileStream.Close();
                fileStream.Dispose();
            }
        }
    }
}
