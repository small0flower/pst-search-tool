using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace PstSearchTool.Config
{
    [Serializable]
    public class StoreSetting
    {
        public string StoreId;
        public string DisplayName;
        /// <summary>已選資料夾，每項格式「kind::路徑」（如 inbox::收件匣）。</summary>
        public List<string> SelectedFolders = new List<string>();
    }

    [Serializable]
    public class AppSettings
    {
        public int SourceMode;                       // 0=目前 PST，1=外掛 PST
        public string ExternalDir = "";
        public List<string> ExternalPstFiles = new List<string>();
        public List<StoreSetting> Stores = new List<StoreSetting>();

        public string LastKeyword = "";
        public bool DateFilterEnabled;
        public string DateFrom = "";
        public string DateTo = "";
        public string SenderFilter = "";

        public int WindowX, WindowY, WindowW, WindowH;
        public bool WindowMaximized;

        public bool HintShown;

        /// <summary>主題："light" / "dark"</summary>
        public string Theme = "light";

        public string DbPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PstSearchTool", "index.db");
    }

    internal static class Settings
    {
        private static readonly string Dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PstSearchTool");
        private static readonly string File = Path.Combine(Dir, "config.xml");

        public static AppSettings Load()
        {
            try
            {
                if (System.IO.File.Exists(File))
                {
                    var xs = new XmlSerializer(typeof(AppSettings));
                    using (var fs = System.IO.File.OpenRead(File))
                        return (AppSettings)xs.Deserialize(fs);
                }
            }
            catch { }
            return new AppSettings();
        }

        public static void Save(AppSettings s)
        {
            try
            {
                Directory.CreateDirectory(Dir);
                var xs = new XmlSerializer(typeof(AppSettings));
                using (var fs = System.IO.File.Create(File))
                    xs.Serialize(fs, s);
            }
            catch { }
        }
    }
}
