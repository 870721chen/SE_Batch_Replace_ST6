using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SEBatchReplace
{
    public static class FolderScanner
    {
        public static IList<string> GetAssemblyFiles(string folderPath)
        {
            if (String.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentException("請先選擇資料夾。", "folderPath");
            if (!Directory.Exists(folderPath))
                throw new DirectoryNotFoundException("找不到指定資料夾：" + folderPath);

            return Directory.GetFiles(folderPath, "*.asm", SearchOption.AllDirectories)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
    }
}