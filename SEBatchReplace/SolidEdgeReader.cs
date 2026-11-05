using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SEBatchReplace
{
    public static class SolidEdgeReader
    {
        public static IList<string> ReadActiveAssemblyOccurrences()
        {
            object application = null;
            object document = null;
            object occurrences = null;
            var results = new List<string>();

            try
            {
                application = Marshal.GetActiveObject("SolidEdge.Application");
                document = GetProperty(application, "ActiveDocument");
                if (document == null)
                    throw new InvalidOperationException("Solid Edge 目前沒有開啟中的文件。");

                try
                {
                    occurrences = GetProperty(document, "Occurrences");
                }
                catch
                {
                    throw new InvalidOperationException("目前作用中的文件不是可讀取的組立文件。請先在 Solid Edge 開啟一個 .asm 文件並切換到該文件。");
                }

                int count = Convert.ToInt32(GetProperty(occurrences, "Count"));
                for (int index = 1; index <= count; index++)
                {
                    object occurrence = null;
                    try
                    {
                        occurrence = GetIndexedProperty(occurrences, "Item", new object[] { index });
                        string fileName = Convert.ToString(GetProperty(occurrence, "OccurrenceFileName"));
                        string occurrenceName = Convert.ToString(GetProperty(occurrence, "Name"));

                        if (String.IsNullOrWhiteSpace(fileName))
                            fileName = "(無法取得零件檔案路徑)";

                        results.Add(index.ToString() + ". " + occurrenceName + " | " + Path.GetFileName(fileName) + " | " + fileName);
                    }
                    finally
                    {
                        ReleaseComObject(occurrence);
                    }
                }

                return results;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    "無法讀取 Solid Edge。請確認 Solid Edge ST6 已啟動，且程式與 Solid Edge 使用相同的 Windows 權限層級。COM 錯誤：" + ex.Message, ex);
            }
            finally
            {
                ReleaseComObject(occurrences);
                ReleaseComObject(document);
                ReleaseComObject(application);
            }
        }

        private static object GetProperty(object target, string propertyName)
        {
            try
            {
                return target.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    target,
                    null);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "讀取 Solid Edge COM 成員 '" + propertyName + "' 失敗。", ex);
            }
        }

        private static object GetIndexedProperty(object target, string propertyName, object[] arguments)
        {
            try
            {
                // Solid Edge collection Item(index) is exposed as an IDispatch method,
                // not as a normal .NET indexed property.
                return target.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    target,
                    arguments);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "存取 Solid Edge COM 索引成員 '" + propertyName + "' 失敗。", ex);
            }
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}