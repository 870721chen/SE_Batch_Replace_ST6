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
            try
            {
                application = Marshal.GetActiveObject("SolidEdge.Application");
                document = GetProperty(application, "ActiveDocument");
                if (document == null)
                    throw new InvalidOperationException("Solid Edge 目前沒有開啟中的文件。");

                occurrences = GetProperty(document, "Occurrences");
                return ReadOccurrences(occurrences);
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

        // Opens the requested assembly only when it is not already open.
        // A document opened by this method is closed without saving after reading.
        // A document that was already open is left open and untouched.
        public static IList<string> OpenReadAndCloseAssembly(string filePath)
        {
            if (String.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("找不到指定的組立檔。", filePath);

            object application = null;
            object documents = null;
            object document = null;
            object occurrences = null;
            bool openedByThisMethod = false;
            bool closeAttempted = false;

            try
            {
                application = Marshal.GetActiveObject("SolidEdge.Application");
                documents = GetProperty(application, "Documents");

                document = FindOpenDocument(documents, filePath);
                if (document == null)
                {
                    document = OpenDocument(documents, filePath);
                    openedByThisMethod = true;
                }

                occurrences = GetProperty(document, "Occurrences");
                IList<string> results = ReadOccurrences(occurrences);
                ReleaseComObject(occurrences);
                occurrences = null;

                if (openedByThisMethod)
                {
                    closeAttempted = true;
                    InvokeMethod(document, "Close", new object[] { false });
                    ReleaseComObject(document);
                    document = null;
                }

                return results;
            }
            catch (Exception ex)
            {
                if (openedByThisMethod && document != null && !closeAttempted)
                {
                    try
                    {
                        InvokeMethod(document, "Close", new object[] { false });
                    }
                    catch
                    {
                        // Preserve the original diagnostic. The caller is told that close may need manual attention.
                    }
                }

                throw new InvalidOperationException(
                    "自動開啟／讀取／關閉組立時失敗。檔案：" + filePath +
                    (openedByThisMethod ? "；此文件由工具開啟。" : "；此文件原本已開啟，工具不會主動關閉它。") +
                    " 原因：" + ex.Message, ex);
            }
            finally
            {
                ReleaseComObject(occurrences);
                ReleaseComObject(document);
                ReleaseComObject(documents);
                ReleaseComObject(application);
            }
        }

        private static object FindOpenDocument(object documents, string filePath)
        {
            int count = Convert.ToInt32(GetProperty(documents, "Count"));
            string target = Path.GetFullPath(filePath);

            for (int index = 1; index <= count; index++)
            {
                object candidate = null;
                try
                {
                    candidate = GetIndexedProperty(documents, "Item", new object[] { index });
                    string fullName = Convert.ToString(GetProperty(candidate, "FullName"));
                    if (!String.IsNullOrWhiteSpace(fullName) &&
                        String.Equals(Path.GetFullPath(fullName), target, StringComparison.OrdinalIgnoreCase))
                    {
                        object match = candidate;
                        candidate = null;
                        return match;
                    }
                }
                finally
                {
                    ReleaseComObject(candidate);
                }
            }
            return null;
        }

        private static object OpenDocument(object documents, string filePath)
        {
            try
            {
                // Documents.Open has one required filename and five optional COM arguments.
                object[] arguments = new object[]
                {
                    filePath,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing
                };

                return documents.GetType().InvokeMember(
                    "Open",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    documents,
                    arguments);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("呼叫 Solid Edge Documents.Open 開啟組立檔失敗。", ex);
            }
        }

        private static IList<string> ReadOccurrences(object occurrences)
        {
            var results = new List<string>();
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

                    results.Add(index.ToString() + ". " + occurrenceName + " | " +
                        Path.GetFileName(fileName) + " | " + fileName);
                }
                finally
                {
                    ReleaseComObject(occurrence);
                }
            }
            return results;
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

        private static object InvokeMethod(object target, string methodName, object[] arguments)
        {
            try
            {
                return target.GetType().InvokeMember(
                    methodName,
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    target,
                    arguments);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("呼叫 Solid Edge COM 方法 '" + methodName + "' 失敗。", ex);
            }
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}