using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SEBatchReplace
{
    public static class SolidEdgeReader
    {
        public static IList<string> ReadActiveAssemblyOccurrences(bool includeNested)
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
                string rootPath = Convert.ToString(GetProperty(document, "FullName"));
                return includeNested ? ReadOccurrencesRecursive(occurrences, 0, rootPath, new HashSet<string>(StringComparer.OrdinalIgnoreCase)) : ReadOccurrencesFirstLevel(occurrences);
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

        // Opens an assembly without creating a document window, while temporarily hiding
        // the Solid Edge main window. Reads first-level occurrences only and never saves.
        public static IList<string> OpenReadAndCloseAssemblyInBackground(string filePath)
        {
            if (String.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("找不到指定的組立檔。", filePath);

            object application = null;
            object documents = null;
            object existingDocument = null;
            object document = null;
            object occurrences = null;
            bool createdApplication = false;
            bool originalVisible = true;
            bool visibilityCaptured = false;
            bool documentOpened = false;
            bool closeAttempted = false;
            string stage = "連接或啟動 Solid Edge";

            try
            {
                try
                {
                    application = Marshal.GetActiveObject("SolidEdge.Application");
                }
                catch
                {
                    Type applicationType = Type.GetTypeFromProgID("SolidEdge.Application");
                    if (applicationType == null)
                        throw new InvalidOperationException("系統未註冊 SolidEdge.Application。請確認 Solid Edge ST6 已安裝並完成 COM 註冊。");
                    application = Activator.CreateInstance(applicationType);
                    createdApplication = true;
                }

                stage = "讀取 Solid Edge 原本的顯示狀態";
                originalVisible = Convert.ToBoolean(GetProperty(application, "Visible"));
                visibilityCaptured = true;

                stage = "隱藏 Solid Edge 主視窗";
                SetProperty(application, "Visible", false);

                stage = "取得 Solid Edge Documents 集合";
                documents = GetProperty(application, "Documents");

                stage = "確認測試組立尚未開啟";
                existingDocument = FindOpenDocument(documents, filePath);
                if (existingDocument != null)
                    throw new InvalidOperationException("這個組立目前已經在 Solid Edge 中開啟。請先關閉該組立，再執行背景開啟測試，以免測試到既有視窗。");

                stage = "使用 JDOCUMENTPROP_NOWINDOW=8 背景開啟組立";
                document = documents.GetType().InvokeMember(
                    "Open",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    documents,
                    new object[] { filePath, (uint)8 });
                if (document == null)
                    throw new InvalidOperationException("Documents.Open 沒有回傳文件物件。");
                documentOpened = true;

                stage = "取得背景組立 Occurrences";
                occurrences = GetProperty(document, "Occurrences");
                stage = "讀取背景組立第一層 Occurrences";
                IList<string> results = ReadOccurrencesFirstLevel(occurrences);

                ReleaseComObject(occurrences);
                occurrences = null;

                stage = "關閉背景組立（不儲存）";
                closeAttempted = true;
                InvokeMethod(document, "Close", new object[] { false });
                documentOpened = false;
                ReleaseComObject(document);
                document = null;

                return results;
            }
            catch (Exception ex)
            {
                if (documentOpened && document != null && !closeAttempted)
                {
                    try
                    {
                        InvokeMethod(document, "Close", new object[] { false });
                        documentOpened = false;
                    }
                    catch
                    {
                        // Keep the original failure; cleanup may need manual attention.
                    }
                }

                throw new InvalidOperationException(
                    "Solid Edge 背景開啟測試失敗。檔案：" + filePath +
                    "；失敗階段：" + stage +
                    (documentOpened ? "；文件可能仍留在 Solid Edge 中，請檢查並手動關閉。" : String.Empty) +
                    "；原因：" + ex.Message, ex);
            }
            finally
            {
                ReleaseComObject(occurrences);
                ReleaseComObject(document);
                ReleaseComObject(existingDocument);
                ReleaseComObject(documents);

                if (createdApplication && application != null)
                {
                    try { InvokeMethod(application, "Quit", new object[0]); }
                    catch { }
                }
                else if (visibilityCaptured && application != null)
                {
                    try { SetProperty(application, "Visible", originalVisible); }
                    catch { }
                }

                ReleaseComObject(application);
            }
        }

        // Opens the requested assembly only when it is not already open.
        // A document opened by this method is closed without saving after reading.
        // A document that was already open is left open and untouched.
        public static IList<string> OpenReadAndCloseAssembly(string filePath, bool includeNested)
        {
            if (String.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("找不到指定的組立檔。", filePath);

            object application = null;
            object documents = null;
            object document = null;
            object occurrences = null;
            bool openedByThisMethod = false;
            bool closeAttempted = false;
            string stage = "連接 Solid Edge";

            try
            {
                application = Marshal.GetActiveObject("SolidEdge.Application");
                stage = "取得 Solid Edge Documents 集合";
                documents = GetProperty(application, "Documents");

                stage = "檢查組立文件是否已開啟";
                document = FindOpenDocument(documents, filePath);
                if (document == null)
                {
                    stage = "由 Solid Edge 開啟組立文件";
                    document = OpenDocument(documents, filePath);
                    openedByThisMethod = true;
                }

                stage = "取得組立 Occurrences";
                occurrences = GetProperty(document, "Occurrences");
                stage = includeNested ? "讀取完整組立階層" : "讀取主組立第一層";
                IList<string> results = includeNested ? ReadOccurrencesRecursive(occurrences, 0, filePath, new HashSet<string>(StringComparer.OrdinalIgnoreCase)) : ReadOccurrencesFirstLevel(occurrences);
                ReleaseComObject(occurrences);
                occurrences = null;

                if (openedByThisMethod)
                {
                    stage = "關閉工具開啟的組立文件（不儲存）";
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
                    "；失敗階段：" + stage +
                    (openedByThisMethod ? "；文件由工具開啟，請確認是否仍留在 Solid Edge 中。" : "；尚未確認文件是否已開啟，不會據此判定文件原本已開啟。") +
                    "；原因：" + ex.Message, ex);
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

        // Reads direct occurrences only. Subassemblies are listed but not expanded.
        private static IList<string> ReadOccurrencesFirstLevel(object occurrences)
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
                    bool isSubassembly = Convert.ToBoolean(GetProperty(occurrence, "Subassembly"));
                    if (String.IsNullOrWhiteSpace(fileName)) fileName = "(無法取得零件檔案路徑)";
                    string itemType = isSubassembly ? "[子組立，未展開]" : "[零件]";
                    results.Add(index.ToString() + ". " + itemType + " " + occurrenceName + " | " + Path.GetFileName(fileName) + " | " + fileName);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("讀取主組立第一層第 " + index + " 個參照時失敗。" + (occurrence == null ? String.Empty : " 參照：" + SafeGetOccurrenceName(occurrence)) + "；原因：" + ex.Message, ex);
                }
                finally { ReleaseComObject(occurrence); }
            }
            return results;
        }
        private static IList<string> ReadOccurrencesRecursive(
            object occurrences,
            int depth,
            string assemblyPath,
            HashSet<string> ancestorPaths)
        {
            if (depth > 64)
                throw new InvalidOperationException("組立階層超過 64 層，已停止讀取以避免異常遞迴。");

            string normalizedPath = NormalizePath(assemblyPath);
            if (!String.IsNullOrWhiteSpace(normalizedPath))
            {
                if (ancestorPaths.Contains(normalizedPath))
                    throw new InvalidOperationException("偵測到循環組立參照：" + normalizedPath);
                ancestorPaths.Add(normalizedPath);
            }

            var results = new List<string>();
            try
            {
                int count = Convert.ToInt32(GetProperty(occurrences, "Count"));

                for (int index = 1; index <= count; index++)
                {
                    object occurrence = null;
                    object subassemblyDocument = null;
                    object subOccurrences = null;
                    try
                    {
                        occurrence = GetIndexedProperty(occurrences, "Item", new object[] { index });
                        string fileName = Convert.ToString(GetProperty(occurrence, "OccurrenceFileName"));
                        string occurrenceName = Convert.ToString(GetProperty(occurrence, "Name"));
                        bool isSubassembly = Convert.ToBoolean(GetProperty(occurrence, "Subassembly"));

                        if (String.IsNullOrWhiteSpace(fileName))
                            fileName = "(無法取得零件檔案路徑)";

                        string indent = new string(' ', depth * 4);
                        string itemType = isSubassembly ? "[子組立]" : "[零件]";
                        results.Add(indent + index.ToString() + ". " + itemType + " " + occurrenceName +
                            " | " + Path.GetFileName(fileName) + " | " + fileName);

                        if (isSubassembly)
                        {
                            string childPath = fileName;
                            string normalizedChildPath = NormalizePath(childPath);
                            if (!String.IsNullOrWhiteSpace(normalizedChildPath) &&
                                ancestorPaths.Contains(normalizedChildPath))
                            {
                                results.Add(indent + "    [停止展開：偵測到循環參照] " + childPath);
                                continue;
                            }

                            subassemblyDocument = GetProperty(occurrence, "OccurrenceDocument");
                            if (subassemblyDocument == null)
                            {
                                results.Add(indent + "    [無法展開：OccurrenceDocument 為空]");
                                continue;
                            }

                            subOccurrences = GetProperty(subassemblyDocument, "Occurrences");
                            IList<string> children = ReadOccurrencesRecursive(
                                subOccurrences,
                                depth + 1,
                                childPath,
                                ancestorPaths);

                            foreach (string child in children)
                                results.Add(child);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(
                            "讀取第 " + index + " 個參照時失敗。" +
                            (occurrence == null ? String.Empty : " 參照：" + SafeGetOccurrenceName(occurrence)) +
                            "；原因：" + ex.Message, ex);
                    }
                    finally
                    {
                        ReleaseComObject(subOccurrences);
                        ReleaseComObject(subassemblyDocument);
                        ReleaseComObject(occurrence);
                    }
                }
                return results;
            }
            finally
            {
                if (!String.IsNullOrWhiteSpace(normalizedPath))
                    ancestorPaths.Remove(normalizedPath);
            }
        }

        private static string SafeGetOccurrenceName(object occurrence)
        {
            try
            {
                return Convert.ToString(GetProperty(occurrence, "Name"));
            }
            catch
            {
                return "(無法取得參照名稱)";
            }
        }

        private static string NormalizePath(string path)
        {
            if (String.IsNullOrWhiteSpace(path) || path.StartsWith("("))
                return null;

            try
            {
                return Path.GetFullPath(path);
            }
            catch
            {
                return null;
            }
        }

        private static void SetProperty(object target, string propertyName, object value)
        {
            try
            {
                target.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    target,
                    new object[] { value });
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("設定 Solid Edge COM 成員 '" + propertyName + "' 失敗。", ex);
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