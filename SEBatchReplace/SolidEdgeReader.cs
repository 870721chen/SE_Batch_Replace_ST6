using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SEBatchReplace
{
    public static class SolidEdgeReader
    {

        // Reads runtime IDispatch type information for the assembly document,
        // its Occurrences collection, and one Occurrence. This method never edits or saves.
        public static IList<string> ProbeAssemblyApiInBackground(string filePath)
        {
            if (String.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("找不到指定的組立檔。", filePath);

            object application = null;
            object documents = null;
            object document = null;
            object occurrences = null;
            object occurrence = null;
            bool createdApplication = false;
            bool originalVisible = true;
            bool visibilityCaptured = false;
            bool documentOpened = false;
            var output = new List<string>();
            string stage = "連接或啟動 Solid Edge";

            try
            {
                try { application = Marshal.GetActiveObject("SolidEdge.Application"); }
                catch
                {
                    Type appType = Type.GetTypeFromProgID("SolidEdge.Application");
                    if (appType == null)
                        throw new InvalidOperationException("系統未註冊 SolidEdge.Application。");
                    application = Activator.CreateInstance(appType);
                    createdApplication = true;
                }

                stage = "讀取原始顯示狀態";
                originalVisible = Convert.ToBoolean(GetProperty(application, "Visible"));
                visibilityCaptured = true;
                SetProperty(application, "Visible", false);

                stage = "取得 Documents";
                documents = GetProperty(application, "Documents");
                stage = "確認組立未開啟";
                object alreadyOpen = FindOpenDocument(documents, filePath);
                if (alreadyOpen != null)
                {
                    ReleaseComObject(alreadyOpen);
                    throw new InvalidOperationException("測試組立目前已開啟。請先關閉該組立，再執行 API 診斷。");
                }

                stage = "背景開啟組立";
                document = documents.GetType().InvokeMember(
                    "Open",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null, documents, new object[] { filePath, (uint)8 });
                if (document == null)
                    throw new InvalidOperationException("Documents.Open 沒有回傳文件物件。");
                documentOpened = true;

                output.Add("[TypeInfo] AssemblyDocument runtime object");
                output.AddRange(ComTypeInfoInspector.Describe(document, "AssemblyDocument"));
                stage = "取得 Occurrences";
                occurrences = GetProperty(document, "Occurrences");
                output.Add("[TypeInfo] Occurrences runtime object");
                output.AddRange(ComTypeInfoInspector.Describe(occurrences, "Occurrences"));

                int count = Convert.ToInt32(GetProperty(occurrences, "Count"));
                output.Add("[資訊] 主組立第一層 occurrence 數量：" + count);
                if (count > 0)
                {
                    stage = "取得第一個 Occurrence";
                    occurrence = GetIndexedProperty(occurrences, "Item", new object[] { 1 });
                    output.Add("[TypeInfo] Occurrence runtime object（第一個項目）");
                    output.AddRange(ComTypeInfoInspector.Describe(occurrence, "Occurrence"));
                }
                else
                {
                    output.Add("[資訊] 組立沒有第一層 occurrence，略過 Occurrence TypeInfo。");
                }

                stage = "關閉組立（不儲存）";
                InvokeMethod(document, "Close", new object[] { false });
                documentOpened = false;
                output.Add("[完成] 背景組立已關閉，未儲存。");
                return output;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "Solid Edge COM API 診斷失敗。檔案：" + filePath +
                    "；失敗階段：" + stage +
                    (documentOpened ? "；文件可能仍留在 Solid Edge，請檢查並手動關閉。" : String.Empty) +
                    "；原因：" + ex.Message, ex);
            }
            finally
            {
                ReleaseComObject(occurrence);
                ReleaseComObject(occurrences);
                if (documentOpened && document != null)
                {
                    try { InvokeMethod(document, "Close", new object[] { false }); }
                    catch { }
                }
                ReleaseComObject(document);
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


        // Match an exact .par filename or a part-code prefix followed by a space.
        // Example: P03-22-00.par matches P03-22-00 燒結-10.1x19.9x10.par,
        // but does not match P03-22-001.par.
        private static bool IsMatchingPartFileName(string candidateFileName, string expectedPartName)
        {
            if (String.IsNullOrWhiteSpace(candidateFileName) || String.IsNullOrWhiteSpace(expectedPartName))
                return false;

            string expected = Path.GetFileName(expectedPartName.Trim());
            if (!expected.EndsWith(".par", StringComparison.OrdinalIgnoreCase))
                expected += ".par";

            if (String.Equals(candidateFileName, expected, StringComparison.OrdinalIgnoreCase))
                return true;

            string partCode = expected.Substring(0, expected.Length - 4);
            return candidateFileName.StartsWith(partCode + " ", StringComparison.OrdinalIgnoreCase);
        }

        // Lists every first-level occurrence and marks exact or part-code-prefix matches.
        // Read-only: no Replace call and no document save.
        public static IList<string> ListFirstLevelOccurrencesInBackground(string assemblyPath, string oldPartName)
        {
            if (String.IsNullOrWhiteSpace(assemblyPath) || !File.Exists(assemblyPath))
                throw new FileNotFoundException("找不到指定的測試組立檔。", assemblyPath);
            if (String.IsNullOrWhiteSpace(oldPartName))
                throw new ArgumentException("舊零件名稱不可空白。", "oldPartName");

            string expectedName = Path.GetFileName(oldPartName.Trim());
            if (!expectedName.EndsWith(".par", StringComparison.OrdinalIgnoreCase))
                expectedName += ".par";

            object application = null, documents = null, document = null, occurrences = null;
            bool createdApplication = false, visibilityCaptured = false, originalVisible = true, documentOpened = false;
            var output = new List<string>();
            string stage = "連接或啟動 Solid Edge";
            try
            {
                try { application = Marshal.GetActiveObject("SolidEdge.Application"); }
                catch
                {
                    Type appType = Type.GetTypeFromProgID("SolidEdge.Application");
                    if (appType == null) throw new InvalidOperationException("系統未註冊 SolidEdge.Application。");
                    application = Activator.CreateInstance(appType);
                    createdApplication = true;
                }

                originalVisible = Convert.ToBoolean(GetProperty(application, "Visible"));
                visibilityCaptured = true;
                SetProperty(application, "Visible", false);
                documents = GetProperty(application, "Documents");

                stage = "確認測試組立未開啟";
                object alreadyOpen = FindOpenDocument(documents, assemblyPath);
                if (alreadyOpen != null)
                {
                    ReleaseComObject(alreadyOpen);
                    throw new InvalidOperationException("測試組立目前已在 Solid Edge 中開啟。請先關閉它，再執行唯讀診斷。");
                }

                stage = "背景開啟測試組立";
                document = documents.GetType().InvokeMember("Open",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null, documents, new object[] { assemblyPath, (uint)8 });
                if (document == null) throw new InvalidOperationException("Documents.Open 沒有回傳組立文件。");
                documentOpened = true;

                stage = "讀取主組立第一層 occurrences";
                occurrences = GetProperty(document, "Occurrences");
                int count = Convert.ToInt32(GetProperty(occurrences, "Count"));
                output.Add("[組立] " + assemblyPath);
                output.Add("[比對目標] " + expectedName);
                output.Add("[第一層數量] " + count);
                int matchCount = 0;

                for (int index = 1; index <= count; index++)
                {
                    object item = null;
                    try
                    {
                        item = GetIndexedProperty(occurrences, "Item", new object[] { index });
                        string name = Convert.ToString(GetProperty(item, "Name"));
                        string path = Convert.ToString(GetProperty(item, "OccurrenceFileName"));
                        bool subassembly = Convert.ToBoolean(GetProperty(item, "Subassembly"));
                        string fileName = String.IsNullOrWhiteSpace(path) ? String.Empty : Path.GetFileName(path);
                        bool match = !subassembly && IsMatchingPartFileName(fileName, expectedName);
                        if (match) matchCount++;
                        output.Add("[Occurrence #" + index + "] " + (subassembly ? "[子組立]" : "[零件]") +
                            "；Name=" + (String.IsNullOrWhiteSpace(name) ? "(空)" : name) +
                            "；FileName=" + (String.IsNullOrWhiteSpace(fileName) ? "(空)" : fileName) +
                            "；Path=" + (String.IsNullOrWhiteSpace(path) ? "(空)" : path) +
                            "；Match=" + (match ? "YES" : "NO"));
                    }
                    finally { ReleaseComObject(item); }
                }

                output.Add("[比對結果] 完全符合檔名的第一層零件 occurrence：" + matchCount);
                output.Add("[限制] 本診斷只列主組立第一層，不展開子組立；不修改、不儲存。");

                stage = "關閉測試組立（不儲存）";
                InvokeMethod(document, "Close", new object[] { false });
                documentOpened = false;
                output.Add("[完成] 已呼叫 Close(false)，未儲存。");
                return output;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("第一層 occurrence 唯讀診斷失敗。檔案：" + assemblyPath +
                    "；失敗階段：" + stage + (documentOpened ? "；文件可能仍開啟，請檢查 Solid Edge。" : "") +
                    "；原因：" + ex.Message, ex);
            }
            finally
            {
                ReleaseComObject(occurrences);
                if (documentOpened && document != null)
                {
                    try { InvokeMethod(document, "Close", new object[] { false }); } catch { }
                }
                ReleaseComObject(document);
                ReleaseComObject(documents);
                if (createdApplication && application != null)
                {
                    try { InvokeMethod(application, "Quit", new object[0]); } catch { }
                }
                else if (visibilityCaptured && application != null)
                {
                    try { SetProperty(application, "Visible", originalVisible); } catch { }
                }
                ReleaseComObject(application);
            }
        }

        // Replaces exactly one matching first-level part occurrence in memory only.
        // The test assembly is always closed with SaveChanges=false.
        public static IList<string> ReplaceOneOccurrenceInBackground(string assemblyPath, string oldPartName, string newPartPath)
        {
            if (String.IsNullOrWhiteSpace(assemblyPath) || !File.Exists(assemblyPath))
                throw new FileNotFoundException("找不到指定的測試組立檔。", assemblyPath);
            if (String.IsNullOrWhiteSpace(newPartPath) || !File.Exists(newPartPath))
                throw new FileNotFoundException("找不到指定的新零件檔。", newPartPath);
            if (String.IsNullOrWhiteSpace(oldPartName))
                throw new ArgumentException("舊零件名稱不可空白。", "oldPartName");

            string expectedOldName = Path.GetFileName(oldPartName.Trim());
            if (!expectedOldName.EndsWith(".par", StringComparison.OrdinalIgnoreCase))
                expectedOldName += ".par";

            object application = null, documents = null, document = null, occurrences = null, targetOccurrence = null;
            bool createdApplication = false, visibilityCaptured = false, originalVisible = true, documentOpened = false;
            string stage = "連接或啟動 Solid Edge";
            var output = new List<string>();

            try
            {
                try { application = Marshal.GetActiveObject("SolidEdge.Application"); }
                catch
                {
                    Type appType = Type.GetTypeFromProgID("SolidEdge.Application");
                    if (appType == null) throw new InvalidOperationException("系統未註冊 SolidEdge.Application。");
                    application = Activator.CreateInstance(appType);
                    createdApplication = true;
                }

                originalVisible = Convert.ToBoolean(GetProperty(application, "Visible"));
                visibilityCaptured = true;
                SetProperty(application, "Visible", false);
                documents = GetProperty(application, "Documents");

                stage = "確認測試組立未開啟";
                object alreadyOpen = FindOpenDocument(documents, assemblyPath);
                if (alreadyOpen != null)
                {
                    ReleaseComObject(alreadyOpen);
                    throw new InvalidOperationException("測試組立目前已在 Solid Edge 中開啟。請先關閉它，再執行替換測試。");
                }

                stage = "背景開啟測試組立";
                document = documents.GetType().InvokeMember("Open",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null, documents, new object[] { assemblyPath, (uint)8 });
                if (document == null) throw new InvalidOperationException("Documents.Open 沒有回傳組立文件。");
                documentOpened = true;

                stage = "讀取主組立第一層 occurrences";
                occurrences = GetProperty(document, "Occurrences");
                int count = Convert.ToInt32(GetProperty(occurrences, "Count"));
                output.Add("[組立] " + assemblyPath);
                output.Add("[資訊] 第一層 occurrence 數量：" + count);

                int matchCount = 0, matchedIndex = 0;
                for (int index = 1; index <= count; index++)
                {
                    object candidate = null;
                    try
                    {
                        candidate = GetIndexedProperty(occurrences, "Item", new object[] { index });
                        bool isSubassembly = Convert.ToBoolean(GetProperty(candidate, "Subassembly"));
                        string candidatePath = Convert.ToString(GetProperty(candidate, "OccurrenceFileName"));
                        if (!isSubassembly && !String.IsNullOrWhiteSpace(candidatePath) &&
                            IsMatchingPartFileName(Path.GetFileName(candidatePath), expectedOldName))
                        {
                            matchCount++;
                            if (matchCount == 1)
                            {
                                targetOccurrence = candidate;
                                candidate = null;
                                matchedIndex = index;
                            }
                        }
                    }
                    finally { ReleaseComObject(candidate); }
                }

                output.Add("[搜尋] " + expectedOldName + "；符合的第一層 occurrence 數量：" + matchCount);
                if (matchCount == 0) throw new InvalidOperationException("主組立第一層找不到符合的零件 occurrence，未執行替換。");
                if (matchCount > 1) throw new InvalidOperationException("符合的 occurrence 多於一個，為避免替錯實例，本次未執行替換。");

                string occurrenceName = Convert.ToString(GetProperty(targetOccurrence, "Name"));
                string oldPath = Convert.ToString(GetProperty(targetOccurrence, "OccurrenceFileName"));
                int? beforeCount = TryGetRelationCount(targetOccurrence);
                output.Add("[目標] occurrence #" + matchedIndex + "；名稱：" + occurrenceName);
                output.Add("[替換前路徑] " + oldPath);
                output.Add("[Relations3d 數量（前）] " + (beforeCount.HasValue ? beforeCount.Value.ToString() : "無法讀取"));

                stage = "呼叫 Occurrence.Replace（ReplaceAll=false）";
                targetOccurrence.GetType().InvokeMember("Replace",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance | BindingFlags.OptionalParamBinding,
                    null, targetOccurrence, new object[] { newPartPath, false, Type.Missing });

                stage = "讀取替換後資訊";
                string afterPath = Convert.ToString(GetProperty(targetOccurrence, "OccurrenceFileName"));
                int? afterCount = TryGetRelationCount(targetOccurrence);
                output.Add("[替換後路徑] " + afterPath);
                output.Add("[Relations3d 數量（後）] " + (afterCount.HasValue ? afterCount.Value.ToString() : "無法讀取"));
                bool pathMatches = false;
                try { pathMatches = String.Equals(Path.GetFullPath(afterPath), Path.GetFullPath(newPartPath), StringComparison.OrdinalIgnoreCase); }
                catch { }
                output.Add("[路徑檢查] " + (pathMatches ? "替換後路徑符合指定的新零件。" : "替換呼叫已返回，但路徑尚未確認符合，需檢查。"));
                output.Add("[限制] Relations3d 數量相同不代表所有 mate/relationship 均有效；本次尚未做完整關聯有效性驗證。");

                stage = "關閉測試組立且不儲存";
                InvokeMethod(document, "Close", new object[] { false });
                documentOpened = false;
                output.Add("[完成] 已呼叫 Close(false)，本次替換未寫入組立檔。");
                return output;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("單一 occurrence 替換測試失敗。檔案：" + assemblyPath +
                    "；失敗階段：" + stage + (documentOpened ? "；文件可能仍開啟，請檢查 Solid Edge。" : "") +
                    "；原因：" + ex.Message, ex);
            }
            finally
            {
                ReleaseComObject(targetOccurrence);
                ReleaseComObject(occurrences);
                if (documentOpened && document != null)
                {
                    try { InvokeMethod(document, "Close", new object[] { false }); } catch { }
                }
                ReleaseComObject(document);
                ReleaseComObject(documents);
                if (createdApplication && application != null)
                {
                    try { InvokeMethod(application, "Quit", new object[0]); } catch { }
                }
                else if (visibilityCaptured && application != null)
                {
                    try { SetProperty(application, "Visible", originalVisible); } catch { }
                }
                ReleaseComObject(application);
            }
        }


        // Diagnostic isolation test: open an assembly in background and call Save without
        // performing any Replace operation. Use only a disposable test copy.
        public static IList<string> TestBackgroundSaveWithoutChanges(string assemblyPath)
        {
            if (String.IsNullOrWhiteSpace(assemblyPath) || !File.Exists(assemblyPath))
                throw new FileNotFoundException("找不到指定的測試組立檔。", assemblyPath);

            object application = null, documents = null, document = null;
            bool createdApplication = false, visibilityCaptured = false, originalVisible = true, documentOpened = false;
            string stage = "連接或啟動 Solid Edge";
            var output = new List<string>();

            try
            {
                try { application = Marshal.GetActiveObject("SolidEdge.Application"); }
                catch
                {
                    Type appType = Type.GetTypeFromProgID("SolidEdge.Application");
                    if (appType == null) throw new InvalidOperationException("系統未註冊 SolidEdge.Application。");
                    application = Activator.CreateInstance(appType);
                    createdApplication = true;
                }

                stage = "讀取原始顯示狀態";
                originalVisible = Convert.ToBoolean(GetProperty(application, "Visible"));
                visibilityCaptured = true;
                SetProperty(application, "Visible", false);
                documents = GetProperty(application, "Documents");

                stage = "確認測試組立未開啟";
                object alreadyOpen = FindOpenDocument(documents, assemblyPath);
                if (alreadyOpen != null)
                {
                    ReleaseComObject(alreadyOpen);
                    throw new InvalidOperationException("測試組立目前已在 Solid Edge 中開啟。請先關閉它，再執行 Save 單獨測試。");
                }

                FileAttributes attributes = File.GetAttributes(assemblyPath);
                bool isReadOnly = (attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly;
                output.Add("[測試模式] 背景開啟 JDOCUMENTPROP_NOWINDOW=8；不執行 Replace；直接呼叫 Save。");
                output.Add("[檔案唯讀屬性] " + (isReadOnly ? "是" : "否"));
                output.Add("[組立檔] " + assemblyPath);
                if (isReadOnly)
                    throw new InvalidOperationException("測試組立檔具有 ReadOnly 屬性，先移除唯讀屬性後再測試。");

                stage = "背景開啟測試組立";
                document = documents.GetType().InvokeMember("Open",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null, documents, new object[] { assemblyPath, (uint)8 });
                if (document == null) throw new InvalidOperationException("Documents.Open 沒有回傳組立文件。");
                documentOpened = true;

                output.Add("[文件 FullName] " + Convert.ToString(GetProperty(document, "FullName")));
                stage = "讀取開啟後文件資訊";
                object occurrences = null;
                try
                {
                    occurrences = GetProperty(document, "Occurrences");
                    output.Add("[第一層 occurrence 數量] " + Convert.ToString(GetProperty(occurrences, "Count")));
                }
                finally { ReleaseComObject(occurrences); }

                stage = "背景文件 Save（未執行 Replace）";
                InvokeMethod(document, "Save", new object[0]);
                output.Add("[Save 結果] 成功；此測試沒有執行任何零件替換。");

                stage = "關閉測試組立";
                InvokeMethod(document, "Close", new object[] { false });
                documentOpened = false;
                output.Add("[完成] 已呼叫 Close(false)，文件已關閉。");
                return output;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("背景 Save 單獨測試失敗。檔案：" + assemblyPath +
                    "；失敗階段：" + stage +
                    (documentOpened ? "；文件可能仍開啟，請檢查 Solid Edge。" : "") +
                    "；原因：" + ex.Message, ex);
            }
            finally
            {
                if (documentOpened && document != null)
                {
                    try { InvokeMethod(document, "Close", new object[] { false }); } catch { }
                }
                ReleaseComObject(document);
                ReleaseComObject(documents);
                if (createdApplication && application != null)
                {
                    try { InvokeMethod(application, "Quit", new object[0]); } catch { }
                }
                else if (visibilityCaptured && application != null)
                {
                    try { SetProperty(application, "Visible", originalVisible); } catch { }
                }
                ReleaseComObject(application);
            }
        }

        // Replaces exactly one first-level occurrence, saves the selected test assembly,
        // closes it, then reopens it in the background to verify the persisted reference.
        public static IList<string> ReplaceOneOccurrenceSaveAndReopenInBackground(
            string assemblyPath, string oldPartName, string newPartPath)
        {
            if (String.IsNullOrWhiteSpace(assemblyPath) || !File.Exists(assemblyPath))
                throw new FileNotFoundException("找不到指定的測試組立檔。", assemblyPath);
            if (String.IsNullOrWhiteSpace(newPartPath) || !File.Exists(newPartPath))
                throw new FileNotFoundException("找不到指定的新零件檔。", newPartPath);
            if (String.IsNullOrWhiteSpace(oldPartName))
                throw new ArgumentException("舊零件名稱不可空白。", "oldPartName");

            string expectedOldName = Path.GetFileName(oldPartName.Trim());
            if (!expectedOldName.EndsWith(".par", StringComparison.OrdinalIgnoreCase))
                expectedOldName += ".par";
            string expectedNewPath = Path.GetFullPath(newPartPath);

            object application = null, documents = null, document = null;
            object occurrences = null, targetOccurrence = null;
            bool createdApplication = false, visibilityCaptured = false, originalVisible = true, documentOpened = false;
            string stage = "連接或啟動 Solid Edge";
            var output = new List<string>();

            try
            {
                try { application = Marshal.GetActiveObject("SolidEdge.Application"); }
                catch
                {
                    Type appType = Type.GetTypeFromProgID("SolidEdge.Application");
                    if (appType == null) throw new InvalidOperationException("系統未註冊 SolidEdge.Application。");
                    application = Activator.CreateInstance(appType);
                    createdApplication = true;
                }

                originalVisible = Convert.ToBoolean(GetProperty(application, "Visible"));
                visibilityCaptured = true;
                SetProperty(application, "Visible", false);
                documents = GetProperty(application, "Documents");

                stage = "確認測試組立未開啟";
                object alreadyOpen = FindOpenDocument(documents, assemblyPath);
                if (alreadyOpen != null)
                {
                    ReleaseComObject(alreadyOpen);
                    throw new InvalidOperationException("測試組立目前已在 Solid Edge 中開啟。請先關閉它，再執行儲存重開測試。");
                }

                stage = "背景開啟測試組立";
                document = documents.GetType().InvokeMember("Open",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null, documents, new object[] { assemblyPath, (uint)8 });
                if (document == null) throw new InvalidOperationException("Documents.Open 沒有回傳組立文件。");
                documentOpened = true;

                stage = "讀取主組立第一層 occurrences";
                occurrences = GetProperty(document, "Occurrences");
                int count = Convert.ToInt32(GetProperty(occurrences, "Count"));
                output.Add("[組立] " + assemblyPath);
                output.Add("[第一層 occurrence 數量] " + count);

                int matchCount = 0;
                for (int index = 1; index <= count; index++)
                {
                    object candidate = null;
                    try
                    {
                        candidate = GetIndexedProperty(occurrences, "Item", new object[] { index });
                        bool isSubassembly = Convert.ToBoolean(GetProperty(candidate, "Subassembly"));
                        string candidatePath = Convert.ToString(GetProperty(candidate, "OccurrenceFileName"));
                        if (!isSubassembly && !String.IsNullOrWhiteSpace(candidatePath) &&
                            IsMatchingPartFileName(Path.GetFileName(candidatePath), expectedOldName))
                        {
                            matchCount++;
                            if (matchCount == 1)
                            {
                                targetOccurrence = candidate;
                                candidate = null;
                            }
                        }
                    }
                    finally { ReleaseComObject(candidate); }
                }

                output.Add("[舊零件比對數量] " + matchCount);
                if (matchCount == 0)
                    throw new InvalidOperationException("主組立第一層找不到符合的舊零件 occurrence，未執行替換。");
                if (matchCount > 1)
                    throw new InvalidOperationException("符合的 occurrence 多於一個，為避免替錯實例，本次未執行替換。");

                string beforePath = Convert.ToString(GetProperty(targetOccurrence, "OccurrenceFileName"));
                int? beforeRelations = TryGetRelationCount(targetOccurrence);
                output.Add("[替換前路徑] " + beforePath);
                output.Add("[Relations3d 數量（替換前）] " +
                    (beforeRelations.HasValue ? beforeRelations.Value.ToString() : "無法讀取"));

                stage = "呼叫 Occurrence.Replace（ReplaceAll=false）";
                targetOccurrence.GetType().InvokeMember("Replace",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance | BindingFlags.OptionalParamBinding,
                    null, targetOccurrence, new object[] { newPartPath, false, Type.Missing });

                stage = "確認替換後路徑";
                string replacedPath = Convert.ToString(GetProperty(targetOccurrence, "OccurrenceFileName"));
                bool replacedPathMatches = false;
                try
                {
                    replacedPathMatches = String.Equals(Path.GetFullPath(replacedPath), expectedNewPath,
                        StringComparison.OrdinalIgnoreCase);
                }
                catch { }
                output.Add("[替換後路徑] " + replacedPath);
                output.Add("[替換後路徑檢查] " + (replacedPathMatches ? "符合指定的新零件。" : "不符合指定的新零件，停止儲存。"));
                if (!replacedPathMatches)
                    throw new InvalidOperationException("替換後路徑與指定的新零件不一致，為避免儲存錯誤結果，本次不執行 Save。");

                int? afterRelations = TryGetRelationCount(targetOccurrence);
                output.Add("[Relations3d 數量（替換後）] " +
                    (afterRelations.HasValue ? afterRelations.Value.ToString() : "無法讀取"));

                ReleaseComObject(targetOccurrence);
                targetOccurrence = null;
                ReleaseComObject(occurrences);
                occurrences = null;

                stage = "儲存測試組立";
                InvokeMethod(document, "Save", new object[0]);
                output.Add("[儲存] AssemblyDocument.Save() 已呼叫。");

                stage = "關閉已儲存的測試組立";
                InvokeMethod(document, "Close", new object[] { false });
                documentOpened = false;
                ReleaseComObject(document);
                document = null;

                stage = "重新背景開啟已儲存的測試組立";
                document = documents.GetType().InvokeMember("Open",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null, documents, new object[] { assemblyPath, (uint)8 });
                if (document == null) throw new InvalidOperationException("重新開啟時 Documents.Open 沒有回傳組立文件。");
                documentOpened = true;

                stage = "重新讀取已儲存的 occurrence";
                occurrences = GetProperty(document, "Occurrences");
                int reopenedCount = Convert.ToInt32(GetProperty(occurrences, "Count"));
                int persistedMatchCount = 0;
                string persistedPath = null;
                int? persistedRelations = null;

                for (int index = 1; index <= reopenedCount; index++)
                {
                    object candidate = null;
                    try
                    {
                        candidate = GetIndexedProperty(occurrences, "Item", new object[] { index });
                        bool isSubassembly = Convert.ToBoolean(GetProperty(candidate, "Subassembly"));
                        string candidatePath = Convert.ToString(GetProperty(candidate, "OccurrenceFileName"));
                        if (!isSubassembly && !String.IsNullOrWhiteSpace(candidatePath))
                        {
                            bool exactPath = false;
                            try
                            {
                                exactPath = String.Equals(Path.GetFullPath(candidatePath), expectedNewPath,
                                    StringComparison.OrdinalIgnoreCase);
                            }
                            catch { }
                            if (exactPath)
                            {
                                persistedMatchCount++;
                                persistedPath = candidatePath;
                                persistedRelations = TryGetRelationCount(candidate);
                            }
                        }
                    }
                    finally { ReleaseComObject(candidate); }
                }

                output.Add("[重新開啟後第一層 occurrence 數量] " + reopenedCount);
                output.Add("[重新開啟後新零件路徑命中數] " + persistedMatchCount);
                output.Add("[重新開啟後新零件路徑] " + (persistedPath ?? "(未找到)"));
                output.Add("[重新開啟後 Relations3d 數量] " +
                    (persistedRelations.HasValue ? persistedRelations.Value.ToString() : "無法讀取"));
                if (persistedMatchCount != 1)
                    throw new InvalidOperationException("重新開啟後，指定的新零件路徑命中數不是 1，無法確認替換結果已正確保存。");

                output.Add("[關聯檢查限制] Relations3d 數量只能作為初步比較，不能單獨證明所有配合關係均有效。");
                stage = "關閉重新開啟的測試組立";
                InvokeMethod(document, "Close", new object[] { false });
                documentOpened = false;
                output.Add("[完成] 已儲存並重新開啟驗證；文件已關閉。");
                return output;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("單一 occurrence 替換＋儲存重開測試失敗。檔案：" + assemblyPath +
                    "；失敗階段：" + stage +
                    (documentOpened ? "；文件可能仍開啟，請檢查 Solid Edge。" : "") +
                    "；原因：" + ex.Message, ex);
            }
            finally
            {
                ReleaseComObject(targetOccurrence);
                ReleaseComObject(occurrences);
                if (documentOpened && document != null)
                {
                    try { InvokeMethod(document, "Close", new object[] { false }); } catch { }
                }
                ReleaseComObject(document);
                ReleaseComObject(documents);
                if (createdApplication && application != null)
                {
                    try { InvokeMethod(application, "Quit", new object[0]); } catch { }
                }
                else if (visibilityCaptured && application != null)
                {
                    try { SetProperty(application, "Visible", originalVisible); } catch { }
                }
                ReleaseComObject(application);
            }
        }

        private static int? TryGetRelationCount(object occurrence)
        {
            object relations = null;
            try
            {
                relations = GetProperty(occurrence, "Relations3d");
                return Convert.ToInt32(GetProperty(relations, "Count"));
            }
            catch { return null; }
            finally { ReleaseComObject(relations); }
        }

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