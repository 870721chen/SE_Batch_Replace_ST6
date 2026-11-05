using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SEBatchReplace
{
    public static class RevisionManagerReader
    {
        public static IList<string> ReadLinkedDocuments(string filePath)
        {
            if (String.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("找不到指定的組立檔。", filePath);

            object application = null;
            object document = null;
            object linkedDocuments = null;
            bool createdApplication = false;
            string stage = "連接 Revision Manager";

            try
            {
                Type applicationType = Type.GetTypeFromProgID("RevisionManager.Application");
                if (applicationType == null)
                    throw new InvalidOperationException("系統未註冊 RevisionManager.Application。請確認 Solid Edge ST6 的 Revision Manager 元件已安裝並註冊。");

                try
                {
                    application = Marshal.GetActiveObject("RevisionManager.Application");
                }
                catch
                {
                    stage = "啟動 Revision Manager";
                    application = Activator.CreateInstance(applicationType);
                    createdApplication = true;
                }

                stage = "透過 Revision Manager 開啟組立文件";
                document = InvokeMethod(application, "Open", new object[] { filePath });
                if (document == null)
                    throw new InvalidOperationException("Revision Manager.Open 沒有回傳文件物件。");

                stage = "取得 LinkedDocuments";
                linkedDocuments = GetProperty(document, "LinkedDocuments");
                int count = Convert.ToInt32(GetProperty(linkedDocuments, "Count"));

                var results = new List<string>();
                results.Add("Revision Manager 直接連結數：" + count);

                for (int index = 1; index <= count; index++)
                {
                    object linkedDocument = null;
                    try
                    {
                        linkedDocument = GetIndexedProperty(linkedDocuments, "Item", new object[] { index });
                        string fullName = Convert.ToString(GetProperty(linkedDocument, "FullName"));
                        string displayName = String.IsNullOrWhiteSpace(fullName)
                            ? "(無法取得 FullName)"
                            : Path.GetFileName(fullName);

                        results.Add(index.ToString() + ". " + displayName + " | " + fullName);
                    }
                    catch (Exception ex)
                    {
                        results.Add(index.ToString() + ". [讀取連結失敗] " + ex.Message);
                    }
                    finally
                    {
                        ReleaseComObject(linkedDocument);
                    }
                }

                return results;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "Revision Manager 讀取失敗。檔案：" + filePath +
                    "；失敗階段：" + stage + "；原因：" + ex.Message, ex);
            }
            finally
            {
                ReleaseComObject(linkedDocuments);
                ReleaseComObject(document);

                // Only quit the Revision Manager instance created by this test.
                // An instance that was already running is left untouched.
                if (createdApplication && application != null)
                {
                    try { InvokeMethod(application, "Quit", new object[0]); }
                    catch { }
                }
                ReleaseComObject(application);
            }
        }

        private static object GetProperty(object target, string propertyName)
        {
            return target.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                null,
                target,
                null);
        }

        private static object GetIndexedProperty(object target, string propertyName, object[] arguments)
        {
            return target.GetType().InvokeMember(
                propertyName,
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                null,
                target,
                arguments);
        }

        private static object InvokeMethod(object target, string methodName, object[] arguments)
        {
            return target.GetType().InvokeMember(
                methodName,
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                null,
                target,
                arguments);
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
