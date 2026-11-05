using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace SEBatchReplace
{
    // IDispatch v-table declaration used only to obtain the live object's ITypeInfo.
    [ComImport]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDispatchTypeInfo
    {
        [PreserveSig]
        int GetTypeInfoCount(out uint count);

        [PreserveSig]
        int GetTypeInfo(uint index, int lcid, out ITypeInfo typeInfo);
    }

    internal static class ComTypeInfoInspector
    {
        public static IList<string> Describe(object comObject, string label)
        {
            var lines = new List<string>();
            if (comObject == null)
            {
                lines.Add("[TypeInfo] " + label + "：COM 物件為 null。");
                return lines;
            }

            ITypeInfo typeInfo = null;
            IntPtr typeAttrPointer = IntPtr.Zero;
            try
            {
                IDispatchTypeInfo dispatch = (IDispatchTypeInfo)comObject;
                uint typeInfoCount;
                int hr = dispatch.GetTypeInfoCount(out typeInfoCount);
                if (hr < 0)
                    Marshal.ThrowExceptionForHR(hr);
                if (typeInfoCount == 0)
                {
                    lines.Add("[TypeInfo] " + label + "：此 IDispatch 物件未提供 TypeInfo。");
                    return lines;
                }

                hr = dispatch.GetTypeInfo(0, 0, out typeInfo);
                if (hr < 0)
                    Marshal.ThrowExceptionForHR(hr);
                if (typeInfo == null)
                {
                    lines.Add("[TypeInfo] " + label + "：GetTypeInfo 回傳 null。");
                    return lines;
                }

                TYPEATTR typeAttr;
                typeInfo.GetTypeAttr(out typeAttrPointer);
                typeAttr = (TYPEATTR)Marshal.PtrToStructure(typeAttrPointer, typeof(TYPEATTR));
                string typeName, typeDoc, helpFile;
                int helpContext;
                typeInfo.GetDocumentation(-1, out typeName, out typeDoc, out helpContext, out helpFile);
                lines.Add("  [介面] " + (String.IsNullOrEmpty(typeName) ? label : typeName) +
                    "；GUID=" + typeAttr.guid + "；函式描述數=" + typeAttr.cFuncs);

                for (int i = 0; i < typeAttr.cFuncs; i++)
                {
                    IntPtr funcDescPointer = IntPtr.Zero;
                    try
                    {
                        typeInfo.GetFuncDesc(i, out funcDescPointer);
                        FUNCDESC func = (FUNCDESC)Marshal.PtrToStructure(funcDescPointer, typeof(FUNCDESC));
                        int nameCapacity = Math.Max(1, Math.Min(65, func.cParams + 1));
                        string[] names = new string[nameCapacity];
                        int namesReturned;
                        typeInfo.GetNames(func.memid, names, names.Length, out namesReturned);

                        string memberName = namesReturned > 0 && !String.IsNullOrEmpty(names[0])
                            ? names[0] : "(未知成員)";
                        string invokeKind = func.invkind.ToString();
                        var parameters = new List<string>();
                        for (int p = 1; p < namesReturned; p++)
                            parameters.Add(names[p]);
                        lines.Add("  " + invokeKind + " " + memberName +
                            "；DISPID=" + func.memid +
                            "；參數數=" + func.cParams +
                            (parameters.Count > 0 ? "；參數=" + String.Join(", ", parameters.ToArray()) : String.Empty));
                    }
                    catch (Exception ex)
                    {
                        lines.Add("  [函式描述讀取失敗] index=" + i + "；" + ex.Message);
                    }
                    finally
                    {
                        if (funcDescPointer != IntPtr.Zero)
                            typeInfo.ReleaseFuncDesc(funcDescPointer);
                    }
                }
            }
            catch (Exception ex)
            {
                lines.Add("[TypeInfo 讀取失敗] " + label + "；" + ex.Message);
            }
            finally
            {
                if (typeAttrPointer != IntPtr.Zero && typeInfo != null)
                    typeInfo.ReleaseTypeAttr(typeAttrPointer);
                if (typeInfo != null && Marshal.IsComObject(typeInfo))
                    Marshal.ReleaseComObject(typeInfo);
            }
            return lines;
        }
    }
}
