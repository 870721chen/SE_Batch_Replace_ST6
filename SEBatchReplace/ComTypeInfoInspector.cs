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

                System.Runtime.InteropServices.ComTypes.TYPEATTR typeAttr;
                typeInfo.GetTypeAttr(out typeAttrPointer);
                typeAttr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)Marshal.PtrToStructure(typeAttrPointer, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));
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
                        System.Runtime.InteropServices.ComTypes.FUNCDESC func = (System.Runtime.InteropServices.ComTypes.FUNCDESC)Marshal.PtrToStructure(funcDescPointer, typeof(System.Runtime.InteropServices.ComTypes.FUNCDESC));
                        int nameCapacity = Math.Max(1, Math.Min(65, func.cParams + 1));
                        string[] names = new string[nameCapacity];
                        int namesReturned;
                        typeInfo.GetNames(func.memid, names, names.Length, out namesReturned);

                        string memberName = namesReturned > 0 && !String.IsNullOrEmpty(names[0])
                            ? names[0] : "(未知成員)";
                        string invokeKind = func.invkind.ToString();
                        string returnType = DescribeType(func.elemdescFunc.tdesc);
                        var parameters = new List<string>();
                        int elementSize = Marshal.SizeOf(typeof(System.Runtime.InteropServices.ComTypes.ELEMDESC));
                        for (int p = 0; p < func.cParams; p++)
                        {
                            System.Runtime.InteropServices.ComTypes.ELEMDESC element =
                                (System.Runtime.InteropServices.ComTypes.ELEMDESC)Marshal.PtrToStructure(
                                    IntPtr.Add(func.lprgelemdescParam, p * elementSize),
                                    typeof(System.Runtime.InteropServices.ComTypes.ELEMDESC));
                            string parameterName = p + 1 < namesReturned && !String.IsNullOrEmpty(names[p + 1])
                                ? names[p + 1] : "arg" + (p + 1);
                            parameters.Add(parameterName + ":" + DescribeType(element.tdesc) +
                                " [" + DescribeParamFlags(element.desc.paramdesc.wParamFlags) + "]");
                        }
                        lines.Add("  " + invokeKind + " " + memberName +
                            "；DISPID=" + func.memid +
                            "；回傳型別=" + returnType +
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

        private static string DescribeType(System.Runtime.InteropServices.ComTypes.TYPEDESC type)
        {
            string name;
            switch ((VarEnum)type.vt)
            {
                case VarEnum.VT_VOID: name = "VOID"; break;
                case VarEnum.VT_I2: name = "I2"; break;
                case VarEnum.VT_I4: name = "I4"; break;
                case VarEnum.VT_R4: name = "R4"; break;
                case VarEnum.VT_R8: name = "R8"; break;
                case VarEnum.VT_BSTR: name = "BSTR"; break;
                case VarEnum.VT_BOOL: name = "BOOL"; break;
                case VarEnum.VT_VARIANT: name = "VARIANT"; break;
                case VarEnum.VT_DISPATCH: name = "IDispatch"; break;
                case VarEnum.VT_UNKNOWN: name = "IUnknown"; break;
                case VarEnum.VT_PTR: name = "PTR"; break;
                case VarEnum.VT_SAFEARRAY: name = "SAFEARRAY"; break;
                case VarEnum.VT_USERDEFINED: name = "USERDEFINED"; break;
                default: name = ((VarEnum)type.vt).ToString(); break;
            }
            return name + "(0x" + type.vt.ToString("X4") + ")";
        }

        private static string DescribeParamFlags(System.Runtime.InteropServices.ComTypes.PARAMFLAG flags)
        {
            return flags == 0 ? "無旗標" : flags.ToString();
        }
    }
}
