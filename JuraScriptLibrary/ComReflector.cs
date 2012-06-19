using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace JuraScriptLibrary
{
    public static class ComReflector
    {
        public enum ComReflectedMemberTypes {
            NO_ITYPEINFO,
            NOT_FOUND,
            Method,
            Property,
            Field
        }

        public static object InvokeMember(object comObject, string memberName, object[] arguments, bool set)
        {
            ComReflectedMemberTypes memberType = GetMemberType(comObject, memberName);
            if (memberType == ComReflectedMemberTypes.Field)
            {
                return comObject.GetType().InvokeMember(memberName, set ? System.Reflection.BindingFlags.SetField : System.Reflection.BindingFlags.GetField, null, comObject, arguments);
            }
            else if (memberType == ComReflectedMemberTypes.Property)
            {
                return comObject.GetType().InvokeMember(memberName, set ? System.Reflection.BindingFlags.SetProperty : System.Reflection.BindingFlags.GetProperty, null, comObject, arguments);
            }
            else if (memberType == ComReflectedMemberTypes.Method)
            {
                return comObject.GetType().InvokeMember(memberName, System.Reflection.BindingFlags.InvokeMethod, null, comObject, arguments);
            }
            else
            {
                return null;
            }
        }

        public static ComReflectedMemberTypes GetCOMMemberType(string memberName, ITypeInfo typeInfo, System.Runtime.InteropServices.ComTypes.TYPEATTR typeAttr)
        {
            ComReflectedMemberTypes returnType = ComReflectedMemberTypes.NOT_FOUND;
            //Methods/subroutines/properties
            for (int i = 0; i < typeAttr.cFuncs; i++)
            {
                IntPtr pFuncDesc = new IntPtr();
                typeInfo.GetFuncDesc(i, out pFuncDesc);
                System.Runtime.InteropServices.ComTypes.FUNCDESC funcDesc = (System.Runtime.InteropServices.ComTypes.FUNCDESC)Marshal.PtrToStructure(pFuncDesc, typeof(System.Runtime.InteropServices.ComTypes.FUNCDESC));

                string[] names = { "" };
                int pcNames = 0;
                typeInfo.GetNames(funcDesc.memid, names, 1, out pcNames);
                string funcName = names[0];

                if (funcName.ToLower() == memberName.ToLower())
                {
                    if (((VarEnum)Enum.Parse(typeof(VarEnum), funcDesc.elemdescFunc.tdesc.vt.ToString()) == VarEnum.VT_VOID))
                    {
                        returnType = ComReflectedMemberTypes.Method;
                    }
                    if ((funcDesc.invkind & System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYGET) != 0)
                    {
                        returnType = ComReflectedMemberTypes.Property;
                    }
                    if ((funcDesc.invkind & System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYPUT) != 0)
                    {
                        returnType = ComReflectedMemberTypes.Property;
                    }
                    if ((funcDesc.invkind & System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYPUTREF) != 0)
                    {
                        returnType = ComReflectedMemberTypes.Property;
                    }
                    if (returnType == ComReflectedMemberTypes.NOT_FOUND)
                    {
                        returnType = ComReflectedMemberTypes.Method;
                    }
                }

                typeInfo.ReleaseFuncDesc(pFuncDesc);
            }

            if (returnType == ComReflectedMemberTypes.NOT_FOUND)
            {
                //Field members
                for (int iVar = 0; iVar < typeAttr.cVars; iVar++)
                {
                    IntPtr pVarDesc = new IntPtr();
                    typeInfo.GetVarDesc(iVar, out pVarDesc);
                    System.Runtime.InteropServices.ComTypes.VARDESC varDesc = (System.Runtime.InteropServices.ComTypes.VARDESC)Marshal.PtrToStructure(pVarDesc, typeof(System.Runtime.InteropServices.ComTypes.VARDESC));
                    string[] varNames = new string[] { "" };
                    int varPcNames = 0;
                    typeInfo.GetNames(varDesc.memid, varNames, 1, out varPcNames);
                    string varName = varNames[0];

                    if (varName.ToLower() == memberName.ToLower())
                    {
                        returnType = ComReflectedMemberTypes.Field;
                    }
                }
            }

            return returnType;
        }

        public static int GetDispId(IDispatch idisp, string memberName)
        {
            Guid dummy = Guid.Empty;
            string[] rgsNames = new string[1];
            rgsNames[0] = memberName;
            int[] rgDispId = new int[1];
            idisp.GetIDsOfNames(ref dummy, rgsNames, 1, 0x800, rgDispId);
            return rgDispId[0];

        }

        public static string DefaultPropertyName(IDispatch idisp, out ComReflectedMemberTypes memberType)
        {
            string returnName = "";

            if (idisp == null)
            {
                memberType = ComReflectedMemberTypes.NOT_FOUND;
                return null;
            }

            UInt32 count = 0;
            idisp.GetTypeInfoCount(ref count);
            if (count < 1) { memberType = ComReflectedMemberTypes.NO_ITYPEINFO; return null; }
            IntPtr _typeInfo = new IntPtr();
            idisp.GetTypeInfo(0, 0, ref _typeInfo);
            if (_typeInfo == IntPtr.Zero)
            {
                memberType = ComReflectedMemberTypes.NO_ITYPEINFO;
                return null;
            }

            ITypeInfo typeInfo = (ITypeInfo)Marshal.GetTypedObjectForIUnknown(_typeInfo, typeof(ITypeInfo));
            Marshal.Release(_typeInfo);

            //AddTypeInfoToDump
            string typeName = "";
            string docString = "";
            int docInt = 0;
            string helpFile = "";
            typeInfo.GetDocumentation(-1, out typeName, out docString, out docInt, out helpFile);

            IntPtr pTypeAttr = new IntPtr();
            typeInfo.GetTypeAttr(out pTypeAttr);

            System.Runtime.InteropServices.ComTypes.TYPEATTR typeAttr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)Marshal.PtrToStructure(pTypeAttr, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));

            ComReflectedMemberTypes returnType = ComReflectedMemberTypes.NOT_FOUND;
            //Methods/subroutines/properties
            int dispId = -1;
            
            for (int i = 0; i < typeAttr.cFuncs; i++)
            {
                IntPtr pFuncDesc = new IntPtr();
                typeInfo.GetFuncDesc(i, out pFuncDesc);
                System.Runtime.InteropServices.ComTypes.FUNCDESC funcDesc = (System.Runtime.InteropServices.ComTypes.FUNCDESC)Marshal.PtrToStructure(pFuncDesc, typeof(System.Runtime.InteropServices.ComTypes.FUNCDESC));

                string[] names = { "" };
                int pcNames = 0;
                typeInfo.GetNames(funcDesc.memid, names, 1, out pcNames);
                string funcName = names[0];

                dispId = GetDispId(idisp, funcName);
                if (dispId == 0)
                {
                    if (((VarEnum)Enum.Parse(typeof(VarEnum), funcDesc.elemdescFunc.tdesc.vt.ToString()) == VarEnum.VT_VOID))
                    {
                        returnType = ComReflectedMemberTypes.Method;
                    }
                    if ((funcDesc.invkind & System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYGET) != 0)
                    {
                        returnType = ComReflectedMemberTypes.Property;
                        returnName = funcName;
                    }
                    if ((funcDesc.invkind & System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYPUT) != 0)
                    {
                        returnType = ComReflectedMemberTypes.Property;
                        returnName = funcName;
                    }
                    if ((funcDesc.invkind & System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYPUTREF) != 0)
                    {
                        returnType = ComReflectedMemberTypes.Property;
                        returnName = funcName;
                    }
                    if (returnType == ComReflectedMemberTypes.NOT_FOUND)
                    {
                        returnType = ComReflectedMemberTypes.Method;
                    }
                }

                typeInfo.ReleaseFuncDesc(pFuncDesc);
            }

            if (returnType == ComReflectedMemberTypes.NOT_FOUND)
            {
                //Field members
                for (int iVar = 0; iVar < typeAttr.cVars; iVar++)
                {
                    IntPtr pVarDesc = new IntPtr();
                    typeInfo.GetVarDesc(iVar, out pVarDesc);
                    System.Runtime.InteropServices.ComTypes.VARDESC varDesc = (System.Runtime.InteropServices.ComTypes.VARDESC)Marshal.PtrToStructure(pVarDesc, typeof(System.Runtime.InteropServices.ComTypes.VARDESC));
                    string[] varNames = new string[] { "" };
                    int varPcNames = 0;
                    typeInfo.GetNames(varDesc.memid, varNames, 1, out varPcNames);
                    string varName = varNames[0];

                    dispId = GetDispId(idisp, varName);
                    if (dispId == 0)
                    {
                        //return properties only
                        //returnName = varName;
                        
                        returnType = ComReflectedMemberTypes.Field;
                    }
                }
            }
            memberType = returnType;
            return returnName;
        }


        public static string GetCOMObjectTypeName(object comObject)
        {
            if (comObject == null) { return ""; }
            var idisp = comObject as IDispatch;
            if (idisp == null) { return ""; }

            UInt32 count = 0;
            idisp.GetTypeInfoCount(ref count);
            if (count < 1) { return ""; }
            IntPtr _typeInfo = new IntPtr();
            idisp.GetTypeInfo(0, 0, ref _typeInfo);
            if (_typeInfo == IntPtr.Zero)
            {
                return "";
            }

            ITypeInfo typeInfo = (ITypeInfo)Marshal.GetTypedObjectForIUnknown(_typeInfo, typeof(ITypeInfo));
            Marshal.Release(_typeInfo);

            //AddTypeInfoToDump
            string typeName = "";
            string docString = "";
            int docInt = 0;
            string helpFile = "";
            typeInfo.GetDocumentation(-1, out typeName, out docString, out docInt, out helpFile);

            return typeName;
        }

        public static ComReflectedMemberTypes GetMemberType(object comObject, string memberName)
        {
            if (comObject == null || String.IsNullOrWhiteSpace(memberName))
            {
                return ComReflectedMemberTypes.NOT_FOUND;
            }

            var idisp = (IDispatch)comObject;
            UInt32 count = 0;
            idisp.GetTypeInfoCount(ref count);
            if (count < 1) { return ComReflectedMemberTypes.NO_ITYPEINFO; }
            IntPtr _typeInfo = new IntPtr();
            idisp.GetTypeInfo(0, 0, ref _typeInfo);
            if (_typeInfo == IntPtr.Zero)
            {
                return ComReflectedMemberTypes.NO_ITYPEINFO;
            }

            ITypeInfo typeInfo = (ITypeInfo)Marshal.GetTypedObjectForIUnknown(_typeInfo, typeof(ITypeInfo));
            Marshal.Release(_typeInfo);

            //AddTypeInfoToDump
            string typeName = "";
            string docString = "";
            int docInt = 0;
            string helpFile = "";
            typeInfo.GetDocumentation(-1, out typeName, out docString, out docInt, out helpFile);

            IntPtr pTypeAttr = new IntPtr();
            typeInfo.GetTypeAttr(out pTypeAttr);
            
            System.Runtime.InteropServices.ComTypes.TYPEATTR typeAttr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)Marshal.PtrToStructure(pTypeAttr, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));

            return GetCOMMemberType(memberName, typeInfo, typeAttr);
        }

        public static string DumpTypeDesc(System.Runtime.InteropServices.ComTypes.TYPEDESC tdesc, System.Runtime.InteropServices.ComTypes.ITypeInfo context)
        {
            System.Runtime.InteropServices.ComTypes.TYPEDESC tdesc2;
            switch ((VarEnum)tdesc.vt)
            {
                case VarEnum.VT_PTR:
                    tdesc2 = (System.Runtime.InteropServices.ComTypes.TYPEDESC)Marshal.PtrToStructure(tdesc.lpValue, typeof(System.Runtime.InteropServices.ComTypes.TYPEDESC));
                    return "ref " + DumpTypeDesc(tdesc2, context);
                case VarEnum.VT_USERDEFINED:
                    int href = tdesc.lpValue.ToInt32();
                    System.Runtime.InteropServices.ComTypes.ITypeInfo refTypeInfo = null;
                    context.GetRefTypeInfo(href, out refTypeInfo);
                    string refTypeName = "";
                    string refDocString = "";
                    string refHelpDoc = "";
                    int refHelpInt = 0;
                    refTypeInfo.GetDocumentation(-1, out refTypeName, out refDocString, out refHelpInt, out refHelpDoc);
                    return refTypeName;
                    break;
                case VarEnum.VT_CARRAY:
                    tdesc2 = (System.Runtime.InteropServices.ComTypes.TYPEDESC)Marshal.PtrToStructure(tdesc.lpValue, typeof(System.Runtime.InteropServices.ComTypes.TYPEDESC));
                    return "array of " + DumpTypeDesc(tdesc2, context);
                case VarEnum.VT_VOID:
                    return "void";
                case VarEnum.VT_VARIANT:
                    return "Object";
                    break;
                case VarEnum.VT_UNKNOWN:
                    return "IUnknown";
                case VarEnum.VT_BSTR:
                    return "String";
                case VarEnum.VT_LPSTR:
                    return "char*";
                case VarEnum.VT_LPWSTR:
                    return "wchar*";
                case VarEnum.VT_HRESULT:
                    return "HResult";
                case VarEnum.VT_BOOL:
                    return "bool";
                case VarEnum.VT_I1:
                    //SByte
                    return "SByte";
                case VarEnum.VT_I2:
                    return "short";
                case VarEnum.VT_UI1:
                    return "byte";
                case VarEnum.VT_UI2:
                    return "ushort";
                case VarEnum.VT_I4:
                    return "int32";
                case VarEnum.VT_INT:
                    return "int32";
                case VarEnum.VT_UI4:
                    return "uint32";
                case VarEnum.VT_I8:
                    return "long";
                case VarEnum.VT_UI8:
                    return "ulong";

            }
            return ((VarEnum)tdesc.vt).ToString();
        }
    }
}
