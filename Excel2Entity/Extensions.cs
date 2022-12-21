using System;

namespace Excel2Entity
{
    /// <summary>
    /// 拡張メソッド
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// エイリアス名取得
        /// </summary>
        /// <param name="self"></param>
        /// <returns></returns>
        public static string GetAliasName(this Type self)
        {
            switch (self.FullName)
            {
                case "System.Boolean": return "bool";
                case "System.Byte": return "byte";
                case "System.SByte": return "sbyte";
                case "System.Char": return "char";
                case "System.Decimal": return "decimal";
                case "System.Double": return "double";
                case "System.Single": return "float";
                case "System.Int32": return "int";
                case "System.UInt32": return "uint";
                case "System.Int64": return "long";
                case "System.UInt64": return "ulong";
                case "System.Object": return "object";
                case "System.Int16": return "short";
                case "System.UInt16": return "ushort";
                case "System.String": return "string";
                case "System.Void": return "void";
            }

            return self.Name;
        }
    }
}
