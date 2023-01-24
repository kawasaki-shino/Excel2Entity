using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static System.Int32;

namespace Excel2Entity
{
    public class Columns : EntityBase
    {
        /// <summary>論理名</summary>
        public string LogicalName { get; set; }

        private string _physicsName;
        /// <summary>物理名</summary>
        public string PhysicsName
        {
            get => _physicsName;
            set
            {
                _physicsName = value;
                CamelCasePhysicsName = GeneratePropertyName(_physicsName);
                PrivateVarName = GeneratePrivateVarName(_physicsName);
            }
        }

        /// <summary>物理名(キャメルケース)</summary>
        public string CamelCasePhysicsName { get; set; }

        /// <summary>private 変数名</summary>
        public string PrivateVarName { get; set; }

        private string _type;
        /// <summary>型</summary>
        public string Type
        {
            get => _type;
            set
            {
                _type = value;

                switch (_type)
                {
                    case "VARCHAR2":
                        CsType = typeof(string);
                        break;
                    case "CHAR":
                        CsType = typeof(char);
                        break;
                    case "DATE":
                        CsType = typeof(DateTime);
                        break;
                    case "NUMBER":
                        CsType = SetNumberType();
                        break;
                    default:
                        CsType = typeof(object);
                        break;
                }
            }
        }

        /// <summary>サイズ</summary>
        public string Size
        {
            get => _size;
            set
            {
                _size = value;
                CsType = SetNumberType();
            }
        }

        private string _size;

        /// <summary>サイズ(C#)</summary>
        public int? CsSize { get; set; }

        /// <summary>型(C#)</summary>
        public Type CsType { get; set; }

        /// <summary>必須</summary>
        private bool _required;

        public bool Required
        {
            get => _required;
            set
            {
                _required = value;
                OnPropertyChanged();
            }
        }

        /// <summary>Undo要否</summary>
        private bool _needUndo = true;

        public bool NeedUndo
        {
            get => _needUndo;
            set
            {
                _needUndo = value;
                OnPropertyChanged();
            }
        }

        public string Nullable => !Required && Type == "NUMBER"
            ? "?"
            : "";

        /// <summary>初期値</summary>
        public string Default { get; set; }

        /// <summary>
        /// プロパティ名生成
        /// </summary>
        /// <param name="physicsName"></param>
        /// <returns></returns>
        private string GeneratePropertyName(string physicsName)
        {
            // パース
            var words = physicsName.Split('_').ToList();
            var convertList = new List<string>();

            words.ForEach(n => convertList.Add(ToUpperCamelCase(n)));

            return string.Join("_", convertList);
        }

        /// <summary>
        /// private 変数名
        /// </summary>
        /// <param name="physicsName"></param>
        /// <returns></returns>
        private string GeneratePrivateVarName(string physicsName)
        {
            var sb = new StringBuilder();
            sb.Append('_');

            // パース
            var words = physicsName.Split('_');

            sb.Append(words.First().ToLower());

            foreach (var word in words.Skip(1))
            {
                sb.Append(ToUpperCamelCase(word));
            }

            return sb.ToString();
        }

        /// <summary>
        /// キャメルケース変換
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string ToUpperCamelCase(string s)
        {
            // 小文字に変換
            s = s.ToLower();

            var sb = new StringBuilder();
            var array = s.ToCharArray();

            // 先頭文字を大文字にする
            array[0] = char.ToUpper(array[0]);
            sb.Append(array);

            return sb.ToString();
        }

        /// <summary>
        /// NUMBER 型の場合に C# の型を何にするかの判定
        /// </summary>
        /// <returns></returns>
        private Type SetNumberType()
        {
            if (Size == null) return typeof(int);

            var arrSize = Size.Split(',');
            if (TryParse(arrSize[0], out int size))
            {
                CsSize = size;
            }

            if (Type != "NUMBER") return CsType;

            if (2 <= arrSize.Length)
            {
                // 小数部の桁指定有り
                return typeof(decimal);
            }

            return 10 <= CsSize
                ? typeof(long)
                : typeof(int);
        }
    }
}
