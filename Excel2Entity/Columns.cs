using System;
using System.Linq;
using System.Text;

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
                        CsType = typeof(string);
                        break;
                    case "DATE":
                        CsType = typeof(DateTime);
                        break;
                    case "NUMBER":
                        CsType = typeof(decimal);
                        break;
                    default:
                        CsType = typeof(object);
                        break;
                }
            }
        }

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
            var sb = new StringBuilder();

            // パース
            var words = physicsName.Split('_');
            Array.ForEach(words, x => sb.Append(ToUpperCamelCase(x)));

            return sb.ToString();
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
    }
}
