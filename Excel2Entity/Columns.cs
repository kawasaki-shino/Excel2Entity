using System;
using System.Text;

namespace Excel2Entity
{
	public class Columns
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
			}
		}

		/// <summary>物理名(キャメルケース)</summary>
		public string CamelCasePhysicsName { get; set; }

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
		public bool Required { get; set; }

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
			foreach (var word in words)
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
