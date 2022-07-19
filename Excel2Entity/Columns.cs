using System.Text;

namespace Excel2Entity
{
	public class Columns
	{
		/// <summary>論理名</summary>
		public string LogicalName { get; set; }

		/// <summary>物理名</summary>
		public string PhysicsName { get; set; }

		/// <summary>物理名(キャメルケース)</summary>
		public string CamelCasePhysicsName => GeneratePropertyName(PhysicsName);

		/// <summary>型</summary>
		public string Type { get; set; }

		/// <summary>型(C#)</summary>
		public string CsType
		{
			get
			{
				// 単純な文字列比較だと引っかからないので Contains で判定する(制御文字を抜ききれていない？)
				if (Type.Contains("VARCHAR2")) return "string";
				if (Type.Contains("CHAR")) return "string";
				if (Type.Contains("DATE")) return "DateTime";
				if (Type.Contains("NUMBER")) return "decimal";
				return "object";
			}
		}

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
