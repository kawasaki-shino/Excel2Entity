using ClosedXML.Excel;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace Excel2Entity
{
	public class Converter
	{
		public string FileInfo { get; set; }
		private XLWorkbook Book { get; set; }
		private ObservableCollection<Sheets> Files { get; set; }

		/// <summary>
		/// コンストラクタ
		/// </summary>
		public Converter(string file)
		{
			FileInfo = file;
		}

		/// <summary>
		/// Excel を読み込む
		/// </summary>
		public ObservableCollection<Sheets> LoadExcel()
		{
			Files = new ObservableCollection<Sheets>();

			try
			{
				// ファイル読み込み
				Book = new XLWorkbook(FileInfo);

				foreach (var sheet in Book.Worksheets)
				{
					// 目次は飛ばす
					if (sheet.Position == 1) continue;
					// 非表示のシートは飛ばす
					if (sheet.Visibility != XLWorksheetVisibility.Visible) continue;

					var table = new Sheets
					{
						LogicalName = CustomTrim(sheet.Cell("C6").Value.ToString()),
						PhysicsName = CustomTrim(sheet.Cell("C7").Value.ToString()),
						ClassName = GenerateClassName(CustomTrim(sheet.Cell("C7").Value.ToString()))
					};

					// カラムを読み込む
					for (var i = 14; i < 1000; i++)
					{
						var b = CustomTrim(sheet.Cell($"B{i}").Value.ToString());
						var c = CustomTrim(sheet.Cell($"C{i}").Value.ToString());
						var d = CustomTrim(sheet.Cell($"D{i}").Value.ToString());
						var g = CustomTrim(sheet.Cell($"G{i}").Value.ToString());
						var j = CustomTrim(sheet.Cell($"J{i}").Value.ToString());

						// 論理名が空ならループを抜ける
						if (string.IsNullOrWhiteSpace(b)) break;

						table.ColumnsList.Add(new Columns
						{
							LogicalName = b,
							PhysicsName = c,
							Type = d,
							Default = g,
							Required = j == "○"
						});
					}

					Files.Add(table);
				}
			}
			catch (IOException e)
			{
				Console.WriteLine(e);
				return null;
			}

			return Files;
		}

		/// <summary>
		/// Entity クラスファイル出力
		/// </summary>
		/// <param name="folder"></param>
		/// <param name="namespc"></param>
		/// <param name="isInheritNotificationObject"></param>
		public void OutputCs(string folder, string namespc, bool isInheritNotificationObject)
		{
			folder = GetOutputFolder(folder);

			foreach (var file in Files)
			{
				// 対象外なら次のシート
				if (!file.Target)
				{
					continue;
				}

				var contents = $@"using System;{(isInheritNotificationObject ? "\r\nusing Wiseman.PJC.WPF.ObjectModel;" : "")}

namespace {namespc}
{{
	public class {file.ClassName}{(isInheritNotificationObject ? " : NotificationObject" : "")}
	{{";

				foreach (var item in file.ColumnsList)
				{
					if (isInheritNotificationObject)
					{
						contents += $@"
		private {item.CsType.GetAliasName()}{GetNullable(item.Required, item.CsType)} {item.PrivateVarName}{GetDefaultString(item.CsType, item.Default, true)}

		/// <summary>{item.LogicalName}</summary>
		public {item.CsType.GetAliasName()}{GetNullable(item.Required, item.CsType)} {item.CamelCasePhysicsName}
		{{
			get => {item.PrivateVarName};
			set
			{{
				{item.PrivateVarName} = value;
				RaisePropertyChanged();
			}}
		}}
";
					}
					else
					{
						contents += $@"
		/// <summary>{item.LogicalName}</summary>
		public {item.CsType.GetAliasName()}{GetNullable(item.Required, item.CsType)} {item.CamelCasePhysicsName} {{ get; set; }}{GetDefaultString(item.CsType, item.Default, false)}
";
					}
				}

				contents += @"	}
}
";

				File.WriteAllText(Path.Combine(folder, $"{file.ClassName}.cs"), contents);
			}
		}

		/// <summary>
		/// 物理名をパースしてクラス名の候補を作成
		/// </summary>
		/// <param name="physicsName"></param>
		/// <returns></returns>
		private string GenerateClassName(string physicsName)
		{
			// パースする
			var words = physicsName.Split('_');
			if (words.Length < 2) return "";

			// 任意名称部分の先頭文字を大文字にして返却
			return $"{Columns.ToUpperCamelCase(words[1])}Entity";
		}

		/// <summary>
		/// Excel から取得した文字列中の余計な文字を抜く
		/// </summary>
		/// <param name="value"></param>
		/// <returns></returns>
		private string CustomTrim(string value)
		{
			// 空白を抜く
			value = value.Trim().Trim('\u200B');
			// 制御文字を抜く
			return new string(value.Where(c => !char.IsControl(c)).ToArray());
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="itemType"></param>
		/// <param name="itemDefault"></param>
		/// <param name="isInheritNotificationObject"></param>
		/// <returns></returns>
		private string GetDefaultString(Type itemType, string itemDefault, bool isInheritNotificationObject)
		{
			// 初期値指定なしなら抜ける
			if (string.IsNullOrWhiteSpace(itemDefault)) return isInheritNotificationObject
				? ";"
				: "";

			// 文字列意外は初期値をそのまま出力
			if (itemType != typeof(string))
			{
				var value = itemDefault.Replace("'", "");
				return string.IsNullOrEmpty(value)
					? isInheritNotificationObject
						? ";"
						: ""
					: $" = {value};";
			}

			// 文字列型かつ空が初期値
			if (itemDefault == "''" || itemDefault == "'''") return @" = """";";
			// それ以外はダブルクオーテーションで囲って出力
			return $@" = ""{itemDefault.Replace("'", "")}"";";
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="required"></param>
		/// <param name="csType"></param>
		/// <returns></returns>
		private string GetNullable(bool required, Type csType)
		{
			if (required) return "";

			if (csType.GetAliasName() == typeof(decimal).GetAliasName()) return "?";

			return "";
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="folder"></param>
		/// <returns></returns>
		private string GetOutputFolder(string folder)
		{
			// 拡張子なしのファイル名を取得
			var name = Path.GetFileNameWithoutExtension(FileInfo);
			var path = folder.Contains(name)
				? folder
				: Path.Combine(folder, name);

			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}

			return path;
		}
	}
}
