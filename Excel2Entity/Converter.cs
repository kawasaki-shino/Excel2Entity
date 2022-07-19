using System;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;

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
						LogicalName = sheet.Cell("C6").Value.ToString(),
						PhysicsName = sheet.Cell("C7").Value.ToString(),
						ClassName = GenerateClassName(sheet.Cell("C7").Value.ToString())
					};

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
		public void OutputCs(string folder, string namespc)
		{
			var index = 0;

			foreach (var sheet in Book.Worksheets)
			{
				var list = new List<Columns>();

				// 目次は飛ばす
				if (sheet.Position == 1) continue;
				// 非表示のシートは飛ばす
				if (sheet.Visibility != XLWorksheetVisibility.Visible) continue;

				// 対象外なら次のシート
				if (!Files[index].Target)
				{
					index++;
					continue;
				}

				// カラムを読み込む
				for (var i = 14; i < 1000; i++)
				{
					var b = sheet.Cell($"B{i}").Value.ToString();
					var c = sheet.Cell($"C{i}").Value.ToString();
					var d = sheet.Cell($"D{i}").Value.ToString();
					var g = sheet.Cell($"G{i}").Value.ToString();

					// 論理名が空ならループを抜ける
					if (string.IsNullOrWhiteSpace(b)) break;

					var columns = new Columns()
					{
						LogicalName = b,
						PhysicsName = c,
						Type = string.IsNullOrWhiteSpace(d)
							? null
							: new string(d.Where(t => !char.IsControl(t)).ToArray()),
						Default = g
					};
					list.Add(columns);
				}


				var contents = $@"namespace {namespc}
{{
	public class {Files[index].ClassName}
	{{";

				foreach (var item in list)
				{
					contents += $@"
		/// <summary>{item.LogicalName}</summary>
		public {item.CsType} {item.CamelCasePhysicsName} {{ get; set; }}
";
				}

				contents += @"	}
}
";

				File.WriteAllText(Path.Combine(folder, $"{Files[index].ClassName}.cs"), contents);

				index++;
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
			return Columns.ToUpperCamelCase(words[1]);
		}
	}
}
