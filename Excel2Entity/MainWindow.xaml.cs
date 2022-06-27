using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Windows;

namespace Excel2Entity
{
	/// <summary>
	/// MainWindow.xaml の相互作用ロジック
	/// </summary>
	public partial class MainWindow : Window
	{
		public Converter Converter { get; set; }

		public MainWindow()
		{
			InitializeComponent();

			Loaded += (s, e) =>
			{
				TbxFolder.Text = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
			};

			// DragOver イベントハンドラ
			TbxExcel.PreviewDragOver += (s, e) =>
			{
				// マウスポインタの変更
				e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
					? DragDropEffects.Copy
					: DragDropEffects.None;

				e.Handled = true;
			};

			// Drop イベントハンドラ
			TbxExcel.Drop += (s, e) =>
			{
				if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

				var files = (string[])e.Data.GetData(DataFormats.FileDrop);
				if (files == null) return;

				TbxExcel.Text = files[0];
				LoadExcel(files[0]);
			};

			// ファイルを開くボタンクリックイベントハンドラ
			BtnOfd.Click += (s, e) =>
			{
				var dialog = new OpenFileDialog
				{
					Filter = "エクセルファイル（*.xlsx）|*.xlsx"
				};

				if (dialog.ShowDialog() == true)
				{
					TbxExcel.Text = dialog.FileName;
					LoadExcel(dialog.FileName);
				}
			};

			// フォルダ選択ボタンクリックイベントハンドラ
			BtnFolder.Click += (s, e) =>
			{
				// フォルダ選択モードで開く
				using (var dialog = new CommonOpenFileDialog()
				{
					IsFolderPicker = true
				})
				{
					if (dialog.ShowDialog() != CommonFileDialogResult.Ok) return;

					TbxFolder.Text = dialog.FileName;
				}
			};

			// 出力ボタンクリック
			BtnOutput.Click += (s, e) =>
			{
				if (string.IsNullOrWhiteSpace(TbxFolder.Text)) return;

				Converter.OutputCs(TbxFolder.Text, TbxNamespace.Text);


				MessageBox.Show(this, "出力が完了しました", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
			};
		}

		/// <summary>
		/// Excel 読み込み
		/// </summary>
		/// <param name="file"></param>
		void LoadExcel(string file)
		{
			var extention = Path.GetExtension(file);
			if (extention != ".xlsx")
			{
				MessageBox.Show(this, "Excel ファイルのみ読み込み可能です", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
				return;
			}

			Converter = new Converter(file);

			// Excel を読み込んで DataGrid にバインド
			var excel = Converter.LoadExcel();
			DgClass.ItemsSource = excel;

			if (excel == null) MessageBox.Show(this, "Excel ファイルが開かれています。閉じてから再度取込を行ってください。", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
		}
	}
}
