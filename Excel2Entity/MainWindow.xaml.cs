using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace Excel2Entity
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public Converter Converter { get; set; }

        public ObservableCollection<CsTypes> CsTypes { get; set; } = new ObservableCollection<CsTypes>()
        {
            new CsTypes(name: typeof(string).GetAliasName(), value: typeof(string)),
            new CsTypes(name: typeof(decimal).GetAliasName(), value: typeof(decimal)),
            new CsTypes(name: typeof(DateTime).GetAliasName(), value: typeof(DateTime)),
            new CsTypes(name: typeof(object).GetAliasName(), value: typeof(object)),
        };

        public MainWindow()
        {
            InitializeComponent();

            DataContext = this;

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

                TbxExcel.Text = LoadExcel(files[0])
                    ? files[0]
                    : "";
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
                    TbxExcel.Text = LoadExcel(dialog.FileName)
                        ? dialog.FileName
                        : "";
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

                Converter.OutputCs(TbxFolder.Text, TbxNamespace.Text, Chk.IsChecked ?? false);

                MessageBox.Show(this, "出力が完了しました", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            };

            // クラス表示 DataGird SelectedCellsChanged
            DgClass.SelectedCellsChanged += (s, e) =>
            {
                // 選択アイテムを取得
                var item = (Sheets)DgClass.SelectedItem;

                if (item?.ColumnsList != null) DgColumn.ItemsSource = item.ColumnsList;

                SetScrollTop(DgColumn);
            };
        }

        /// <summary>
        /// Excel 読み込み
        /// </summary>
        /// <param name="file"></param>
        private bool LoadExcel(string file)
        {
            var extention = Path.GetExtension(file);
            if (extention != ".xlsx")
            {
                MessageBox.Show(this, "Excel ファイルのみ読み込み可能です", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            Converter = new Converter(file);

            // Excel を読み込んで DataGrid にバインド
            var excel = Converter.LoadExcel();
            DgClass.ItemsSource = excel;
            SetScrollTop(DgClass);

            DgColumn.ItemsSource = null;

            if (excel == null)
            {
                MessageBox.Show(this, "Excel ファイルが開かれています。閉じてから再度取込を行ってください。", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        /// <summary>
        /// スクロール位置を Top へ移動
        /// </summary>
        /// <param name="dg"></param>
        private void SetScrollTop(DataGrid dg)
        {
            var border = VisualTreeHelper.GetChild(dg, 0) as Decorator;
            var scrollViewer = border?.Child as ScrollViewer;
            scrollViewer?.ScrollToTop();
        }

        /// <summary>
        /// 必須と Undo の全選択管理用フィールド
        /// </summary>
        private bool _currentIsRequiredChecked;
        private bool _currentIsNeedUndoChecked = true;

        /// <summary>
        /// ヘッダークリック時の全選択/全解除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void columnHeader_Click(object sender, RoutedEventArgs e)
        {
            var header = sender as DataGridColumnHeader;
            if (header == null) return;

            if (header.Content.ToString() == "必須")
            {
                foreach (Columns entity in DgColumn.ItemsSource)
                {
                    entity.Required = !_currentIsRequiredChecked;
                }

                _currentIsRequiredChecked = !_currentIsRequiredChecked;
            }

            if (header.Content.ToString() == "Undo")
            {
                foreach (Columns entity in DgColumn.ItemsSource)
                {
                    entity.NeedUndo = !_currentIsNeedUndoChecked;
                }

                _currentIsNeedUndoChecked = !_currentIsNeedUndoChecked;
            }
        }
    }
}
