using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using FastWpfGrid;
using ExcelMerge.GUI.Utilities;
using ExcelMerge.GUI.ViewModels;
using ExcelMerge.GUI.Settings;
using ExcelMerge.GUI.Models;
using ExcelMerge.GUI.Services;

namespace ExcelMerge.GUI.Views
{
    public partial class DiffView : UserControl
    {
        private ExcelSheetDiffConfig diffConfig = new ExcelSheetDiffConfig();
        private IContainer container;
        private const string srcKey = "src";
        private const string dstKey = "dst";

        private FastGridControl copyTargetGrid;
        private MergeResult mergeResult;
        private ThreeWayDiffResult _threeWayResult;
        private const string baseKey = "base";
        private readonly SearchService searchService = new SearchService();

        private ExcelWorkbook _srcWorkbook;
        private ExcelWorkbook _dstWorkbook;

        public DiffView()
        {
            InitializeComponent();
            InitializeContainer();
            InitializeEventListeners();

            App.Instance.OnSettingUpdated += OnApplicationSettingUpdated;

            SearchTextCombobox.ItemsSource = App.Instance.Setting.SearchHistory.ToList();
        }

        private DiffViewModel GetViewModel()
        {
            return DataContext as DiffViewModel;
        }

        private void InitializeContainer()
        {
            container = new SimpleContainer();
            container
                .RegisterInstance(srcKey, SrcDataGrid)
                .RegisterInstance(dstKey, DstDataGrid)
                .RegisterInstance(baseKey, BaseDataGrid)
                .RegisterInstance(srcKey, SrcLocationGrid)
                .RegisterInstance(dstKey, DstLocationGrid)
                .RegisterInstance(srcKey, SrcViewRectangle)
                .RegisterInstance(dstKey, DstViewRectangle)
                .RegisterInstance(srcKey, SrcValueTextBox)
                .RegisterInstance(dstKey, DstValueTextBox)
                .RegisterInstance(baseKey, BaseValueTextBox);
        }

        private void InitializeEventListeners()
        {
            var srcEventHandler = new DiffViewEventHandler(srcKey);
            var dstEventHandler = new DiffViewEventHandler(dstKey);
            var baseEventHandler = new DiffViewEventHandler(baseKey);

            DataGridEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            DataGridEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            DataGridEventDispatcher.Instance.Listeners.Add(baseEventHandler);
            LocationGridEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            LocationGridEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            ViewportEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            ViewportEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            ValueTextBoxEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            ValueTextBoxEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            ValueTextBoxEventDispatcher.Instance.Listeners.Add(baseEventHandler);
        }

        private void OnApplicationSettingUpdated()
        {
            var e = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchApplicationSettingUpdateEvent(e);
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchParentLoadEvent(args);

            ExecuteDiff(isStartup: true);
        }

        private ExcelSheetDiffConfig CreateDiffConfig(FileSetting srcFileSetting, FileSetting dstFileSetting, bool isStartup)
        {
            var config = new ExcelSheetDiffConfig();

            config.SrcSheetIndex = SrcSheetCombobox.SelectedIndex;
            config.DstSheetIndex = DstSheetCombobox.SelectedIndex;
            config.CompareFormula = CompareFormulaCheckbox.IsChecked == true;
            config.IgnoreWhitespace = IgnoreWhitespaceCheckbox.IsChecked == true;
            if (double.TryParse(NumericPrecisionTextBox.Text, out var precisionVal))
                config.NumericPrecision = precisionVal;

            if (srcFileSetting != null)
            {
                if (isStartup)
                    config.SrcSheetIndex = GetSheetIndex(srcFileSetting, SrcSheetCombobox.Items);

                config.SrcHeaderIndex = srcFileSetting.ColumnHeaderIndex;
            }

            if (dstFileSetting != null)
            {
                if (isStartup)
                    config.DstSheetIndex = GetSheetIndex(dstFileSetting, DstSheetCombobox.Items);

                config.DstHeaderIndex = dstFileSetting.ColumnHeaderIndex;
            }

            return config;
        }

        private int GetSheetIndex(FileSetting fileSetting, ItemCollection sheetNames)
        {
            if (fileSetting == null)
                return -1;

            var index = fileSetting.SheetIndex;
            if (!string.IsNullOrEmpty(fileSetting.SheetName))
                index = sheetNames.IndexOf(fileSetting.SheetName);

            if (index < 0 || index >= sheetNames.Count)
            {
                MessageBox.Show(Properties.Resources.Msg_OutofSheetRange);
                index = 0;
            }

            return index;
        }

        private void LocationGrid_MouseDown(object sender, MouseEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchMouseDownEvent(args, e);
        }

        private void LocationGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
                LocationGridEventDispatcher.Instance.DispatchMouseDownEvent(args, e);
            }
        }

        private void LocationGrid_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchMouseWheelEvent(args, e);
        }

        private void DataGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(sender as FastGridControl, container);
            DataGridEventDispatcher.Instance.DispatchSizeChangeEvent(args, e);
        }

        private void LocationGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchSizeChangeEvent(args, e);
        }

        private void DataGrid_SelectedCellsChanged(object sender, FastWpfGrid.SelectionChangedEventArgs e)
        {
            var grid = copyTargetGrid = sender as FastGridControl;
            if (grid == null)
                return;

            copyTargetGrid = grid;

            var args = new DiffViewEventArgs<FastGridControl>(sender as FastGridControl, container);
            DataGridEventDispatcher.Instance.DispatchSelectedCellChangeEvent(args);

            if (!SrcDataGrid.CurrentCell.Row.HasValue || !DstDataGrid.CurrentCell.Row.HasValue)
                return;

            if (!SrcDataGrid.CurrentCell.Column.HasValue || !DstDataGrid.CurrentCell.Column.HasValue)
                return;

            if (SrcDataGrid.Model == null || DstDataGrid.Model == null)
                return;

            var srcModel = SrcDataGrid.Model as DiffGridModel;
            var dstModel = DstDataGrid.Model as DiffGridModel;

            var srcValue = srcModel.GetCellText(SrcDataGrid.CurrentCell.Row.Value, SrcDataGrid.CurrentCell.Column.Value, true);
            var dstValue = dstModel.GetCellText(DstDataGrid.CurrentCell.Row.Value, DstDataGrid.CurrentCell.Column.Value, true);

            ExcelCellDiff cellDiff;
            srcModel.TryGetCellDiffPublic(SrcDataGrid.CurrentCell.Row.Value, SrcDataGrid.CurrentCell.Column.Value, out cellDiff);

            UpdateValueDiff(srcValue, dstValue, cellDiff);

            if (App.Instance.Setting.AlwaysExpandCellDiff)
            {
                var a = new DiffViewEventArgs<RichTextBox>(null, container, TargetType.First);
                ValueTextBoxEventDispatcher.Instance.DispatchGotFocusEvent(a);
            }
        }

        private void ValueTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<RichTextBox>(sender as RichTextBox, container, TargetType.First);
            ValueTextBoxEventDispatcher.Instance.DispatchGotFocusEvent(args);
        }

        private void ValueTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<RichTextBox>(sender as RichTextBox, container, TargetType.First);
            ValueTextBoxEventDispatcher.Instance.DispatchLostFocusEvent(args);
        }

        private string GetRichTextString(RichTextBox textBox)
        {
            var textRange = new TextRange(textBox.Document.ContentStart, textBox.Document.ContentEnd);

            return textRange.Text;
        }

        private void UpdateValueDiff(string srcValue, string dstValue, ExcelCellDiff cellDiff = null)
        {
            SrcValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Clear();
            DstValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Clear();

            var srcParagraph = SrcValueTextBox.Document.Blocks.First().ContentStart.Paragraph;
            var dstParagraph = DstValueTextBox.Document.Blocks.First().ContentStart.Paragraph;

            var modifiedColor = App.Instance.Setting.ModifiedColor;
            var highlightBrush = new SolidColorBrush(
                Color.FromArgb(180, modifiedColor.R, modifiedColor.G, modifiedColor.B));

            if (!string.IsNullOrEmpty(srcValue) && !string.IsNullOrEmpty(dstValue) && srcValue != dstValue)
            {
                // Both sides have values and they differ: use character-level diff
                var srcSegments = TextDiffUtil.ComputeInlineDiffSrc(srcValue, dstValue);
                var dstSegments = TextDiffUtil.ComputeInlineDiff(srcValue, dstValue);

                foreach (var seg in srcSegments)
                {
                    var run = new Run(seg.Text);
                    if (seg.IsModified)
                        run.Background = highlightBrush;
                    srcParagraph.Inlines.Add(run);
                }

                foreach (var seg in dstSegments)
                {
                    var run = new Run(seg.Text);
                    if (seg.IsModified)
                        run.Background = highlightBrush;
                    dstParagraph.Inlines.Add(run);
                }
            }
            else
            {
                // Equal values, or one/both sides empty: show plain text
                srcParagraph.Inlines.Add(new Run(srcValue));
                dstParagraph.Inlines.Add(new Run(dstValue));
            }

            // Append comment text in italic gray if the cell has a comment
            if (cellDiff != null)
            {
                if (!string.IsNullOrEmpty(cellDiff.SrcCell.Comment))
                {
                    srcParagraph.Inlines.Add(new Run("\n[Comment] " + cellDiff.SrcCell.Comment)
                    {
                        Foreground = Brushes.Gray,
                        FontStyle = FontStyles.Italic
                    });
                }

                if (!string.IsNullOrEmpty(cellDiff.DstCell.Comment))
                {
                    dstParagraph.Inlines.Add(new Run("\n[Comment] " + cellDiff.DstCell.Comment)
                    {
                        Foreground = Brushes.Gray,
                        FontStyle = FontStyles.Italic
                    });
                }
            }
        }

        private void DiffButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteDiff();
        }

        private ExcelSheetReadConfig CreateReadConfig()
        {
            var setting = ((App)Application.Current).Setting;

            return new ExcelSheetReadConfig()
            {
                TrimFirstBlankRows = setting.SkipFirstBlankRows,
                TrimFirstBlankColumns = setting.SkipFirstBlankColumns,
                TrimLastBlankRows = setting.TrimLastBlankRows,
                TrimLastBlankColumns = setting.TrimLastBlankColumns,
            };
        }

        private Tuple<ExcelWorkbook, ExcelWorkbook> ReadWorkbooks()
        {
            ExcelWorkbook swb = null;
            ExcelWorkbook dwb = null;
            var srcPath = File.Exists(SrcPathTextBox.Text) ? SrcPathTextBox.Text : GetOrCreateEmptyFile();
            var dstPath = File.Exists(DstPathTextBox.Text) ? DstPathTextBox.Text : GetOrCreateEmptyFile();
            ProgressWindow.DoWorkWithModal(progress =>
            {
                progress.Report(Properties.Resources.Msg_ReadingFiles);

                var config = CreateReadConfig();
                swb = ExcelWorkbook.Create(srcPath, config);
                dwb = ExcelWorkbook.Create(dstPath, config);
            });

            return Tuple.Create(swb, dwb);
        }

        private Tuple<FileSetting, FileSetting> FindFileSettings(bool isStartup)
        {
            FileSetting srcSetting = null;
            FileSetting dstSetting = null;
            var srcPath = SrcPathTextBox.Text;
            var dstPath = DstPathTextBox.Text;

            var srcSelectedItem = SrcSheetCombobox.SelectedItem;
            var dstSelectedItem = DstSheetCombobox.SelectedItem;

            if (!IgnoreFileSettingCheckbox.IsChecked.Value
                && srcSelectedItem != null && dstSelectedItem != null)
            {
                srcSetting =
                    FindFilseSetting(Path.GetFileName(srcPath), SrcSheetCombobox.SelectedIndex, srcSelectedItem.ToString(), isStartup);

                dstSetting =
                    FindFilseSetting(Path.GetFileName(dstPath), DstSheetCombobox.SelectedIndex, dstSelectedItem.ToString(), isStartup);

                diffConfig = CreateDiffConfig(srcSetting, dstSetting, isStartup);
            }
            else
            {
                diffConfig = new ExcelSheetDiffConfig();

                diffConfig.SrcSheetIndex = Math.Max(SrcSheetCombobox.SelectedIndex, 0);
                diffConfig.DstSheetIndex = Math.Max(DstSheetCombobox.SelectedIndex, 0);
                diffConfig.CompareFormula = CompareFormulaCheckbox.IsChecked == true;
                diffConfig.IgnoreWhitespace = IgnoreWhitespaceCheckbox.IsChecked == true;
                if (double.TryParse(NumericPrecisionTextBox.Text, out var precisionFallback))
                    diffConfig.NumericPrecision = precisionFallback;
            }

            return Tuple.Create(srcSetting, dstSetting);
        }

        private ExcelSheetDiff ExecuteDiff(ExcelSheet srcSheet, ExcelSheet dstSheet)
        {
            ExcelSheetDiff diff = null;
            ProgressWindow.DoWorkWithModal(progress =>
            {
                progress.Report(Properties.Resources.Msg_ExtractingDiff);
                diff = ExcelSheet.Diff(srcSheet, dstSheet, diffConfig);
            });

            return diff;
        }

        public void ExecuteDiff(bool isStartup = false)
        {
            var srcExists = File.Exists(SrcPathTextBox.Text);
            var dstExists = File.Exists(DstPathTextBox.Text);

            if (!srcExists && !dstExists)
                return;

            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchPreExecuteDiffEvent(args);

            Tuple<ExcelWorkbook, ExcelWorkbook> workbooks;
            try
            {
                workbooks = ReadWorkbooks();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var srcWorkbook = workbooks.Item1;
            var dstWorkbook = workbooks.Item2;
            _srcWorkbook = srcWorkbook;
            _dstWorkbook = dstWorkbook;

            // Always sync ComboBox items from actual workbooks
            SrcSheetCombobox.ItemsSource = srcWorkbook.Sheets.Keys.ToList();
            DstSheetCombobox.ItemsSource = dstWorkbook.Sheets.Keys.ToList();

            var fileSettings = FindFileSettings(isStartup);
            var srcFileSetting = fileSettings.Item1;
            var dstFileSetting = fileSettings.Item2;

            var srcSheetIdx = Math.Min(Math.Max(diffConfig.SrcSheetIndex, 0), srcWorkbook.Sheets.Count - 1);
            var dstSheetIdx = Math.Min(Math.Max(diffConfig.DstSheetIndex, 0), dstWorkbook.Sheets.Count - 1);
            SrcSheetCombobox.SelectedIndex = srcSheetIdx;
            DstSheetCombobox.SelectedIndex = dstSheetIdx;

            var srcSheetName = srcWorkbook.Sheets.Keys.ElementAtOrDefault(srcSheetIdx);
            var dstSheetName = dstWorkbook.Sheets.Keys.ElementAtOrDefault(dstSheetIdx);
            if (srcSheetName == null || dstSheetName == null)
                return;

            var srcSheet = srcWorkbook.Sheets[srcSheetName];
            var dstSheet = dstWorkbook.Sheets[dstSheetName];

            if (srcSheet.Rows.Count > 10000 || dstSheet.Rows.Count > 10000)
                MessageBox.Show(Properties.Resources.Msg_WarnSize);

            var diff = ExecuteDiff(srcSheet, dstSheet);
            mergeResult = new MergeResult(diff);

            var compareFormula = CompareFormulaCheckbox.IsChecked == true;
            var srcModel = new DiffGridModel(diff, DiffType.Source) { MergeResult = mergeResult, CompareFormula = compareFormula };
            var dstModel = new DiffGridModel(diff, DiffType.Dest) { MergeResult = mergeResult, CompareFormula = compareFormula };
            SrcDataGrid.Model = srcModel;
            DstDataGrid.Model = dstModel;

            // 3-way merge mode: load BASE workbook and compute three-way diff
            if (IsMergeMode && !string.IsNullOrEmpty(_mergeBasePath) && File.Exists(_mergeBasePath))
            {
                try
                {
                    var baseWb = ExcelWorkbook.Create(_mergeBasePath, CreateReadConfig());
                    var baseSheetName = baseWb.Sheets.Keys.ElementAtOrDefault(
                        Math.Min(Math.Max(diffConfig.SrcSheetIndex, 0), baseWb.Sheets.Count - 1));

                    if (baseSheetName != null)
                    {
                        var baseSheet = baseWb.Sheets[baseSheetName];

                        // Compute 3-way diff: srcSheet=THEIRS(left), dstSheet=MINE(right)
                        _threeWayResult = ThreeWayDiff.Compute(baseSheet, dstSheet, srcSheet, diffConfig);

                        // Pass 3-way merge result to both grid models for cell coloring
                        srcModel.ThreeWayMergeResult = _threeWayResult;
                        dstModel.ThreeWayMergeResult = _threeWayResult;

                        // Create BASE grid model by diffing base against itself (no differences shown)
                        var baseDiff = ExcelSheet.Diff(baseSheet, baseSheet, diffConfig);
                        var baseModel = new DiffGridModel(baseDiff, DiffType.Source) { CompareFormula = compareFormula };
                        BaseDataGrid.Model = baseModel;

                        // Enable the Show BASE toggle
                        ShowBaseToggle.IsEnabled = true;

                        // Auto-show BASE panel if conflicts exist
                        if (_threeWayResult.HasConflicts)
                        {
                            ShowBaseToggle.IsChecked = true;
                            ShowBasePanel();
                        }

                        UpdateMergeStatus();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to load BASE file: {ex.Message}");
                }
            }
            else
            {
                _threeWayResult = null;
                ShowBaseToggle.IsEnabled = false;
                ShowBaseToggle.IsChecked = false;
                HideBasePanel();
                ConflictCountLabel.Visibility = Visibility.Collapsed;
            }

            args = new DiffViewEventArgs<FastGridControl>(SrcDataGrid, container);
            DataGridEventDispatcher.Instance.DispatchFileSettingUpdateEvent(args, srcFileSetting);

            args = new DiffViewEventArgs<FastGridControl>(DstDataGrid, container);
            DataGridEventDispatcher.Instance.DispatchFileSettingUpdateEvent(args, dstFileSetting);

            args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchDisplayFormatChangeEvent(args, ShowOnlyDiffRadioButton.IsChecked.Value);
            DataGridEventDispatcher.Instance.DispatchPostExecuteDiffEvent(args);

            var summary = diff.CreateSummary();
            GetViewModel().UpdateDiffSummary(summary);

            if (!App.Instance.KeepFileHistory)
                App.Instance.UpdateRecentFiles(SrcPathTextBox.Text, DstPathTextBox.Text);

            if (App.Instance.Setting.NotifyEqual && !summary.HasDiff)
                MessageBox.Show(Properties.Resources.Message_NoDiff);

            if (App.Instance.Setting.FocusFirstDiff)
                MoveNextModifiedCell();
        }

        private FileSetting FindFilseSetting(string fileName, int sheetIndex, string sheetName, bool isStartup)
        {
            var results = new List<FileSetting>();
            foreach (var setting in App.Instance.Setting.FileSettings)
            {
                if (setting.UseRegex)
                {
                    var regex = new System.Text.RegularExpressions.Regex(setting.Name);

                    if (regex.IsMatch(fileName))
                        results.Add(setting);
                }
                else
                {
                    if (setting.ExactMatch)
                    {
                        if (setting.Name == fileName)
                            results.Add(setting);
                    }
                    else
                    {
                        if (fileName.Contains(setting.Name))
                            results.Add(setting);
                    }
                }
            }

            if (isStartup)
                return results.FirstOrDefault(r => r.IsStartupSheet) ?? results.FirstOrDefault() ?? null;

            return results.FirstOrDefault(r => r.SheetName == sheetName) ?? results.FirstOrDefault(r => r.SheetIndex == sheetIndex) ?? null;
        }

        private void SetRowHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchRowHeaderChagneEvent(args);
                }
            }
        }

        private void ResetRowHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchRowHeaderResetEvent(args);
                }
            }
        }

        private void SetColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchColumnHeaderChangeEvent(args);
                }
            }
        }

        private void ResetColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchColumnHeaderResetEvent(args);
                }
            }
        }

        private void SwapButton_Click(object sender, RoutedEventArgs e)
        {
            Swap();
        }

        private void SrcBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            BrowseFile(SrcPathTextBox);
        }

        private void DstBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            BrowseFile(DstPathTextBox);
        }

        private void BrowseFile(TextBox target)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls;*.csv;*.tsv|All Files|*.*"
            };
            if (!string.IsNullOrEmpty(target.Text) && Directory.Exists(Path.GetDirectoryName(target.Text)))
                dlg.InitialDirectory = Path.GetDirectoryName(target.Text);

            if (dlg.ShowDialog() == true)
            {
                target.Text = dlg.FileName;
                target.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
                TryAutoExecuteDiff();
            }
        }

        private void PathTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var tb = sender as TextBox;
                tb?.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
                TryAutoExecuteDiff();
            }
        }

        private string _lastSrcPath = string.Empty;
        private string _lastDstPath = string.Empty;

        private void PathTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var src = SrcPathTextBox.Text;
            var dst = DstPathTextBox.Text;

            if (src == _lastSrcPath && dst == _lastDstPath)
                return;

            if (!File.Exists(src) && !File.Exists(dst))
                return;

            _lastSrcPath = src;
            _lastDstPath = dst;

            SrcPathTextBox.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
            DstPathTextBox.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();

            Dispatcher.InvokeAsync(() => ExecuteDiff(),
                System.Windows.Threading.DispatcherPriority.Background);
        }

        private void TryAutoExecuteDiff()
        {
            if (File.Exists(SrcPathTextBox.Text) || File.Exists(DstPathTextBox.Text))
                ExecuteDiff();
        }

        private string _emptyFilePath;

        private string GetOrCreateEmptyFile()
        {
            if (_emptyFilePath != null && File.Exists(_emptyFilePath))
                return _emptyFilePath;

            _emptyFilePath = Path.Combine(Path.GetTempPath(), "ExcelMerge_empty.xlsx");
            ExcelUtility.CreateWorkbook(_emptyFilePath, ExcelWorkbookType.XLSX);
            return _emptyFilePath;
        }

        private void Swap()
        {
            var srcTmp = SrcSheetCombobox.SelectedIndex;
            var dstTmp = DstSheetCombobox.SelectedIndex;

            var tmp = SrcPathTextBox.Text;
            SrcPathTextBox.Text = DstPathTextBox.Text;
            DstPathTextBox.Text = tmp;

            diffConfig.SrcSheetIndex = dstTmp;
            diffConfig.DstSheetIndex = srcTmp;

            ExecuteDiff();
        }

        private void DiffByHeaderSrc_Click(object sender, RoutedEventArgs e)
        {
            var headerIndex = SrcDataGrid.CurrentCell.Row.HasValue ? SrcDataGrid.CurrentCell.Row.Value : -1;

            diffConfig.SrcHeaderIndex= headerIndex;

            ExecuteDiff();
        }

        private void DiffByHeaderDst_Click(object sender, RoutedEventArgs e)
        {
            var headerIndex = DstDataGrid.CurrentCell.Row.HasValue ? DstDataGrid.CurrentCell.Row.Value : -1;

            diffConfig.DstSheetIndex = headerIndex;

            ExecuteDiff();
        }

        private void ShowAllRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchDisplayFormatChangeEvent(args, false);
        }

        private void ShowOnlyDiffRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchDisplayFormatChangeEvent(args, true);
        }

        private void CompareFormulaCheckbox_Changed(object sender, RoutedEventArgs e)
        {
            ExecuteDiff();
        }

        private void IgnoreRulesChanged(object sender, RoutedEventArgs e)
        {
            ExecuteDiff();
        }

        private bool ValidateDataGrids()
        {
            return SrcDataGrid.Model != null && DstDataGrid.Model != null;
        }

        private void ValuteTextBox_ScrollChanged(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<RichTextBox>(sender as RichTextBox, container);
            ValueTextBoxEventDispatcher.Instance.DispatchScrolledEvent(args, (ScrollChangedEventArgs)e);
        }

        private bool SwitchToNextSheetWithDiff(bool forward)
        {
            if (_srcWorkbook == null || _dstWorkbook == null) return false;

            var currentSheetIndex = SrcSheetCombobox.SelectedIndex;
            var sheetCount = Math.Min(_srcWorkbook.Sheets.Count, _dstWorkbook.Sheets.Count);

            for (int i = 1; i < sheetCount; i++)
            {
                var nextIndex = forward
                    ? (currentSheetIndex + i) % sheetCount
                    : (currentSheetIndex - i + sheetCount) % sheetCount;

                var srcSheetName = _srcWorkbook.Sheets.Keys.ElementAtOrDefault(nextIndex);
                var dstSheetName = _dstWorkbook.Sheets.Keys.ElementAtOrDefault(nextIndex);
                if (srcSheetName == null || dstSheetName == null) continue;

                var srcSheet = _srcWorkbook.Sheets[srcSheetName];
                var dstSheet = _dstWorkbook.Sheets[dstSheetName];

                var diff = ExcelSheet.Diff(srcSheet, dstSheet, diffConfig);
                var summary = diff.CreateSummary();
                if (summary.HasDiff)
                {
                    // Switch to this sheet — update ComboBox and re-execute diff
                    SrcSheetCombobox.SelectedIndex = nextIndex;
                    DstSheetCombobox.SelectedIndex = nextIndex;
                    ExecuteDiff();
                    return true;
                }
            }

            return false;
        }

        private void NextModifiedCellButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextModifiedCell();
        }

        private void MoveNextModifiedCell()
        {
            if (!ValidateDataGrids()) return;
            var navigated = DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetNextModifiedCell(c));
            if (!navigated && SwitchToNextSheetWithDiff(true))
                DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetNextModifiedCell(FastGridCellAddress.Empty));
        }

        private void PrevModifiedCellButton_Click(object sender, RoutedEventArgs e) => MovePrevModifiedCell();

        private void MovePrevModifiedCell()
        {
            if (!ValidateDataGrids()) return;
            var navigated = DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetPreviousModifiedCell(c));
            if (!navigated && SwitchToNextSheetWithDiff(false))
            {
                var model = SrcDataGrid.Model as DiffGridModel;
                var lastCell = new FastGridCellAddress(model.RowCount - 1, model.ColumnCount - 1);
                DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetPreviousModifiedCell(lastCell));
            }
        }

        private void NextModifiedRowButton_Click(object sender, RoutedEventArgs e) => MoveNextModifiedRow();

        private void MoveNextModifiedRow()
        {
            if (!ValidateDataGrids()) return;
            var navigated = DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetNextModifiedRow(c));
            if (!navigated && SwitchToNextSheetWithDiff(true))
                DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetNextModifiedRow(FastGridCellAddress.Empty));
        }

        private void PrevModifiedRowButton_Click(object sender, RoutedEventArgs e) => MovePrevModifiedRow();

        private void MovePrevModifiedRow()
        {
            if (!ValidateDataGrids()) return;
            var navigated = DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetPreviousModifiedRow(c));
            if (!navigated && SwitchToNextSheetWithDiff(false))
            {
                var model = SrcDataGrid.Model as DiffGridModel;
                var lastCell = new FastGridCellAddress(model.RowCount - 1, model.ColumnCount - 1);
                DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetPreviousModifiedRow(lastCell));
            }
        }

        private void NextAddedRowButton_Click(object sender, RoutedEventArgs e) => MoveNextAddedRow();

        private void MoveNextAddedRow()
        {
            if (!ValidateDataGrids()) return;
            var navigated = DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetNextAddedRow(c));
            if (!navigated && SwitchToNextSheetWithDiff(true))
                DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetNextAddedRow(FastGridCellAddress.Empty));
        }

        private void PrevAddedRowButton_Click(object sender, RoutedEventArgs e) => MovePrevAddedRow();

        private void MovePrevAddedRow()
        {
            if (!ValidateDataGrids()) return;
            var navigated = DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetPreviousAddedRow(c));
            if (!navigated && SwitchToNextSheetWithDiff(false))
            {
                var model = SrcDataGrid.Model as DiffGridModel;
                var lastCell = new FastGridCellAddress(model.RowCount - 1, model.ColumnCount - 1);
                DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetPreviousAddedRow(lastCell));
            }
        }

        private void NextRemovedRowButton_Click(object sender, RoutedEventArgs e) => MoveNextRemovedRow();

        private void MoveNextRemovedRow()
        {
            if (!ValidateDataGrids()) return;
            var navigated = DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetNextRemovedRow(c));
            if (!navigated && SwitchToNextSheetWithDiff(true))
                DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetNextRemovedRow(FastGridCellAddress.Empty));
        }

        private void PrevRemovedRowButton_Click(object sender, RoutedEventArgs e) => MovePrevRemovedRow();

        private void MovePrevRemovedRow()
        {
            if (!ValidateDataGrids()) return;
            var navigated = DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetPreviousRemovedRow(c));
            if (!navigated && SwitchToNextSheetWithDiff(false))
            {
                var model = SrcDataGrid.Model as DiffGridModel;
                var lastCell = new FastGridCellAddress(model.RowCount - 1, model.ColumnCount - 1);
                DiffNavigator.Navigate(SrcDataGrid, (m, c) => m.GetPreviousRemovedRow(lastCell));
            }
        }

        private void PrevMatchCellButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevMatchCell();
        }

        private void MovePrevMatchCell()
        {
            var text = GetSearchTextAndUpdateHistory();
            if (text == null) return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousMatchCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell, text,
                ExactMatchCheckBox.IsChecked.Value, CaseSensitiveCheckBox.IsChecked.Value, RegexCheckBox.IsChecked.Value, ShowOnlyDiffRadioButton.IsChecked.Value);
            if (!nextCell.IsEmpty)
                SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextMatchCellButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextMatchCell();
        }

        private void MoveNextMatchCell()
        {
            var text = GetSearchTextAndUpdateHistory();
            if (text == null) return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextMatchCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell, text,
                ExactMatchCheckBox.IsChecked.Value, CaseSensitiveCheckBox.IsChecked.Value, RegexCheckBox.IsChecked.Value, ShowOnlyDiffRadioButton.IsChecked.Value);
            if (!nextCell.IsEmpty)
                SrcDataGrid.CurrentCell = nextCell;
        }

        private string GetSearchTextAndUpdateHistory()
        {
            if (!ValidateDataGrids())
                return null;

            var text = searchService.UpdateSearchHistory(SearchTextCombobox.Text);
            if (text != null)
                SearchTextCombobox.ItemsSource = searchService.GetSearchHistory();

            return text;
        }

        #region Search Overlay

        private void ShowSearchOverlay()
        {
            SearchOverlayPanel.Visibility = Visibility.Visible;

            // Sync current search text from expander ComboBox if overlay is empty
            if (string.IsNullOrEmpty(SearchOverlayComboBox.Text) && !string.IsNullOrEmpty(SearchTextCombobox.Text))
                SearchOverlayComboBox.Text = SearchTextCombobox.Text;

            SearchOverlayComboBox.ItemsSource = App.Instance.Setting.SearchHistory.ToList();

            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Input,
                new Action(() =>
                {
                    SearchOverlayComboBox.Focus();
                    Keyboard.Focus(SearchOverlayComboBox);
                    var tb = SearchOverlayComboBox.Template.FindName("PART_EditableTextBox", SearchOverlayComboBox) as TextBox;
                    tb?.SelectAll();
                }));
        }

        private void HideSearchOverlay()
        {
            SearchOverlayPanel.Visibility = Visibility.Collapsed;
            SrcDataGrid.Focus();
        }

        private void SyncSearchTextFromOverlay()
        {
            SearchTextCombobox.Text = SearchOverlayComboBox.Text;
        }

        private void SearchOverlayComboBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                SyncSearchTextFromOverlay();
                if ((Keyboard.Modifiers & ModifierKeys.Shift) != 0)
                    MovePrevMatchCell();
                else
                    MoveNextMatchCell();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                HideSearchOverlay();
                e.Handled = true;
            }
        }

        private void SearchOverlayPrev_Click(object sender, RoutedEventArgs e)
        {
            SyncSearchTextFromOverlay();
            MovePrevMatchCell();
        }

        private void SearchOverlayNext_Click(object sender, RoutedEventArgs e)
        {
            SyncSearchTextFromOverlay();
            MoveNextMatchCell();
        }

        private void SearchOverlayClose_Click(object sender, RoutedEventArgs e)
        {
            HideSearchOverlay();
        }

        private void FindButton_Click(object sender, RoutedEventArgs e)
        {
            ShowSearchOverlay();
        }

        private void SearchTextCombobox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if ((Keyboard.Modifiers & ModifierKeys.Shift) != 0)
                    MovePrevMatchCell();
                else
                    MoveNextMatchCell();
                e.Handled = true;
            }
        }

        #endregion

        #region History

        private void HistoryButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var menu = new ContextMenu();

            var recentSets = App.Instance.GetRecentFileSets().ToList();
            if (recentSets.Count == 0)
            {
                menu.Items.Add(new MenuItem { Header = "(empty)", IsEnabled = false });
            }
            else
            {
                foreach (var set in recentSets)
                {
                    var src = set.Item1;
                    var dst = set.Item2;
                    var srcName = Path.GetFileName(src);
                    var dstName = Path.GetFileName(dst);
                    var item = new MenuItem
                    {
                        Header = $"{srcName}  \u2194  {dstName}",
                        ToolTip = $"{src}\n\u2194\n{dst}",
                    };
                    item.Click += (s, args) => ApplyFileSet(src, dst);
                    menu.Items.Add(item);
                }
            }

            menu.PlacementTarget = button;
            menu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
            menu.IsOpen = true;
        }

        private void ApplyFileSet(string srcPath, string dstPath)
        {
            SrcPathTextBox.Text = srcPath;
            DstPathTextBox.Text = dstPath;
            SrcPathTextBox.GetBindingExpression(System.Windows.Controls.TextBox.TextProperty)?.UpdateSource();
            DstPathTextBox.GetBindingExpression(System.Windows.Controls.TextBox.TextProperty)?.UpdateSource();
            ExecuteDiff(isStartup: false);
        }

        #endregion

        private void CopyToClipboardSelectedCells(string separator)
        {
            ClipboardService.CopySelectedCells(copyTargetGrid, separator);
        }

        private void UserControl_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Right:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextModifiedCell();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Left:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevModifiedCell();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Down:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextModifiedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Up:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevModifiedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.L:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextRemovedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.O:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevRemovedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.K:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextAddedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.I:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevAddedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.F8:
                    {
                        MovePrevMatchCell();
                        e.Handled = true;
                    }
                    break;
                case Key.F9:
                    {
                        MoveNextMatchCell();
                        e.Handled = true;
                    }
                    break;
                case Key.Escape:
                    {
                        if (SearchOverlayPanel.Visibility == Visibility.Visible)
                        {
                            HideSearchOverlay();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.F3:
                    {
                        if (SearchOverlayPanel.Visibility == Visibility.Visible)
                            SyncSearchTextFromOverlay();
                        if ((Keyboard.Modifiers & ModifierKeys.Shift) != 0)
                            MovePrevMatchCell();
                        else
                            MoveNextMatchCell();
                        e.Handled = true;
                    }
                    break;
                case Key.F:
                    {
                        if ((Keyboard.Modifiers & ModifierKeys.Control) != 0)
                        {
                            ShowSearchOverlay();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.C:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            CopyToClipboardSelectedCells(Keyboard.IsKeyDown(Key.RightShift) || Keyboard.IsKeyDown(Key.LeftShift) ? "," : "\t");
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.B:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            ShowLog();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.F5:
                    {
                        ExecuteDiff();
                        e.Handled = true;
                    }
                    break;
            }
        }

        private void ShowLog()
        {
            var log = BuildCellBaseLog();

            (App.Current.MainWindow as MainWindow).WriteToConsole(log);
        }

        private void BuildCellBaseLog_Click(object sender, RoutedEventArgs e)
        {
            ShowLog();
        }

        private string BuildCellBaseLog()
        {
            var srcModel = SrcDataGrid.Model as DiffGridModel;
            var dstModel = DstDataGrid.Model as DiffGridModel;

            return LogBuilder.BuildCellBaseLog(
                srcModel,
                dstModel,
                SrcDataGrid.SelectedCells,
                App.Instance.Setting.LogFormat,
                App.Instance.Setting.AddedRowLogFormat,
                App.Instance.Setting.RemovedRowLogFormat,
                Properties.Resources.Word_Blank);
        }

        private void CopyAsTsv_Click(object sender, RoutedEventArgs e)
        {
            CopyToClipboardSelectedCells("\t");
        }

        private void CopyAsCsv_Click(object sender, RoutedEventArgs e)
        {
            CopyToClipboardSelectedCells(",");
        }

        private void AcceptSrc_Click(object sender, RoutedEventArgs e)
        {
            ApplyMergeToSelectedCells(MergeSide.Src);
        }

        private void AcceptDst_Click(object sender, RoutedEventArgs e)
        {
            ApplyMergeToSelectedCells(MergeSide.Dst);
        }

        private void AcceptSrcRow_Click(object sender, RoutedEventArgs e)
        {
            ApplyMergeToSelectedRows(MergeSide.Src);
        }

        private void AcceptDstRow_Click(object sender, RoutedEventArgs e)
        {
            ApplyMergeToSelectedRows(MergeSide.Dst);
        }

        private void ApplyMergeToSelectedCells(MergeSide side)
        {
            if (mergeResult == null || copyTargetGrid == null) return;
            var model = copyTargetGrid.Model as DiffGridModel;
            MergeApplicator.ApplyToSelectedCells(mergeResult, model, copyTargetGrid.SelectedCells, side);
            SrcDataGrid.InvalidateAll();
            DstDataGrid.InvalidateAll();
        }

        private void ApplyMergeToSelectedRows(MergeSide side)
        {
            if (mergeResult == null || copyTargetGrid == null) return;
            var model = copyTargetGrid.Model as DiffGridModel;
            MergeApplicator.ApplyToSelectedRows(mergeResult, model, copyTargetGrid.SelectedCells, side);
            SrcDataGrid.InvalidateAll();
            DstDataGrid.InvalidateAll();
        }

        private void SaveMergeButton_Click(object sender, RoutedEventArgs e)
        {
            // 3-way merge mode: use MergeWriter with ThreeWayDiffResult
            if (IsMergeMode && _threeWayResult != null)
            {
                if (_threeWayResult.UnresolvedConflictCount > 0)
                {
                    var result = MessageBox.Show(
                        $"{_threeWayResult.UnresolvedConflictCount} unresolved conflicts.\nSave anyway (MINE values will be used for unresolved)?",
                        "Unresolved Conflicts", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    if (result != MessageBoxResult.Yes)
                        return;

                    // Fill unresolved conflicts with MINE values
                    foreach (var row in _threeWayResult.Rows.Values)
                        foreach (var cell in row.Values)
                            if (cell.Status == CellMergeStatus.Conflict && cell.ResolvedValue == null)
                                cell.ResolvedValue = cell.MineValue;
                }

                var outputPath = _mergeOutputPath;
                if (string.IsNullOrEmpty(outputPath))
                {
                    var dlg = new Microsoft.Win32.SaveFileDialog
                    {
                        Filter = "Excel Files|*.xlsx",
                        DefaultExt = ".xlsx",
                        FileName = "merged.xlsx"
                    };
                    if (dlg.ShowDialog() != true)
                        return;
                    outputPath = dlg.FileName;
                }

                try
                {
                    // Find current sheet name for MergeWriter
                    var sheetName = SrcSheetCombobox.SelectedItem?.ToString() ?? "Sheet1";
                    MergeWriter.Write(_mergeBasePath, outputPath, _threeWayResult, sheetName);
                    MessageBox.Show(string.Format(Properties.Resources.Msg_MergeSaved, outputPath),
                        "ExcelMerge", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to save merge: {ex.Message}",
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                return;
            }

            // 2-way merge mode: use existing MergeResult
            if (mergeResult == null || mergeResult.DecisionCount == 0)
            {
                MessageBox.Show(Properties.Resources.Msg_NoMergeDecisions, "ExcelMerge",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var saveDlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                DefaultExt = ".xlsx",
                FileName = "merged.xlsx"
            };

            if (saveDlg.ShowDialog() == true)
            {
                mergeResult.WriteToFile(saveDlg.FileName);
                MessageBox.Show(string.Format(Properties.Resources.Msg_MergeSaved, saveDlg.FileName),
                    "ExcelMerge", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // Merge mode properties
        private string _mergeBasePath;
        private string _mergeOutputPath;
        public bool IsMergeMode { get; private set; }
        public bool IsReadonlyLeft { get; private set; }

        public void SetMergeMode(string basePath, string outputPath)
        {
            IsMergeMode = true;
            _mergeBasePath = basePath;
            _mergeOutputPath = outputPath;
        }

        public void SetReadonlyLeft()
        {
            IsReadonlyLeft = true;
            ApplyReadonlyLeft();
        }

        private void ApplyReadonlyLeft()
        {
            // When left side is read-only (SVN base file), hide Save Merge in diff mode
            // In merge mode, Save writes to --output path, so it stays visible
            if (!IsMergeMode)
                SaveMergeButton.Visibility = Visibility.Collapsed;
        }

        #region BASE Panel

        public void ShowBasePanel()
        {
            BaseColumnDef.Width = new GridLength(30, GridUnitType.Star);
            BaseSplitterColumnDef.Width = GridLength.Auto;
            BaseDataGrid.Visibility = Visibility.Visible;
            BaseValueTextBox.Visibility = Visibility.Visible;
            BaseSplitter.Visibility = Visibility.Visible;
        }

        public void HideBasePanel()
        {
            BaseColumnDef.Width = new GridLength(0);
            BaseSplitterColumnDef.Width = new GridLength(0);
            BaseDataGrid.Visibility = Visibility.Collapsed;
            BaseValueTextBox.Visibility = Visibility.Collapsed;
            BaseSplitter.Visibility = Visibility.Collapsed;
        }

        private void ShowBaseToggle_Checked(object sender, RoutedEventArgs e)
        {
            ShowBasePanel();
        }

        private void ShowBaseToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            HideBasePanel();
        }

        private void UpdateMergeStatus()
        {
            if (_threeWayResult == null)
            {
                ConflictCountLabel.Visibility = Visibility.Collapsed;
                return;
            }

            var conflicts = _threeWayResult.UnresolvedConflictCount;
            if (conflicts > 0)
            {
                ConflictCountLabel.Text = $"Conflicts: {conflicts}";
                ConflictCountLabel.Visibility = Visibility.Visible;

                var window = Window.GetWindow(this);
                if (window != null)
                    window.Title = $"ExcelMerge - Merge ({conflicts} conflicts)";
            }
            else
            {
                var autoMerged = _threeWayResult.AutoMergedCount;
                if (autoMerged > 0)
                {
                    ConflictCountLabel.Text = $"Auto-merged: {autoMerged}";
                    ConflictCountLabel.Foreground = new SolidColorBrush(Color.FromRgb(0x16, 0xA3, 0x4A));
                    ConflictCountLabel.Visibility = Visibility.Visible;
                }
                else
                {
                    ConflictCountLabel.Visibility = Visibility.Collapsed;
                }

                var window = Window.GetWindow(this);
                if (window != null)
                    window.Title = "ExcelMerge - Merge (no conflicts)";
            }
        }

        #endregion
    }
}
