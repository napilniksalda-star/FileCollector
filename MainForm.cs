// FileCollector V3.0 "Phoenix" - controller half of MainForm.
// UI construction lives in MainForm.Layout.cs. This file holds state,
// event handlers, and wires the Core/ services together.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using FileCollector.Core;
using FileCollector.Core.Copy;
using FileCollector.Core.Excel;
using FileCollector.Core.Logging;
using FileCollector.Core.Reporting;
using FileCollector.Core.Search;
using FileCollector.Plugins;
using FileCollector.Plugins.BuiltIn;

namespace FileCollector
{
    public partial class MainForm : Form
    {
        // ---------- State (underscore-prefixed names referenced by MainForm.Layout.cs) ----------
        private readonly List<string> _searchFolders = new();
        private string _excelPath = string.Empty;
        private string _destinationPath = string.Empty;
        private bool _isRunning;

        private List<PreviewItem> _previewItems = new();
        private readonly List<CopyOperation> _completedOps = new();

        private CancellationTokenSource? _cts;
        private BatchUiLogger? _logger;
        private PluginManager? _plugins;

        // Counters surfaced to the status label.
        private int _total, _copied, _notFound, _skipped;

        // Hidden defaults.
        private const bool SkipLockedFiles = true;
        private const bool SkipSystemHidden = true;

        public MainForm()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = VersionInfo.TitleBar;
            this.Size = new Size(900, 700);
            this.MinimumSize = new Size(850, 600);
            this.KeyPreview = true;
            this.BackColor = SystemColors.Control;
            this.ForeColor = SystemColors.ControlText;

            InitializeComponents();
            SetupToolTips();
            SetupContextMenus();
            SetupDragDrop();
            SetupKeyboardShortcuts();

            _logger = new BatchUiLogger(AppendLog);
            _plugins = new PluginManager(msg => _logger?.Info(msg));
            RegisterBuiltInPlugins();
            // External plugins from a sibling folder, optional.
            _plugins.LoadPlugins(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Plugins"));

            this.FormClosing += (_, __) =>
            {
                try { _cts?.Cancel(); } catch { }
                _plugins?.Shutdown();
                _logger?.Dispose();
            };

            DisplayInitialInstructions();
        }

        private void RegisterBuiltInPlugins()
        {
            if (_plugins == null) return;
            _plugins.Register(new HtmlExportPlugin());
            _plugins.Register(new DuplicateDetectorPlugin());
            _plugins.Register(new PdfAnalyzerPlugin());
        }

        private void DisplayInitialInstructions()
        {
            const string instructions =
                "========== ИНСТРУКЦИЯ ==========\n" +
                "1. Выберите Excel со списком имён в столбце B (.xlsx / .xls).\n" +
                "2. Добавьте одну или несколько папок для поиска.\n" +
                "3. Укажите папку для сбора результатов.\n" +
                "4. Отметьте нужные расширения (по умолчанию .pdf и .dxf).\n" +
                "5. Нажмите 'Предпросмотр' (Ctrl+P), затем 'Запустить копирование' (Ctrl+S).\n\n" +
                "Поля поддерживают drag-and-drop. Esc - отмена операции.\n" +
                "================================\n";
            AppendLog(instructions);
        }

        // ---------- Excel ----------
        private void BtnSelectExcel_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*",
                Title = "Выберите Excel файл",
                CheckFileExists = true
            };
            if (ofd.ShowDialog() != DialogResult.OK) return;
            _excelPath = ofd.FileName;
            txtExcelPath.Text = _excelPath;
            UpdateStartButtonState();
        }

        // ---------- Search folders ----------
        private void BtnAddSearchFolder_Click(object? sender, EventArgs e)
        {
            using var fbd = new FolderBrowserDialog
            {
                Description = "Выберите папку для поиска файлов",
                ShowNewFolderButton = false
            };
            if (fbd.ShowDialog() != DialogResult.OK) return;

            if (_searchFolders.Contains(fbd.SelectedPath))
            {
                MessageBox.Show("Эта папка уже добавлена в список.", "Информация",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _searchFolders.Add(fbd.SelectedPath);
            lstSearchFolders.Items.Add(fbd.SelectedPath);
            UpdateSearchFoldersHorizontalExtent();
            UpdateStartButtonState();
            UpdateSearchFolderButtons();
        }

        private void BtnRemoveSearchFolder_Click(object? sender, EventArgs e)
        {
            int idx = lstSearchFolders.SelectedIndex;
            if (idx < 0) return;
            _searchFolders.RemoveAt(idx);
            lstSearchFolders.Items.RemoveAt(idx);
            UpdateSearchFoldersHorizontalExtent();
            UpdateStartButtonState();
            UpdateSearchFolderButtons();
        }

        private void BtnClearSearchFolders_Click(object? sender, EventArgs e)
        {
            if (_searchFolders.Count == 0) return;
            if (MessageBox.Show("Удалить все папки из списка поиска?", "Подтверждение",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;
            _searchFolders.Clear();
            lstSearchFolders.Items.Clear();
            UpdateSearchFoldersHorizontalExtent();
            UpdateStartButtonState();
            UpdateSearchFolderButtons();
        }

        private void LstSearchFolders_SelectedIndexChanged(object? sender, EventArgs e)
        {
            UpdateSearchFolderButtons();
        }

        private void UpdateSearchFolderButtons()
        {
            btnRemoveSearchFolder.Enabled = !_isRunning && lstSearchFolders.SelectedIndex >= 0;
            btnClearSearchFolders.Enabled = !_isRunning && _searchFolders.Count > 0;
        }

        private void UpdateSearchFoldersHorizontalExtent()
        {
            int maxWidth = lstSearchFolders.ClientSize.Width;
            foreach (var item in lstSearchFolders.Items)
            {
                string text = item?.ToString() ?? string.Empty;
                int w = TextRenderer.MeasureText(text, lstSearchFolders.Font).Width;
                if (w > maxWidth) maxWidth = w;
            }
            lstSearchFolders.HorizontalExtent = maxWidth + 20;
        }

        // ---------- Destination ----------
        private void BtnSelectDestination_Click(object? sender, EventArgs e)
        {
            using var fbd = new FolderBrowserDialog
            {
                Description = "Выберите папку для сбора файлов",
                ShowNewFolderButton = true
            };
            if (fbd.ShowDialog() != DialogResult.OK) return;
            _destinationPath = fbd.SelectedPath;
            txtDestinationPath.Text = _destinationPath;
            UpdateStartButtonState();
        }

        // ---------- Extensions tooltip helpers ----------
        private void ClbExtensions_MouseMove(object? sender, MouseEventArgs e)
        {
            int index = clbExtensions.IndexFromPoint(e.Location);
            if (index < 0 || index == lastExtensionTooltipIndex) return;

            lastExtensionTooltipIndex = index;
            string ext = clbExtensions.Items[index]?.ToString() ?? string.Empty;

            if (SolidWorksExtensionTooltips.TryGetValue(ext, out var tip))
            {
                toolTip.SetToolTip(clbExtensions, tip);
                toolTip.Show(tip, clbExtensions, e.Location.X + 14, e.Location.Y + 18, 2500);
            }
            else
            {
                toolTip.SetToolTip(clbExtensions, ExtensionsToolTipText);
                toolTip.Hide(clbExtensions);
            }
        }

        private void ClbExtensions_MouseLeave(object? sender, EventArgs e)
        {
            lastExtensionTooltipIndex = -1;
            toolTip.Hide(clbExtensions);
            toolTip.SetToolTip(clbExtensions, ExtensionsToolTipText);
        }

        // ---------- Preview ----------
        private async void BtnPreview_Click(object? sender, EventArgs e)
        {
            if (_isRunning) return;
            if (_logger == null) return;

            ResetCounters();
            UpdateStats();

            try
            {
                SetOperationState(true);
                dgvPreview.Rows.Clear();
                _previewItems.Clear();
                progressBar.Value = 0;
                UpdateProgressLabel("Подготовка предпросмотра...");

                _cts = new CancellationTokenSource();
                var token = _cts.Token;

                var extensions = GetSelectedExtensions();
                var fileNames  = new ExcelFileNameReader(_logger).Read(_excelPath);

                _previewItems = await Task.Run(() => BuildPreview(fileNames, extensions, token), token);

                dgvPreview.Visible = true;
                txtLog.Visible = false;

                foreach (var item in _previewItems)
                {
                    string fileDate = item.FileDate.HasValue
                        ? item.FileDate.Value.ToString("dd.MM.yyyy HH:mm")
                        : "Не найден";
                    dgvPreview.Rows.Add(item.IsFound, item.FileName, item.Extension,
                        item.SourcePath ?? string.Empty, fileDate);
                }

                _total = _previewItems.Count;
                _notFound = _previewItems.Count(i => !i.IsFound);
                UpdateStats();

                UpdateProgressLabel(
                    $"Найдено {_previewItems.Count - _notFound} из {_previewItems.Count}. " +
                    "Снимите галочки с лишних и нажмите 'Запустить копирование'.");
            }
            catch (OperationCanceledException)
            {
                UpdateProgressLabel("Операция прервана пользователем.");
            }
            catch (Exception ex)
            {
                _logger.Error("Ошибка предпросмотра: " + ex.Message);
                UpdateProgressLabel("Ошибка при создании предпросмотра.");
            }
            finally
            {
                _cts?.Dispose();
                _cts = null;
                SetOperationState(false);
            }
        }

        private List<PreviewItem> BuildPreview(
            List<string> fileNames, List<string> extensions, CancellationToken ct)
        {
            if (_logger == null) return new List<PreviewItem>();

            var progress = new Progress<string>(folder => UpdateCurrentFolderLabel("Сканирование: " + folder));
            var idx = FileIndex.Build(_searchFolders, _logger, progress, SkipSystemHidden, ct);

            var result = new List<PreviewItem>(fileNames.Count * extensions.Count);
            int total = fileNames.Count * extensions.Count;
            int processed = 0;

            foreach (var name in fileNames)
            {
                ct.ThrowIfCancellationRequested();
                foreach (var ext in extensions)
                {
                    processed++;
                    var hit = idx.FindLatest(name, ext);
                    result.Add(new PreviewItem
                    {
                        FileName = name,
                        Extension = ext,
                        SourcePath = hit?.FullName,
                        FileDate = hit?.LastWriteTime
                    });
                    if (processed % 20 == 0 || processed == total)
                    {
                        UpdateProgressBar((int)(processed * 100.0 / total));
                        UpdateProgressLabel($"Поиск {processed} из {total}: {name}{ext}");
                    }
                }
            }

            UpdateCurrentFolderLabel(string.Empty);
            return result.OrderBy(p => p.FileName).ThenBy(p => p.Extension).ToList();
        }

        // ---------- Copy run ----------
        private async void BtnStart_Click(object? sender, EventArgs e)
        {
            if (_isRunning) return;
            if (_logger == null) return;

            ResetCounters();
            UpdateStats();

            bool usePreviewData = dgvPreview.Visible && dgvPreview.Rows.Count > 0;
            List<PreviewItem> queue;

            if (usePreviewData)
            {
                queue = CollectCheckedPreviewRows();
                if (queue.Count == 0)
                {
                    MessageBox.Show("Выберите хотя бы один файл для копирования.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            else
            {
                var extensions = GetSelectedExtensions();
                var fileNames  = new ExcelFileNameReader(_logger).Read(_excelPath);
                queue = await Task.Run(() => BuildPreview(fileNames, extensions, CancellationToken.None));
            }

            try
            {
                SetOperationState(true);
                _completedOps.Clear();

                dgvPreview.Visible = false;
                txtLog.Visible = true;
                txtLog.Clear();
                progressBar.Value = 0;

                UpdateProgressLabel("Начало обработки...");
                UpdateCurrentFolderLabel(string.Empty);

                _cts = new CancellationTokenSource();
                var token = _cts.Token;

                await Task.Run(() => RunCopy(queue, token), token);

                if (!token.IsCancellationRequested)
                {
                    new NotFoundReportWriter(_logger).Write(_destinationPath, _completedOps);
                    progressBar.Value = 100;
                    UpdateProgressLabel("Обработка завершена.");
                    UpdateCurrentFolderLabel(string.Empty);
                    MessageBox.Show(
                        $"Готово.\n\nСкопировано: {_copied}\nПропущено: {_skipped}\nНе найдено: {_notFound}",
                        "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (OperationCanceledException)
            {
                UpdateProgressLabel("Операция прервана пользователем.");
            }
            catch (Exception ex)
            {
                _logger.Error("Критическая ошибка: " + ex.Message);
                UpdateProgressLabel("Ошибка при выполнении операции.");
            }
            finally
            {
                _cts?.Dispose();
                _cts = null;
                SetOperationState(false);
                Cleanup();
            }
        }

        private async Task RunCopy(List<PreviewItem> queue, CancellationToken token)
        {
            if (_logger == null) return;
            var copier = new FileCopyService(_logger);

            _total = queue.Count;
            UpdateStats();
            UpdateProgressBar(0);

            for (int i = 0; i < queue.Count; i++)
            {
                token.ThrowIfCancellationRequested();
                var item = queue[i];

                int progress = (int)((i + 1) * 100.0 / Math.Max(queue.Count, 1));
                UpdateProgressBar(progress);
                UpdateProgressLabel($"Копирование {i + 1} из {queue.Count}: {item.FileName}{item.Extension}");

                if (string.IsNullOrEmpty(item.SourcePath) || !File.Exists(item.SourcePath))
                {
                    _completedOps.Add(new CopyOperation
                    {
                        FileName = item.FileName,
                        Extension = item.Extension,
                        SourcePath = item.SourcePath ?? string.Empty,
                        Status = CopyStatus.NotFound,
                        Message = "Файл не найден в папках поиска"
                    });
                    _notFound++;
                    UpdateStats();
                    continue;
                }

                var op = await copier.CopyAsync(
                    new FileInfo(item.SourcePath),
                    _destinationPath,
                    item.FileName,
                    item.Extension,
                    SkipLockedFiles,
                    token).ConfigureAwait(false);

                _completedOps.Add(op);
                switch (op.Status)
                {
                    case CopyStatus.Success: _copied++; break;
                    case CopyStatus.AlreadyAtDestination: _skipped++; break;
                    case CopyStatus.NotFound: _notFound++; break;
                    default: _notFound++; break;
                }
                UpdateStats();
            }
        }

        private List<PreviewItem> CollectCheckedPreviewRows()
        {
            var list = new List<PreviewItem>();
            foreach (DataGridViewRow row in dgvPreview.Rows)
            {
                if (row.Cells["Copy"] is not DataGridViewCheckBoxCell chk) continue;
                if (chk.Value is not bool b || !b) continue;

                string srcRaw = row.Cells["SourcePath"].Value?.ToString() ?? string.Empty;
                list.Add(new PreviewItem
                {
                    FileName  = row.Cells["FileName"].Value?.ToString()  ?? string.Empty,
                    Extension = row.Cells["Extension"].Value?.ToString() ?? string.Empty,
                    SourcePath = string.IsNullOrEmpty(srcRaw) ? null : srcRaw
                });
            }
            return list;
        }

        // ---------- Cancel ----------
        private void BtnCancel_Click(object? sender, EventArgs e)
        {
            if (!_isRunning || _cts == null) return;
            if (MessageBox.Show("Прервать текущую операцию?", "Подтверждение",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;
            _cts.Cancel();
            btnCancel.Enabled = false;
            btnCancel.Text = "Отмена...";
            UpdateProgressLabel("Завершение операции...");
        }

        // ---------- Preview context-menu actions ----------
        private void SelectAllPreviewItems(bool check)
        {
            foreach (DataGridViewRow row in dgvPreview.Rows)
                if (row.Cells["Copy"] is DataGridViewCheckBoxCell chk)
                    chk.Value = check;
        }

        private void UncheckSelectedPreviewItems()
        {
            foreach (DataGridViewRow row in dgvPreview.SelectedRows)
                if (row.Cells["Copy"] is DataGridViewCheckBoxCell chk)
                    chk.Value = false;
        }

        private string? GetSelectedPreviewFilePath()
        {
            if (dgvPreview.SelectedRows.Count == 0) return null;
            return dgvPreview.SelectedRows[0].Cells["SourcePath"].Value?.ToString();
        }

        private void OpenSelectedPreviewFile()
        {
            string? path = GetSelectedPreviewFilePath();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
            try
            {
                Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                _logger?.Error($"Ошибка открытия файла {path}: {ex.Message}");
            }
        }

        private void OpenSelectedPreviewFolder()
        {
            string? path = GetSelectedPreviewFilePath();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
            string? folder = Path.GetDirectoryName(path);
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder)) return;
            try { Process.Start("explorer.exe", $"\"{folder}\""); }
            catch (Exception ex) { _logger?.Error($"Ошибка открытия папки {folder}: {ex.Message}"); }
        }

        private void CopySelectedPreviewNamesToClipboard()
        {
            var names = dgvPreview.SelectedRows
                .Cast<DataGridViewRow>()
                .Select(r =>
                    (r.Cells["FileName"].Value?.ToString() ?? string.Empty) +
                    (r.Cells["Extension"].Value?.ToString() ?? string.Empty))
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (names.Count == 0) return;
            try { Clipboard.SetText(string.Join(Environment.NewLine, names)); }
            catch (Exception ex) { _logger?.Error("Ошибка буфера обмена: " + ex.Message); }
        }

        private void DgvPreview_CellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right || e.RowIndex < 0) return;
            if (!dgvPreview.Rows[e.RowIndex].Selected)
            {
                dgvPreview.ClearSelection();
                dgvPreview.Rows[e.RowIndex].Selected = true;
            }
            if (e.ColumnIndex >= 0)
                dgvPreview.CurrentCell = dgvPreview.Rows[e.RowIndex].Cells[e.ColumnIndex];
        }

        // ---------- Export ----------
        private async void ExportAsHtml()
        {
            if (_completedOps.Count == 0 && _previewItems.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта. Сначала выполните предпросмотр или копирование.",
                    "Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Filter = "HTML файл|*.html",
                FileName = $"FileCollector-report-{DateTime.Now:yyyyMMdd-HHmm}.html"
            };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            var plugin = _plugins?.GetPlugin<HtmlExportPlugin>();
            if (plugin == null)
            {
                MessageBox.Show("HTML-плагин не загружен.", "Экспорт",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var results = new SearchResults
            {
                SearchDate = DateTime.Now,
                TotalFound = _completedOps.Count > 0
                    ? _completedOps.Count(o => o.Status == CopyStatus.Success)
                    : _previewItems.Count(i => i.IsFound),
                NotFound = _completedOps.Count > 0
                    ? _completedOps.Count(o => o.Status != CopyStatus.Success && o.Status != CopyStatus.AlreadyAtDestination)
                    : _previewItems.Count(i => !i.IsFound),
                Files = (_completedOps.Count > 0
                        ? _completedOps.Where(o => o.Status == CopyStatus.Success)
                            .Select(o => new FileResult
                            {
                                FileName = o.FileName,
                                Extension = o.Extension,
                                SourcePath = o.SourcePath,
                                FileDate = File.Exists(o.SourcePath) ? new FileInfo(o.SourcePath).LastWriteTime : DateTime.MinValue,
                                FileSize = o.FileSize
                            })
                        : _previewItems.Where(i => i.IsFound)
                            .Select(i => new FileResult
                            {
                                FileName = i.FileName,
                                Extension = i.Extension,
                                SourcePath = i.SourcePath ?? string.Empty,
                                FileDate = i.FileDate ?? DateTime.MinValue,
                                FileSize = !string.IsNullOrEmpty(i.SourcePath) && File.Exists(i.SourcePath) ? new FileInfo(i.SourcePath).Length : 0
                            }))
                    .ToList()
            };

            try
            {
                await plugin.ExportAsync(results, sfd.FileName);
                MessageBox.Show("Отчёт сохранён:\n" + sfd.FileName, "Экспорт",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _logger?.Error("Ошибка экспорта HTML: " + ex.Message);
                MessageBox.Show("Ошибка экспорта: " + ex.Message, "Экспорт",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ---------- Helpers ----------
        private List<string> GetSelectedExtensions()
        {
            var list = new List<string>();
            foreach (var item in clbExtensions.CheckedItems)
            {
                string? s = item?.ToString();
                if (!string.IsNullOrWhiteSpace(s)) list.Add(s);
            }

            if (!string.IsNullOrWhiteSpace(txtCustomExtension.Text))
            {
                var parts = txtCustomExtension.Text.Split(
                    new[] { ',', ';', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var p in parts)
                {
                    string ext = p.Trim();
                    if (!ext.StartsWith('.')) ext = "." + ext;
                    if (!list.Contains(ext, StringComparer.OrdinalIgnoreCase))
                        list.Add(ext);
                }
            }

            if (list.Count == 0) list.Add(".pdf");
            return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private void UpdateStartButtonState()
        {
            bool ready = !string.IsNullOrEmpty(_excelPath)
                         && _searchFolders.Count > 0
                         && !string.IsNullOrEmpty(_destinationPath);

            btnPreview.Enabled = ready && !_isRunning;
            btnStart.Enabled = ready && !_isRunning;
            btnSelectExcel.Enabled = !_isRunning;
            btnAddSearchFolder.Enabled = !_isRunning;
            btnSelectDestination.Enabled = !_isRunning;
            clbExtensions.Enabled = !_isRunning;
            txtCustomExtension.Enabled = !_isRunning;
        }

        private void SetOperationState(bool running)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => SetOperationState(running)));
                return;
            }
            _isRunning = running;
            btnCancel.Visible = running;
            btnCancel.Enabled = running;
            btnCancel.Text = "Отмена";
            UpdateStartButtonState();
            UpdateSearchFolderButtons();
        }

        private void ResetCounters()
        {
            _total = _copied = _notFound = _skipped = 0;
        }

        private void Cleanup()
        {
            // Drop large transient buffers but keep _previewItems so HTML export still has data.
            _completedOps.TrimExcess();
        }

        private void AppendLog(string text)
        {
            if (txtLog.IsDisposed) return;
            if (txtLog.InvokeRequired)
            {
                try { txtLog.BeginInvoke(new Action<string>(AppendLog), text); }
                catch (InvalidOperationException) { /* shutting down */ }
                return;
            }
            txtLog.AppendText(text.EndsWith(Environment.NewLine) ? text : text + Environment.NewLine);
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
        }

        private void UpdateStats()
        {
            if (lblStats.InvokeRequired)
            {
                lblStats.BeginInvoke(new Action(UpdateStats));
                return;
            }
            lblStats.Text =
                $"Статистика: всего: {_total} | скопировано: {_copied} | " +
                $"пропущено: {_skipped} | не найдено: {_notFound}";
        }

        private void UpdateProgressBar(int value)
        {
            int v = Math.Max(0, Math.Min(value, 100));
            if (progressBar.InvokeRequired)
            {
                progressBar.BeginInvoke(new Action<int>(UpdateProgressBar), v);
                return;
            }
            progressBar.Value = v;
        }

        private void UpdateProgressLabel(string text)
        {
            if (lblProgress.InvokeRequired)
            {
                lblProgress.BeginInvoke(new Action<string>(UpdateProgressLabel), text);
                return;
            }
            lblProgress.Text = text;
        }

        private void UpdateCurrentFolderLabel(string text)
        {
            if (lblCurrentFolder.InvokeRequired)
            {
                lblCurrentFolder.BeginInvoke(new Action<string>(UpdateCurrentFolderLabel), text);
                return;
            }
            lblCurrentFolder.Text = text;
        }
    }
}
