// FileCollector V3.0 - layout half of MainForm (UI construction).
// Extracted from V2.1's monolithic MainForm.cs (was 2145 lines).

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using FileCollector.Core;

namespace FileCollector
{
    public partial class MainForm
    {
        // UI controls
        private Button btnSelectExcel = null!;
        private Button btnAddSearchFolder = null!;
        private Button btnRemoveSearchFolder = null!;
        private Button btnClearSearchFolders = null!;
        private Button btnSelectDestination = null!;
        private Button btnPreview = null!;
        private Button btnStart = null!;
        private Button btnCancel = null!;
        private CheckedListBox clbExtensions = null!;
        private TextBox txtCustomExtension = null!;
        private TextBox txtExcelPath = null!;
        private ListBox lstSearchFolders = null!;
        private TextBox txtDestinationPath = null!;
        private DataGridView dgvPreview = null!;
        private TextBox txtLog = null!;
        private Label lblStats = null!;
        private ProgressBar progressBar = null!;
        private Label lblProgress = null!;
        private Label lblCurrentFolder = null!;
        private ToolTip toolTip = null!;
        private ContextMenuStrip folderContextMenu = null!;
        private ContextMenuStrip previewContextMenu = null!;
        private MenuStrip mainMenu = null!;
        private int lastExtensionTooltipIndex = -1;

        private const string ExtensionsToolTipText =
            "Выберите форматы файлов для поиска\n" +
            "* По умолчанию отмечены: .pdf, .dxf\n" +
            "* SolidWorks: .sldprt, .sldasm, .slddrw\n" +
            "* Нейтральные CAD: .step, .stp, .iges\n" +
            "* Inventor: .ipt, .iam, .idw\n" +
            "* Creo: .prt, .asm, .drw\n" +
            "* CATIA: .catpart, .catproduct\n" +
            "* Solid Edge: .par, .psm";

        private static readonly Dictionary<string, string> SolidWorksExtensionTooltips =
            new(StringComparer.OrdinalIgnoreCase)
            {
                [".sldprt"] = "SolidWorks: .sldprt - Деталь",
                [".sldasm"] = "SolidWorks: .sldasm - Сборка",
                [".slddrw"] = "SolidWorks: .slddrw - Чертёж"
            };

        private static readonly string[] DefaultExtensions =
        {
            ".pdf", ".dxf",
            ".sldprt", ".sldasm", ".slddrw",
            ".dwg", ".doc", ".docx", ".xlsx", ".xls", ".txt", ".jpg", ".png",
            ".step", ".stp", ".iges", ".igs",
            ".ipt", ".iam", ".idw",
            ".prt", ".asm", ".drw",
            ".catpart", ".catproduct", ".catdrawing",
            ".par", ".psm", ".dft",
            ".3dm", ".skp", ".dgn", ".rvt", ".rfa", ".rte"
        };

        private void InitializeComponents()
        {
            int margin = 10;
            int currentY = margin;
            int labelWidth = 140;
            int buttonWidth = 130;
            int fullWidth = this.ClientSize.Width - 2 * margin;

            BuildMenu();
            currentY += mainMenu.Height; // shift below menu

            // ---------- Excel selection ----------
            btnSelectExcel = NewButton("Выбрать Excel", margin, currentY, buttonWidth, 28);
            btnSelectExcel.Click += BtnSelectExcel_Click;
            this.Controls.Add(btnSelectExcel);

            txtExcelPath = new TextBox
            {
                Location = new Point(margin + buttonWidth + 5, currentY),
                Width = fullWidth - buttonWidth - 5,
                Height = 28,
                ReadOnly = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(txtExcelPath);

            currentY += 35;

            // ---------- Search folders ----------
            btnAddSearchFolder = NewButton("Добавить папку", margin, currentY, buttonWidth, 28);
            btnAddSearchFolder.Click += BtnAddSearchFolder_Click;
            this.Controls.Add(btnAddSearchFolder);

            btnRemoveSearchFolder = NewButton("Удалить", margin, currentY + 32, buttonWidth, 28);
            btnRemoveSearchFolder.Enabled = false;
            btnRemoveSearchFolder.Click += BtnRemoveSearchFolder_Click;
            this.Controls.Add(btnRemoveSearchFolder);

            btnClearSearchFolders = NewButton("Очистить все", margin, currentY + 64, buttonWidth, 28);
            btnClearSearchFolders.Enabled = false;
            btnClearSearchFolders.Click += BtnClearSearchFolders_Click;
            this.Controls.Add(btnClearSearchFolders);

            lstSearchFolders = new ListBox
            {
                Location = new Point(margin + buttonWidth + 5, currentY),
                Width = fullWidth - buttonWidth - 5,
                Height = 90,
                HorizontalScrollbar = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            lstSearchFolders.SelectedIndexChanged += LstSearchFolders_SelectedIndexChanged;
            this.Controls.Add(lstSearchFolders);

            currentY += 100;

            // ---------- Extensions ----------
            this.Controls.Add(new Label
            {
                Text = "Форматы для поиска:",
                Location = new Point(margin, currentY + 5),
                Width = labelWidth,
                Height = 20
            });

            clbExtensions = new CheckedListBox
            {
                Location = new Point(margin + labelWidth, currentY),
                Width = 300,
                Height = 120,
                CheckOnClick = true,
                MultiColumn = true,
                ColumnWidth = 100
            };
            foreach (var ext in DefaultExtensions)
                clbExtensions.Items.Add(ext);
            // Default: pdf and dxf checked.
            clbExtensions.SetItemChecked(0, true);
            clbExtensions.SetItemChecked(1, true);
            this.Controls.Add(clbExtensions);

            this.Controls.Add(new Label
            {
                Text = "Доп. форматы (через запятую):",
                Location = new Point(margin + labelWidth + 305, currentY + 5),
                Width = 180,
                Height = 20
            });

            txtCustomExtension = new TextBox
            {
                Location = new Point(margin + labelWidth + 305, currentY + 25),
                Width = fullWidth - labelWidth - 305,
                Height = 28,
                PlaceholderText = ".ifc, .sat, .x_t, .x_b, .jt, .stl, .fbx"
            };
            this.Controls.Add(txtCustomExtension);

            currentY += 130;

            // ---------- Destination ----------
            this.Controls.Add(new Label
            {
                Text = "Папка для сбора:",
                Location = new Point(margin, currentY + 5),
                Width = labelWidth,
                Height = 20
            });

            btnSelectDestination = NewButton("Выбрать папку", margin + labelWidth, currentY, buttonWidth, 28);
            btnSelectDestination.Click += BtnSelectDestination_Click;
            this.Controls.Add(btnSelectDestination);

            txtDestinationPath = new TextBox
            {
                Location = new Point(margin + labelWidth + buttonWidth + 5, currentY),
                Width = fullWidth - labelWidth - buttonWidth - 5,
                Height = 28,
                ReadOnly = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(txtDestinationPath);

            currentY += 35;

            // ---------- Action buttons ----------
            btnPreview = NewButton("Предпросмотр", margin, currentY, 130, 35);
            btnPreview.Enabled = false;
            btnPreview.Click += BtnPreview_Click;
            this.Controls.Add(btnPreview);

            btnStart = NewButton("Запустить копирование", margin + 135, currentY, 160, 35);
            btnStart.Enabled = false;
            btnStart.Click += BtnStart_Click;
            this.Controls.Add(btnStart);

            btnCancel = NewButton("Отмена", margin + 300, currentY, 120, 35);
            btnCancel.Visible = false;
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);

            currentY += 45;

            // ---------- Progress ----------
            progressBar = new ProgressBar
            {
                Location = new Point(margin, currentY),
                Width = fullWidth,
                Height = 25,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Style = ProgressBarStyle.Continuous
            };
            this.Controls.Add(progressBar);

            currentY += 30;

            lblProgress = new Label
            {
                Location = new Point(margin, currentY),
                Width = fullWidth,
                Height = 20,
                Text = "Готов к работе",
                Font = new Font(this.Font, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblProgress);

            currentY += 25;

            lblCurrentFolder = new Label
            {
                Location = new Point(margin, currentY),
                Width = fullWidth,
                Height = 20,
                Text = string.Empty,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblCurrentFolder);

            currentY += 30;

            // ---------- Preview / log ----------
            dgvPreview = new DataGridView
            {
                Location = new Point(margin, currentY),
                Width = fullWidth,
                Height = 180,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                ReadOnly = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ScrollBars = ScrollBars.Vertical,
                RowHeadersVisible = false,
                Visible = false
            };
            dgvPreview.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Copy", HeaderText = "Копировать", Width = 80 });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "FileName", HeaderText = "Имя файла", Width = 150, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "Extension", HeaderText = "Формат", Width = 80, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "SourcePath", HeaderText = "Откуда", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 220, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "FileDate", HeaderText = "Дата", Width = 120, ReadOnly = true });
            this.Controls.Add(dgvPreview);

            txtLog = new TextBox
            {
                Location = new Point(margin, currentY),
                Width = fullWidth,
                Height = 180,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Font = new Font("Consolas", 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(txtLog);

            currentY += 185;

            lblStats = new Label
            {
                Location = new Point(margin, currentY),
                Width = fullWidth,
                Height = 30,
                Text = "Статистика: всего: 0 | скопировано: 0 | не найдено: 0",
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblStats);
        }

        private void BuildMenu()
        {
            mainMenu = new MenuStrip();
            var miFile = new ToolStripMenuItem("Файл");
            miFile.DropDownItems.Add(new ToolStripMenuItem("Выйти", null, (_, __) => this.Close()));
            mainMenu.Items.Add(miFile);

            var miExport = new ToolStripMenuItem("Экспорт");
            miExport.DropDownItems.Add(new ToolStripMenuItem("HTML отчёт...", null, (_, __) => ExportAsHtml())
            {
                Name = "miExportHtml"
            });
            mainMenu.Items.Add(miExport);

            var miHelp = new ToolStripMenuItem("Справка");
            miHelp.DropDownItems.Add(new ToolStripMenuItem("О программе", null, (_, __) => ShowAboutBox()));
            mainMenu.Items.Add(miHelp);

            this.MainMenuStrip = mainMenu;
            this.Controls.Add(mainMenu);
        }

        private void ShowAboutBox()
        {
            MessageBox.Show(
                $"{VersionInfo.ShortTitle}\n" +
                $"Кодовое имя: {VersionInfo.CodeName}\n" +
                $"Компания: {VersionInfo.Company}\n\n" +
                "Архитектурный пересбор V3.0:\n" +
                "* монолит разделён на Core/ (бизнес-логика) и UI/\n" +
                "* индекс файлов вместо повторных обходов\n" +
                "* асинхронное копирование с защитой от path-traversal\n" +
                "* пакетный логгер (UI больше не блокируется)\n" +
                "* подключённая система плагинов",
                "О программе",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static Button NewButton(string text, int x, int y, int w, int h) => new()
        {
            Text = text,
            Location = new Point(x, y),
            Width = w,
            Height = h
        };

        private void SetupToolTips()
        {
            toolTip = new ToolTip
            {
                AutoPopDelay = 5000,
                InitialDelay = 500,
                ReshowDelay = 250,
                ShowAlways = true
            };
            toolTip.SetToolTip(btnSelectExcel, "Выберите Excel файл с именами файлов (столбец B)");
            toolTip.SetToolTip(txtExcelPath, "Путь к Excel файлу с именами файлов для поиска");
            toolTip.SetToolTip(btnAddSearchFolder, "Добавить папку для поиска файлов");
            toolTip.SetToolTip(btnRemoveSearchFolder, "Удалить выбранную папку из списка поиска");
            toolTip.SetToolTip(btnClearSearchFolders, "Очистить весь список папок для поиска");
            toolTip.SetToolTip(clbExtensions, ExtensionsToolTipText);
            clbExtensions.MouseMove -= ClbExtensions_MouseMove;
            clbExtensions.MouseMove += ClbExtensions_MouseMove;
            clbExtensions.MouseLeave -= ClbExtensions_MouseLeave;
            clbExtensions.MouseLeave += ClbExtensions_MouseLeave;
            toolTip.SetToolTip(txtCustomExtension, "Дополнительные форматы через запятую\nНапример: .ifc, .sat, .x_t, .jt, .stl, .fbx");
            toolTip.SetToolTip(btnSelectDestination, "Выберите папку, куда будут скопированы найденные файлы");
            toolTip.SetToolTip(btnPreview, "Предварительный просмотр найденных файлов\nПозволяет выбрать, какие файлы копировать");
            toolTip.SetToolTip(btnStart, "Начать копирование выбранных файлов");
            toolTip.SetToolTip(btnCancel, "Прервать текущую операцию");
        }

        private void SetupContextMenus()
        {
            folderContextMenu = new ContextMenuStrip();
            folderContextMenu.Items.Add("Удалить", null, (s, ev) => BtnRemoveSearchFolder_Click(s, ev));
            folderContextMenu.Items.Add("Очистить все", null, (s, ev) => BtnClearSearchFolders_Click(s, ev));
            folderContextMenu.Items.Add("-");
            folderContextMenu.Items.Add("Открыть в проводнике", null, (s, ev) =>
            {
                if (lstSearchFolders.SelectedItem is string p && Directory.Exists(p))
                    System.Diagnostics.Process.Start("explorer.exe", p);
            });
            lstSearchFolders.ContextMenuStrip = folderContextMenu;

            previewContextMenu = new ContextMenuStrip();
            previewContextMenu.Items.Add("Выбрать все", null, (_, __) => SelectAllPreviewItems(true));
            previewContextMenu.Items.Add("Снять все", null, (_, __) => SelectAllPreviewItems(false));
            previewContextMenu.Items.Add("Снять галочку у выделенных", null, (_, __) => UncheckSelectedPreviewItems());
            previewContextMenu.Items.Add("-");
            previewContextMenu.Items.Add("Открыть файл", null, (_, __) => OpenSelectedPreviewFile());
            previewContextMenu.Items.Add("Открыть папку", null, (_, __) => OpenSelectedPreviewFolder());
            previewContextMenu.Items.Add("Копировать имя файла", null, (_, __) => CopySelectedPreviewNamesToClipboard());
            dgvPreview.ContextMenuStrip = previewContextMenu;
            dgvPreview.CellMouseDown -= DgvPreview_CellMouseDown;
            dgvPreview.CellMouseDown += DgvPreview_CellMouseDown;
        }

        private void SetupDragDrop()
        {
            this.AllowDrop = true;
            txtExcelPath.AllowDrop = true;
            txtDestinationPath.AllowDrop = true;
            lstSearchFolders.AllowDrop = true;

            txtExcelPath.DragEnter += (s, e) =>
            {
                if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                    e.Effect = DragDropEffects.Copy;
            };
            txtExcelPath.DragDrop += (s, e) =>
            {
                if (e.Data?.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0
                    && (files[0].EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                        || files[0].EndsWith(".xls", StringComparison.OrdinalIgnoreCase)))
                {
                    _excelPath = files[0];
                    txtExcelPath.Text = _excelPath;
                    UpdateStartButtonState();
                }
            };

            txtDestinationPath.DragEnter += (s, e) =>
            {
                if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                    e.Effect = DragDropEffects.Copy;
            };
            txtDestinationPath.DragDrop += (s, e) =>
            {
                if (e.Data?.GetData(DataFormats.FileDrop) is not string[] items || items.Length == 0)
                    return;
                string item = items[0];
                if (File.Exists(item))
                    item = Path.GetDirectoryName(item) ?? item;
                if (Directory.Exists(item))
                {
                    _destinationPath = item;
                    txtDestinationPath.Text = _destinationPath;
                    UpdateStartButtonState();
                }
            };

            lstSearchFolders.DragEnter += (s, e) =>
            {
                if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                    e.Effect = DragDropEffects.Copy;
            };
            lstSearchFolders.DragDrop += (s, e) =>
            {
                if (e.Data?.GetData(DataFormats.FileDrop) is not string[] items) return;
                foreach (var item in items)
                {
                    if (Directory.Exists(item) && !_searchFolders.Contains(item))
                    {
                        _searchFolders.Add(item);
                        lstSearchFolders.Items.Add(item);
                        UpdateSearchFoldersHorizontalExtent();
                    }
                }
                UpdateStartButtonState();
                UpdateSearchFolderButtons();
            };
        }

        private void SetupKeyboardShortcuts()
        {
            this.KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.O && !_isRunning)
                {
                    e.SuppressKeyPress = true;
                    BtnSelectExcel_Click(s, e);
                }
                else if (e.Control && e.KeyCode == Keys.P && !_isRunning)
                {
                    e.SuppressKeyPress = true;
                    BtnPreview_Click(s, e);
                }
                else if (e.Control && e.KeyCode == Keys.S && !_isRunning)
                {
                    e.SuppressKeyPress = true;
                    BtnStart_Click(s, e);
                }
                else if (e.KeyCode == Keys.Escape && _isRunning)
                {
                    e.SuppressKeyPress = true;
                    BtnCancel_Click(s, e);
                }
            };
        }
    }
}
