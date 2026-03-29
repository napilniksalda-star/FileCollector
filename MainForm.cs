using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using FileCollector.Plugins;
using FileCollector.Plugins.BuiltIn;

namespace FileCollector
{
    public class MainForm : Form
    {
        // UI Controls
        private Button? btnSelectExcel = null;
        private Button? btnAddSearchFolder = null;
        private Button? btnRemoveSearchFolder = null;
        private Button? btnClearSearchFolders = null;
        private Button? btnSelectDestination = null;
        private Button? btnPreview = null;
        private Button? btnStart = null;
        private Button? btnCancel = null;
        private CheckedListBox? clbExtensions = null;
        private TextBox? txtCustomExtension = null;
        private TextBox? txtExcelPath = null;
        private ListBox? lstSearchFolders = null;
        private TextBox? txtDestinationPath = null;
        private DataGridView? dgvPreview = null;
        private TextBox? txtLog = null;
        private Label? lblStats = null;
        private ProgressBar? progressBar = null;
        private Label? lblProgress = null;
        private Label? lblCurrentFolder = null;
        private ToolTip? toolTip = null;
        private ContextMenuStrip? folderContextMenu = null;
        private ContextMenuStrip? previewContextMenu = null;
        private PictureBox? pbLogo = null;
        private int lastExtensionTooltipIndex = -1;

        private const string ExtensionsToolTipText =
            "Выберите форматы файлов для поиска\n" +
            "• По умолчанию отмечены: .pdf, .dxf\n" +
            "• SolidWorks: .sldprt (Деталь), .sldasm (Сборка), .slddrw (Чертеж)\n" +
            "• Нейтральные CAD: .step, .stp, .iges\n" +
            "• Inventor: .ipt, .iam, .idw\n" +
            "• Creo: .prt, .asm, .drw\n" +
            "• CATIA: .catpart, .catproduct\n" +
            "• Solid Edge: .par, .psm\n" +
            "• Другие САПР и документы";

        private static readonly Dictionary<string, string> SolidWorksExtensionTooltips =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                [".sldprt"] = "SolidWorks: .sldprt - Деталь",
                [".sldasm"] = "SolidWorks: .sldasm - Сборка",
                [".slddrw"] = "SolidWorks: .slddrw - Чертеж"
            };

        // State
        private List<string> searchFolders = new List<string>();
        private string excelPath = string.Empty;
        private string destinationPath = string.Empty;
        private string logoLoadedPath = string.Empty;  // Для отладки
        private int totalFiles = 0;
        private int copiedFiles = 0;
        private int notFoundFiles = 0;
        private int skippedFiles = 0;
        private List<PreviewItem> previewItems = new List<PreviewItem>();
        private List<CopyOperation> completedOperations = new List<CopyOperation>();
        private CancellationTokenSource? cancellationTokenSource = null;
        private bool isOperationRunning = false;

        // Default settings (hidden from user)
        private bool skipLockedFiles = true;       // Пропускать заблокированные файлы
        private bool skipSystemHidden = true;      // Пропускать системные/скрытые файлы
        private bool continueOnError = false;      // Продолжать при ошибках

        // Cache for performance
        private static readonly ConcurrentDictionary<string, FileSystemCacheEntry> fileSystemCache = 
            new ConcurrentDictionary<string, FileSystemCacheEntry>();
        // Lock used when updating the shared latestFile in parallel searches
        private readonly object latestFileLock = new object();

        public MainForm()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Сборщик файлов V2.1";
            this.Size = new Size(900, 700);
            this.MinimumSize = new Size(850, 600);
            this.KeyPreview = true;

            // Применяем стандартный Windows стиль
            this.BackColor = SystemColors.Control;
            this.ForeColor = SystemColors.ControlText;

            InitializeComponents();
            SetupToolTips();
            SetupContextMenus();
            SetupDragDrop();
            SetupKeyboardShortcuts();

            // Добавляем начальную инструкцию в логи
            DisplayInitialInstructions();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void DisplayInitialInstructions()
        {
            string instructions = @"========== ИНСТРУКЦИЯ ПОЛЬЗОВАНИЯ ==========

1. ВЫБОР ФАЙЛА EXCEL:
   • Нажмите кнопку '📁 Выбрать...' или перетащите файл Excel
   • Поддерживаемые форматы: .xlsx, .xls
   • Файл должен содержать список имён файлов в столбце B

2. ПАПКИ ДЛЯ ПОИСКА:
   • Нажмите '➕ Добавить папку' или перетащите несколько папок в список
   • Программа будет искать файлы во всех папках и подпапках
   • Используйте '❌ Удалить' для удаления папки из списка

3. ПАПКА ДЛЯ СОХРАНЕНИЯ:
   • Нажмите '📁 Выбрать...' или перетащите папку
   • Найденные файлы будут скопированы в эту папку

4. ВЫБОР ФОРМАТОВ ФАЙЛОВ:
   • Отметьте нужные расширения (PDF, DXF и т.д.)
   • Программа будет искать файлы только с выбранными расширениями
   • Если в Excel указаны расширения, они удаляются автоматически

5. ЗАПУСК ПОИСКА И КОПИРОВАНИЯ:
   • Нажмите кнопку 'Старт' (Ctrl+S)
   • Программа найдёт и скопирует все файлы
   • Прогресс отображается в нижней части окна

6. ПРЕДПРОСМОТР:
   • Нажмите 'Предпросмотр' (Ctrl+P) для проверки найденных файлов
   • Параметры: фильтр по статусу, сортировка, поиск

⚙️ СОЧЕТАНИЯ КЛАВИШ:
   • Ctrl+O: Открыть Excel файл
   • Ctrl+P: Предпросмотр найденных файлов
   • Ctrl+S: Запустить процесс
   • Esc: Отменить операцию

💡 ПОДСКАЗКИ:
   • Перетаскивайте файлы и папки прямо в поля
   • Программа автоматически удаляет спецсимволы из Excel
   • Если файл не найден, проверьте логи ниже

✅ Готовы начать? Выполните шаги 1-4 выше!
============================================";

            LogMessage(instructions);
        }

        private void InitializeComponents()
        {
            int margin = 10;
            int currentY = margin;
            int labelWidth = 140;
            int buttonWidth = 130;
            int fullWidth = this.ClientSize.Width - 2 * margin;

            // Логотип
            pbLogo = new PictureBox
            {
                Location = new Point(fullWidth - 120, margin),
                Width = 110,
                Height = 21,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                SizeMode = PictureBoxSizeMode.Zoom,
                Visible = false  // Скрываем логотип для стандартного стиля
            };
            
            this.Controls.Add(pbLogo);

            // Excel Selection
            btnSelectExcel = new Button
            {
                Text = "Выбрать Excel",
                Location = new Point(margin, currentY),
                Width = buttonWidth,
                Height = 28
            };
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

            // Search Folders
            btnAddSearchFolder = new Button
            {
                Text = "Добавить папку",
                Location = new Point(margin, currentY),
                Width = buttonWidth,
                Height = 28
            };
            btnAddSearchFolder.Click += BtnAddSearchFolder_Click;
            this.Controls.Add(btnAddSearchFolder);

            btnRemoveSearchFolder = new Button
            {
                Text = "Удалить",
                Location = new Point(margin, currentY + 32),
                Width = buttonWidth,
                Height = 28,
                Enabled = false
            };
            btnRemoveSearchFolder.Click += BtnRemoveSearchFolder_Click;
            this.Controls.Add(btnRemoveSearchFolder);

            btnClearSearchFolders = new Button
            {
                Text = "Очистить все",
                Location = new Point(margin, currentY + 64),
                Width = buttonWidth,
                Height = 28,
                Enabled = false
            };
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

            // File Formats
            Label lblFormat = new Label
            {
                Text = "Форматы для поиска:",
                Location = new Point(margin, currentY + 5),
                Width = labelWidth,
                Height = 20
            };
            this.Controls.Add(lblFormat);

            clbExtensions = new CheckedListBox
            {
                Location = new Point(margin + labelWidth, currentY),
                Width = 300,
                Height = 120,
                CheckOnClick = true,
                MultiColumn = true,
                ColumnWidth = 100
            };
            
            // Добавляем форматы в нужном порядке
            // Первые два - pdf и dxf (будут отмечены галочками)
            clbExtensions.Items.Add(".pdf");
            clbExtensions.Items.Add(".dxf");
            
            // Затем форматы SolidWorks (без галочек)
            clbExtensions.Items.Add(".sldprt");
            clbExtensions.Items.Add(".sldasm");
            clbExtensions.Items.Add(".slddrw");
            
            // Затем остальные форматы САПР и документов (без галочек)
            clbExtensions.Items.Add(".dwg");
            clbExtensions.Items.Add(".doc");
            clbExtensions.Items.Add(".docx");
            clbExtensions.Items.Add(".xlsx");
            clbExtensions.Items.Add(".xls");
            clbExtensions.Items.Add(".txt");
            clbExtensions.Items.Add(".jpg");
            clbExtensions.Items.Add(".png");
            clbExtensions.Items.Add(".step");
            clbExtensions.Items.Add(".stp");
            clbExtensions.Items.Add(".iges");
            clbExtensions.Items.Add(".igs");
            clbExtensions.Items.Add(".ipt");
            clbExtensions.Items.Add(".iam");
            clbExtensions.Items.Add(".idw");
            clbExtensions.Items.Add(".prt");
            clbExtensions.Items.Add(".asm");
            clbExtensions.Items.Add(".drw");
            clbExtensions.Items.Add(".catpart");
            clbExtensions.Items.Add(".catproduct");
            clbExtensions.Items.Add(".catdrawing");
            clbExtensions.Items.Add(".par");
            clbExtensions.Items.Add(".psm");
            clbExtensions.Items.Add(".dft");
            clbExtensions.Items.Add(".3dm");
            clbExtensions.Items.Add(".skp");
            clbExtensions.Items.Add(".dgn");
            clbExtensions.Items.Add(".rvt");
            clbExtensions.Items.Add(".rfa");
            clbExtensions.Items.Add(".rte");
            
            // Устанавливаем галочки только для pdf и dxf
            clbExtensions.SetItemChecked(0, true); // .pdf
            clbExtensions.SetItemChecked(1, true); // .dxf
            
            this.Controls.Add(clbExtensions);

            Label lblCustomFormat = new Label
            {
                Text = "Доп. форматы (через запятую):",
                Location = new Point(margin + labelWidth + 305, currentY + 5),
                Width = 180,
                Height = 20
            };
            this.Controls.Add(lblCustomFormat);

            txtCustomExtension = new TextBox
            {
                Location = new Point(margin + labelWidth + 305, currentY + 25),
                Width = fullWidth - labelWidth - 305,
                Height = 28,
                PlaceholderText = ".ifc, .sat, .x_t, .x_b, .jt, .u3d, .dae, .fbx, .obj, .stl"
            };
            this.Controls.Add(txtCustomExtension);

            currentY += 130;

            // Destination Folder
            Label lblDestination = new Label
            {
                Text = "Папка для сбора:",
                Location = new Point(margin, currentY + 5),
                Width = labelWidth,
                Height = 20
            };
            this.Controls.Add(lblDestination);

            btnSelectDestination = new Button
            {
                Text = "Выбрать папку",
                Location = new Point(margin + labelWidth, currentY),
                Width = buttonWidth,
                Height = 28
            };
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

            // Action Buttons
            btnPreview = new Button
            {
                Text = "Предпросмотр",
                Location = new Point(margin, currentY),
                Width = 130,
                Height = 35,
                Enabled = false
            };
            btnPreview.Click += BtnPreview_Click;
            this.Controls.Add(btnPreview);

            btnStart = new Button
            {
                Text = "Запустить копирование",
                Location = new Point(margin + 135, currentY),
                Width = 160,
                Height = 35,
                Enabled = false
            };
            btnStart.Click += BtnStart_Click;
            this.Controls.Add(btnStart);

            btnCancel = new Button
            {
                Text = "Отмена",
                Location = new Point(margin + 300, currentY),
                Width = 120,
                Height = 35,
                Visible = false
            };
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);

            currentY += 45;

            // Progress
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
                Text = "",
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblCurrentFolder);

            currentY += 30;

            // Preview/Log area
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
            
            dgvPreview.Columns.Add(new DataGridViewCheckBoxColumn { 
                Name = "Copy", 
                HeaderText = "Копировать", 
                Width = 80 
            });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { 
                Name = "FileName", 
                HeaderText = "Имя файла", 
                Width = 150, 
                ReadOnly = true 
            });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { 
                Name = "Extension", 
                HeaderText = "Формат", 
                Width = 80, 
                ReadOnly = true 
            });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { 
                Name = "SourcePath", 
                HeaderText = "Откуда", 
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 220,
                ReadOnly = true 
            });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { 
                Name = "FileDate", 
                HeaderText = "Дата", 
                Width = 120, 
                ReadOnly = true 
            });
            
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

            // Statistics
            lblStats = new Label
            {
                Location = new Point(margin, currentY),
                Width = fullWidth,
                Height = 30,
                Text = "Статистика: Всего файлов: 0 | Скопировано: 0 | Не найдено: 0",
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblStats);
        }

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

        private void ClbExtensions_MouseMove(object? sender, MouseEventArgs e)
        {
            if (toolTip == null || clbExtensions == null)
                return;

            int index = clbExtensions.IndexFromPoint(e.Location);
            if (index < 0 || index == lastExtensionTooltipIndex)
                return;

            lastExtensionTooltipIndex = index;
            string extension = clbExtensions.Items[index]?.ToString() ?? string.Empty;

            if (SolidWorksExtensionTooltips.TryGetValue(extension, out string? extensionTip))
            {
                toolTip.SetToolTip(clbExtensions, extensionTip);
                toolTip.Show(extensionTip, clbExtensions, e.Location.X + 14, e.Location.Y + 18, 2500);
            }
            else
            {
                toolTip.SetToolTip(clbExtensions, ExtensionsToolTipText);
                toolTip.Hide(clbExtensions);
            }
        }

        private void ClbExtensions_MouseLeave(object? sender, EventArgs e)
        {
            if (toolTip == null || clbExtensions == null)
                return;

            lastExtensionTooltipIndex = -1;
            toolTip.Hide(clbExtensions);
            toolTip.SetToolTip(clbExtensions, ExtensionsToolTipText);
        }

        private void SetupContextMenus()
        {
            // Context menu for folder list
            folderContextMenu = new ContextMenuStrip();
            folderContextMenu.Items.Add("Удалить", null, (s, ev) => BtnRemoveSearchFolder_Click(s, ev));
            folderContextMenu.Items.Add("Очистить все", null, (s, ev) => BtnClearSearchFolders_Click(s, ev));
            folderContextMenu.Items.Add("-");
            folderContextMenu.Items.Add("Открыть в проводнике", null, (s, ev) =>
            {
                if (lstSearchFolders.SelectedItem != null)
                {
                    string path = lstSearchFolders.SelectedItem.ToString();
                    if (Directory.Exists(path))
                        System.Diagnostics.Process.Start("explorer.exe", path);
                }
            });
            lstSearchFolders.ContextMenuStrip = folderContextMenu;

            // Context menu for preview grid
            previewContextMenu = new ContextMenuStrip();
            previewContextMenu.Items.Add("Выбрать все", null, (s, ev) => SelectAllPreviewItems(true));
            previewContextMenu.Items.Add("Снять все", null, (s, ev) => SelectAllPreviewItems(false));
            previewContextMenu.Items.Add("Снять галочку у выделенных", null, (s, ev) => UncheckSelectedPreviewItems());
            previewContextMenu.Items.Add("-");
            previewContextMenu.Items.Add("Открыть файл", null, (s, ev) => OpenSelectedPreviewFile());
            previewContextMenu.Items.Add("Открыть папку", null, (s, ev) => OpenSelectedPreviewFolder());
            previewContextMenu.Items.Add("Копировать имя файла", null, (s, ev) => CopySelectedPreviewNamesToClipboard());
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
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    e.Effect = DragDropEffects.Copy;
            };

            txtExcelPath.DragDrop += (s, e) =>
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && (files[0].EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                    files[0].EndsWith(".xls", StringComparison.OrdinalIgnoreCase)))
                {
                    excelPath = files[0];
                    txtExcelPath.Text = excelPath;
                    UpdateStartButtonState();
                }
            };

            // Drag-and-drop для папки назначения
            txtDestinationPath.DragEnter += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    e.Effect = DragDropEffects.Copy;
            };

            txtDestinationPath.DragDrop += (s, e) =>
            {
                string[] items = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (items.Length > 0)
                {
                    string item = items[0];
                    // Если это файл, берём папку, в которой он находится
                    if (File.Exists(item))
                    {
                        item = Path.GetDirectoryName(item) ?? item;
                    }

                    if (Directory.Exists(item))
                    {
                        destinationPath = item;
                        txtDestinationPath.Text = destinationPath;
                        UpdateStartButtonState();
                    }
                }
            };

            lstSearchFolders.DragEnter += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    e.Effect = DragDropEffects.Copy;
            };

            lstSearchFolders.DragDrop += (s, e) =>
            {
                string[] items = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string item in items)
                {
                    if (Directory.Exists(item) && !searchFolders.Contains(item))
                    {
                        searchFolders.Add(item);
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
                if (e.Control && e.KeyCode == Keys.O && !isOperationRunning)
                {
                    e.SuppressKeyPress = true;
                    BtnSelectExcel_Click(s, e);
                }
                else if (e.Control && e.KeyCode == Keys.P && !isOperationRunning)
                {
                    e.SuppressKeyPress = true;
                    BtnPreview_Click(s, e);
                }
                else if (e.Control && e.KeyCode == Keys.S && !isOperationRunning)
                {
                    e.SuppressKeyPress = true;
                    BtnStart_Click(s, e);
                }
                else if (e.KeyCode == Keys.Escape && isOperationRunning)
                {
                    e.SuppressKeyPress = true;
                    BtnCancel_Click(s, e);
                }
            };
        }

        #region Event Handlers

        private void BtnSelectExcel_Click(object? sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*";
                ofd.Title = "Выберите Excel файл";
                ofd.CheckFileExists = true;
                
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    excelPath = ofd.FileName;
                    txtExcelPath.Text = excelPath;
                    UpdateStartButtonState();
                }
            }
        }

        private void BtnAddSearchFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Выберите папку для поиска файлов";
                fbd.ShowNewFolderButton = false;
                
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    if (!searchFolders.Contains(fbd.SelectedPath))
                    {
                        searchFolders.Add(fbd.SelectedPath);
                        lstSearchFolders.Items.Add(fbd.SelectedPath);
                        UpdateSearchFoldersHorizontalExtent();
                        UpdateStartButtonState();
                        UpdateSearchFolderButtons();
                    }
                    else
                    {
                        MessageBox.Show("Эта папка уже добавлена в список", "Информация", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void BtnRemoveSearchFolder_Click(object sender, EventArgs e)
        {
            if (lstSearchFolders.SelectedIndex >= 0)
            {
                int selectedIndex = lstSearchFolders.SelectedIndex;
                searchFolders.RemoveAt(selectedIndex);
                lstSearchFolders.Items.RemoveAt(selectedIndex);
                UpdateSearchFoldersHorizontalExtent();
                UpdateStartButtonState();
                UpdateSearchFolderButtons();
            }
        }

        private void BtnClearSearchFolders_Click(object sender, EventArgs e)
        {
            if (searchFolders.Count > 0)
            {
                DialogResult result = MessageBox.Show(
                    "Удалить все папки из списка поиска?",
                    "Подтверждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    searchFolders.Clear();
                    lstSearchFolders.Items.Clear();
                    UpdateSearchFoldersHorizontalExtent();
                    UpdateStartButtonState();
                    UpdateSearchFolderButtons();
                }
            }
        }

        private void LstSearchFolders_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateSearchFolderButtons();
        }

        private void UpdateSearchFolderButtons()
        {
            btnRemoveSearchFolder.Enabled = lstSearchFolders.SelectedIndex >= 0;
            btnClearSearchFolders.Enabled = searchFolders.Count > 0;
        }

        private void UpdateSearchFoldersHorizontalExtent()
        {
            if (lstSearchFolders == null)
                return;

            int maxWidth = lstSearchFolders.ClientSize.Width;
            foreach (var item in lstSearchFolders.Items)
            {
                string text = item?.ToString() ?? string.Empty;
                int textWidth = TextRenderer.MeasureText(text, lstSearchFolders.Font).Width;
                if (textWidth > maxWidth)
                {
                    maxWidth = textWidth;
                }
            }

            lstSearchFolders.HorizontalExtent = maxWidth + 20;
        }

        private void BtnSelectDestination_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Выберите папку для сбора файлов";
                fbd.ShowNewFolderButton = true;
                
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    destinationPath = fbd.SelectedPath;
                    txtDestinationPath.Text = destinationPath;
                    UpdateStartButtonState();
                }
            }
        }

        private void UpdateStartButtonState()
        {
            bool ready = !string.IsNullOrEmpty(excelPath) && 
                        searchFolders.Count > 0 && 
                        !string.IsNullOrEmpty(destinationPath);
            
            btnPreview.Enabled = ready && !isOperationRunning;
            btnStart.Enabled = ready && !isOperationRunning;
            
            btnSelectExcel.Enabled = !isOperationRunning;
            btnAddSearchFolder.Enabled = !isOperationRunning;
            btnSelectDestination.Enabled = !isOperationRunning;
            clbExtensions.Enabled = !isOperationRunning;
            txtCustomExtension.Enabled = !isOperationRunning;
        }

        private async void BtnPreview_Click(object sender, EventArgs e)
        {
            if (isOperationRunning) return;

            totalFiles = 0;
            copiedFiles = 0;
            notFoundFiles = 0;
            UpdateStats();

            try
            {
                SetOperationState(true);
                dgvPreview.Rows.Clear();
                previewItems.Clear();
                progressBar.Value = 0;
                UpdateProgressLabel("Подготовка предпросмотра...");

                await Task.Run(() => GeneratePreviewParallel());

                dgvPreview.Visible = true;
                txtLog.Visible = false;

                foreach (var item in previewItems)
                {
                    string fileDate = item.FileDate != DateTime.MinValue ? 
                        item.FileDate.ToString("dd.MM.yyyy HH:mm") : "Не найден";
                    
                    dgvPreview.Rows.Add(true, item.FileName, item.Extension, 
                        item.SourcePath, fileDate);
                }

                totalFiles = previewItems.Count;
                copiedFiles = 0;
                notFoundFiles = previewItems.Count(item => string.IsNullOrWhiteSpace(item.SourcePath));
                UpdateStats();

                UpdateProgressLabel($"Найдено {previewItems.Count} файлов. Снимите галочки с ненужных и нажмите 'Запустить копирование'");
            }
            catch (OperationCanceledException)
            {
                UpdateProgressLabel("Операция прервана пользователем");
            }
            catch (Exception ex)
            {
                LogMessage($"Ошибка при создании предпросмотра: {ex.Message}");
                UpdateProgressLabel("Ошибка при создании предпросмотра");
            }
            finally
            {
                SetOperationState(false);
            }
        }

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            if (isOperationRunning) return;

            totalFiles = 0;
            copiedFiles = 0;
            notFoundFiles = 0;
            skippedFiles = 0;
            UpdateStats();

            bool usePreviewData = dgvPreview.Visible && dgvPreview.Rows.Count > 0;

            if (usePreviewData)
            {
                int selectedCount = 0;
                foreach (DataGridViewRow row in dgvPreview.Rows)
                {
                    DataGridViewCheckBoxCell chk = row.Cells["Copy"] as DataGridViewCheckBoxCell;
                    if (chk != null && chk.Value != null && (bool)chk.Value)
                        selectedCount++;
                }

                if (selectedCount == 0)
                {
                    MessageBox.Show("Выберите хотя бы один файл для копирования", 
                        "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            try
            {
                SetOperationState(true);
                completedOperations.Clear();
                
                dgvPreview.Visible = false;
                txtLog.Visible = true;
                txtLog.Clear();
                progressBar.Value = 0;
                
                UpdateProgressLabel("Начало обработки...");
                UpdateCurrentFolderLabel("");
                UpdateStats();

                cancellationTokenSource = new CancellationTokenSource();
                var cancellationToken = cancellationTokenSource.Token;

                await Task.Run(() => ProcessFiles(usePreviewData, cancellationToken), cancellationToken);

                if (!cancellationToken.IsCancellationRequested)
                {
                    CreateNotFoundReport();
                    
                    progressBar.Value = 100;
                    UpdateProgressLabel("Обработка завершена");
                    UpdateCurrentFolderLabel("");
                    
                    MessageBox.Show($"Обработка завершена\n\nСкопировано: {copiedFiles}\nНе найдено: {notFoundFiles}", 
                        "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (OperationCanceledException)
            {
                UpdateProgressLabel("Операция прервана пользователем");
            }
            catch (Exception ex)
            {
                LogMessage($"Критическая ошибка: {ex.Message}");
                UpdateProgressLabel("Ошибка при выполнении операции");
            }
            finally
            {
                SetOperationState(false);
                cancellationTokenSource?.Dispose();
                cancellationTokenSource = null;
                Cleanup();
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            if (isOperationRunning && cancellationTokenSource != null)
            {
                if (MessageBox.Show("Прервать текущую операцию?", "Подтверждение", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    cancellationTokenSource.Cancel();
                    btnCancel.Enabled = false;
                    btnCancel.Text = "Отмена...";
                    UpdateProgressLabel("Завершение операции...");
                }
            }
        }

        private void SelectAllPreviewItems(bool select)
        {
            foreach (DataGridViewRow row in dgvPreview.Rows)
            {
                if (row.Cells["Copy"] is DataGridViewCheckBoxCell checkCell)
                {
                    checkCell.Value = select;
                }
            }
        }

        private void DgvPreview_CellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right || e.RowIndex < 0)
                return;

            if (!dgvPreview.Rows[e.RowIndex].Selected)
            {
                dgvPreview.ClearSelection();
                dgvPreview.Rows[e.RowIndex].Selected = true;
            }

            if (e.ColumnIndex >= 0)
            {
                dgvPreview.CurrentCell = dgvPreview.Rows[e.RowIndex].Cells[e.ColumnIndex];
            }
        }

        private string? GetSelectedPreviewFilePath()
        {
            if (dgvPreview.SelectedRows.Count == 0)
                return null;

            return dgvPreview.SelectedRows[0].Cells["SourcePath"].Value?.ToString();
        }

        private void OpenSelectedPreviewFile()
        {
            string? path = GetSelectedPreviewFilePath();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return;

            try
            {
                var startInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = path,
                    UseShellExecute = true
                };
                System.Diagnostics.Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                LogMessage($"Ошибка открытия файла {path}: {ex.Message}");
            }
        }

        private void OpenSelectedPreviewFolder()
        {
            string? path = GetSelectedPreviewFilePath();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return;

            string? folder = Path.GetDirectoryName(path);
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                return;

            try
            {
                System.Diagnostics.Process.Start("explorer.exe", $"\"{folder}\"");
            }
            catch (Exception ex)
            {
                LogMessage($"Ошибка открытия папки {folder}: {ex.Message}");
            }
        }

        private void UncheckSelectedPreviewItems()
        {
            if (dgvPreview.SelectedRows.Count == 0)
                return;

            foreach (DataGridViewRow row in dgvPreview.SelectedRows)
            {
                if (row.Cells["Copy"] is DataGridViewCheckBoxCell checkCell)
                {
                    checkCell.Value = false;
                }
            }
        }

        private void CopySelectedPreviewNamesToClipboard()
        {
            var selectedNames = dgvPreview.SelectedRows
                .Cast<DataGridViewRow>()
                .Select(row =>
                {
                    string fileName = row.Cells["FileName"].Value?.ToString() ?? string.Empty;
                    string extension = row.Cells["Extension"].Value?.ToString() ?? string.Empty;
                    return $"{fileName}{extension}";
                })
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (selectedNames.Count == 0)
                return;

            try
            {
                Clipboard.SetText(string.Join(Environment.NewLine, selectedNames));
            }
            catch (Exception ex)
            {
                LogMessage($"Ошибка копирования в буфер обмена: {ex.Message}");
            }
        }


        #endregion

        #region Core Logic

        private void SetOperationState(bool running)
        {
            isOperationRunning = running;
            
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => SetOperationState(running)));
                return;
            }

            btnCancel.Visible = running;
            btnCancel.Enabled = running;
            btnCancel.Text = "Отмена";
            
            UpdateStartButtonState();
        }

        private void GeneratePreviewParallel()
        {
            List<string> extensions = GetSelectedExtensions();
            List<string> fileNames = ReadExcelFileNames();
            
            previewItems.Clear();
            var parallelPreviewItems = new ConcurrentBag<PreviewItem>();
            int totalCombinations = fileNames.Count * extensions.Count;
            int processed = 0;

            Parallel.ForEach(fileNames, new ParallelOptions 
            { 
                MaxDegreeOfParallelism = Environment.ProcessorCount 
            }, fileName =>
            {
                foreach (string extension in extensions)
                {
                    int currentProcessed = Interlocked.Increment(ref processed);
                    int progress = (int)(currentProcessed * 100.0 / totalCombinations);
                    
                    UpdateProgressBar(progress);
                    UpdateProgressLabel($"Поиск {currentProcessed} из {totalCombinations}: {fileName}{extension}");

                    string foundFilePath = FindLatestFile(fileName, extension);

                    if (!string.IsNullOrEmpty(foundFilePath))
                    {
                        FileInfo fi = new FileInfo(foundFilePath);
                        parallelPreviewItems.Add(new PreviewItem
                        {
                            FileName = fileName,
                            Extension = extension,
                            SourcePath = foundFilePath,
                            FileDate = fi.LastWriteTime
                        });
                    }
                    else
                    {
                        parallelPreviewItems.Add(new PreviewItem
                        {
                            FileName = fileName,
                            Extension = extension,
                            SourcePath = string.Empty,
                            FileDate = DateTime.MinValue
                        });
                    }
                }
            });

            previewItems = parallelPreviewItems.OrderBy(p => p.FileName)
                                              .ThenBy(p => p.Extension)
                                              .ToList();
        }

        private void ProcessFiles(bool usePreviewData, CancellationToken cancellationToken)
        {
            if (usePreviewData)
            {
                ProcessFromPreview(cancellationToken);
            }
            else
            {
                List<string> extensions = GetSelectedExtensions();
                List<string> fileNames = ReadExcelFileNames();
                ProcessFromExcel(fileNames, extensions, cancellationToken);
            }
        }

        private void ProcessFromPreview(CancellationToken cancellationToken)
        {
            List<PreviewItem> selectedItems = new List<PreviewItem>();
            
            this.Invoke(new Action(() =>
            {
                foreach (DataGridViewRow row in dgvPreview.Rows)
                {
                    DataGridViewCheckBoxCell chk = row.Cells["Copy"] as DataGridViewCheckBoxCell;
                    if (chk != null && chk.Value != null && (bool)chk.Value)
                    {
                        selectedItems.Add(new PreviewItem
                        {
                            FileName = row.Cells["FileName"].Value?.ToString() ?? "",
                            Extension = row.Cells["Extension"].Value?.ToString() ?? "",
                            SourcePath = row.Cells["SourcePath"].Value?.ToString() ?? ""
                        });
                    }
                }
            }));

            totalFiles = selectedItems.Count;
            UpdateStats();

            for (int i = 0; i < selectedItems.Count; i++)
            {
                if (cancellationToken.IsCancellationRequested)
                    return;

                var item = selectedItems[i];
                int progress = (int)((i + 1) * 100.0 / selectedItems.Count);
                
                UpdateProgressBar(progress);
                UpdateProgressLabel($"Копирование {i + 1} из {selectedItems.Count}: {item.FileName}{item.Extension}");
                
                if (string.IsNullOrEmpty(item.SourcePath))
                {
                    item.SourcePath = FindLatestFile(item.FileName, item.Extension);
                }

                if (!string.IsNullOrEmpty(item.SourcePath))
                {
                    LogMessage($"Копирование: {item.FileName}{item.Extension}");
                    if (CopyFileWithRetry(item.FileName, item.Extension, item.SourcePath, cancellationToken))
                        copiedFiles++;
                    else
                        notFoundFiles++;
                }
                else
                {
                    LogMessage($"Не найден: {item.FileName}{item.Extension}");
                    completedOperations.Add(new CopyOperation 
                    { 
                        FileName = item.FileName, 
                        Extension = item.Extension, 
                        SourcePath = string.Empty,
                        Status = "Не найден" 
                    });
                    notFoundFiles++;
                }
                    
                UpdateStats();
            }
        }

        private void ProcessFromExcel(List<string> fileNames, List<string> extensions, CancellationToken cancellationToken)
        {
            totalFiles = fileNames.Count * extensions.Count;
            UpdateStats();
            UpdateProgressBar(0);

            int processed = 0;

            foreach (string fileName in fileNames)
            {
                if (cancellationToken.IsCancellationRequested)
                    return;

                foreach (string extension in extensions)
                {
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    processed++;
                    int progress = (int)(processed * 100.0 / totalFiles);
                    
                    UpdateProgressBar(progress);
                    UpdateProgressLabel($"Обработка {processed} из {totalFiles}: {fileName}{extension}");
                    LogMessage($"Поиск файла: {fileName}{extension}");
                    
                    string foundFilePath = FindLatestFile(fileName, extension);

                    if (!string.IsNullOrEmpty(foundFilePath))
                    {
                        if (CopyFileWithRetry(fileName, extension, foundFilePath, cancellationToken))
                            copiedFiles++;
                        else
                            notFoundFiles++;
                    }
                    else
                    {
                        LogMessage($"Не найден: {fileName}{extension}");
                        completedOperations.Add(new CopyOperation 
                        { 
                            FileName = fileName, 
                            Extension = extension, 
                            SourcePath = string.Empty,
                            Status = "Не найден" 
                        });
                        notFoundFiles++;
                    }

                    UpdateStats();
                }
            }
        }

        private bool CopyFileWithRetry(string fileName, string extension, string sourcePath, CancellationToken cancellationToken)
        {
            try
            {
                if (cancellationToken.IsCancellationRequested)
                    return false;

                if (!File.Exists(sourcePath))
                {
                    LogMessage($"Файл не существует: {sourcePath}");
                    completedOperations.Add(new CopyOperation 
                    { 
                        FileName = fileName, 
                        Extension = extension, 
                        SourcePath = sourcePath,
                        Status = "Не существует" 
                    });
                    notFoundFiles++;
                    return false;
                }

                string destFile = Path.Combine(destinationPath, Path.GetFileName(sourcePath));
                
                if (Path.GetFullPath(sourcePath).Equals(Path.GetFullPath(destFile), StringComparison.OrdinalIgnoreCase))
                {
                    LogMessage($"Пропущен (уже в папке назначения): {sourcePath}");
                    completedOperations.Add(new CopyOperation 
                    { 
                        FileName = fileName, 
                        Extension = extension, 
                        SourcePath = sourcePath, 
                        DestinationPath = destFile,
                        Status = "Уже в папке назначения" 
                    });
                    skippedFiles++;
                    return true;
                }

                int attempts = 3;
                
                for (int attempt = 1; attempt <= attempts; attempt++)
                {
                    if (cancellationToken.IsCancellationRequested)
                        return false;

                    try
                    {
                        File.Copy(sourcePath, destFile, true);
                        LogMessage($"Скопирован: {sourcePath}");
                        completedOperations.Add(new CopyOperation 
                        { 
                            FileName = fileName, 
                            Extension = extension, 
                            SourcePath = sourcePath, 
                            DestinationPath = destFile,
                            Status = "Успешно",
                            FileSize = new FileInfo(sourcePath).Length
                        });
                        return true;
                    }
                    catch (IOException ioEx) when (attempt < attempts)
                    {
                        if (skipLockedFiles && IsFileLocked(ioEx))
                        {
                            LogMessage($"Пропущен (заблокирован): {sourcePath}");
                            completedOperations.Add(new CopyOperation 
                            { 
                                FileName = fileName, 
                                Extension = extension, 
                                SourcePath = sourcePath, 
                                Status = "Пропущен (заблокирован)" 
                            });
                            skippedFiles++;
                            return false;
                        }
                        
                        LogMessage($"Попытка {attempt}: Файл занят, повтор через 1 сек...");
                        Thread.Sleep(1000);
                    }
                }
                
                LogMessage($"Не удалось скопировать (заблокирован): {sourcePath}");
                completedOperations.Add(new CopyOperation 
                { 
                    FileName = fileName, 
                    Extension = extension, 
                    SourcePath = sourcePath, 
                    Status = "Не удалось скопировать" 
                });
                notFoundFiles++;
                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Ошибка копирования {sourcePath}: {ex.Message}");
                completedOperations.Add(new CopyOperation 
                { 
                    FileName = fileName, 
                    Extension = extension, 
                    SourcePath = sourcePath, 
                    Status = $"Ошибка: {ex.Message}" 
                });
                
                if (continueOnError)
                {
                    notFoundFiles++;
                    return false;
                }
                else
                {
                    throw;
                }
            }
        }

        private bool IsFileLocked(IOException ex)
        {
            int errorCode = System.Runtime.InteropServices.Marshal.GetHRForException(ex) & 0xFFFF;
            return errorCode == 32 || errorCode == 33; // ERROR_SHARING_VIOLATION or ERROR_LOCK_VIOLATION
        }

        private void CreateNotFoundReport()
        {
            if (string.IsNullOrEmpty(destinationPath) || !Directory.Exists(destinationPath))
                return;

            var notFoundOperations = completedOperations
                .Where(op => op.Status == "Не найден" || op.Status == "Не существует")
                .ToList();
            
            if (notFoundOperations.Count == 0)
                return;

            string dateTimeStr = DateTime.Now.ToString("dd.MM.yy HH.mm");
            string reportPath = Path.Combine(destinationPath, $"Отчет не найденных файлов {dateTimeStr}.txt");

            try
            {
                using (StreamWriter sw = new StreamWriter(reportPath, false, Encoding.UTF8))
                {
                    sw.WriteLine($"Отчет о ненайденных файлах");
                    sw.WriteLine($"Дата создания: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
                    sw.WriteLine($"Всего не найдено: {notFoundOperations.Count}");
                    sw.WriteLine(new string('=', 60));
                    sw.WriteLine();

                    var groupedByExtension = notFoundOperations
                        .GroupBy(op => op.Extension)
                        .OrderBy(g => g.Key);

                    foreach (var group in groupedByExtension)
                    {
                        sw.WriteLine($"Формат: {group.Key} (всего: {group.Count()})");
                        sw.WriteLine(new string('-', 40));
                        
                        foreach (var op in group.OrderBy(op => op.FileName))
                        {
                            sw.WriteLine($"  {op.FileName}");
                        }
                        
                        sw.WriteLine();
                    }
                }

                LogMessage($"Создан отчет о ненайденных файлов: {reportPath}");
            }
            catch (Exception ex)
            {
                LogMessage($"Ошибка при создании отчета: {ex.Message}");
            }
        }

        #endregion

        #region Helper Methods

        private List<string> GetSelectedExtensions()
        {
            List<string> extensions = new List<string>();
            
            foreach (var item in clbExtensions.CheckedItems)
            {
                extensions.Add(item.ToString());
            }

            if (!string.IsNullOrWhiteSpace(txtCustomExtension.Text))
            {
                string[] customExts = txtCustomExtension.Text.Split(
                    new[] { ',', ';', ' ', '\t' }, 
                    StringSplitOptions.RemoveEmptyEntries);
                
                foreach (string ext in customExts)
                {
                    string cleaned = ext.Trim();
                    if (!cleaned.StartsWith("."))
                        cleaned = "." + cleaned;
                    
                    if (!extensions.Contains(cleaned, StringComparer.OrdinalIgnoreCase))
                        extensions.Add(cleaned);
                }
            }

            if (extensions.Count == 0)
                extensions.Add(".pdf");

            return extensions.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private List<string> ReadExcelFileNames()
        {
            List<string> fileNames = new List<string>();

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelPath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    
                    if (worksheet.Dimension == null)
                    {
                        LogMessage("Excel файл пустой или не содержит данных");
                        return fileNames;
                    }

                    int rowCount = worksheet.Dimension.End.Row;
                    
                    LogMessage($"Размер таблицы: {rowCount} строк");

                    // Читаем ТОЛЬКО столбец B (индекс 2)
                    int fileNameColumn = 2;
                    
                    // Определяем строку начала данных (игнорируем заголовки)
                    int startRow = FindDataStartRow(worksheet, fileNameColumn, rowCount);
                    LogMessage($"Чтение данных из столбца B, начиная со строки {startRow}");

                    for (int row = startRow; row <= rowCount; row++)
                    {
                        var cell = worksheet.Cells[row, fileNameColumn];
                        string cellValue = GetCellValueAsString(cell);

                        LogMessage($"ОТЛАДКА Excel: Строка {row}, сырое значение: [{cellValue}]");

                        if (!string.IsNullOrWhiteSpace(cellValue))
                        {
                            string fileName = cellValue.Trim();

                            // Проверяем, не является ли это стандартным заголовком или исключением
                            if (!IsHeaderOrException(fileName))
                            {
                                // Получаем имя файла без расширения
                                string nameWithoutExt = GetFileNameWithoutExtension(fileName);

                                LogMessage($"ОТЛАДКА Excel: После обработки '{fileName}' -> '{nameWithoutExt}'");

                                // Если после удаления расширения имя не пустое, используем его
                                // Иначе используем оригинальное имя
                                if (!string.IsNullOrWhiteSpace(nameWithoutExt))
                                {
                                    fileNames.Add(nameWithoutExt);
                                    LogMessage($"Добавлено имя файла: '{nameWithoutExt}' (строка {row}: '{cellValue}')");
                                }
                                else
                                {
                                    fileNames.Add(fileName);
                                    LogMessage($"Добавлено имя файла: '{fileName}' (строка {row})");
                                }
                            }
                            else
                            {
                                LogMessage($"Пропущена ячейка B{row} (заголовок или исключение): '{cellValue}'");
                            }
                        }
                        else
                        {
                            LogMessage($"Пустая ячейка B{row}, пропуск");
                        }
                    }
                }

                LogMessage($"Загружено {fileNames.Count} имен файлов из столбца B");
                
                // Логируем все имена для проверки
                for (int i = 0; i < fileNames.Count; i++)
                {
                    LogMessage($"Файл #{i + 1}: {fileNames[i]}");
                }
                
                if (fileNames.Count == 0)
                {
                    LogMessage("ВНИМАНИЕ: Не найдено ни одного имени файла в столбце B!");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Ошибка чтения Excel: {ex.Message}");
                LogMessage($"Детали ошибки: {ex}");
            }

            return fileNames.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        // Вспомогательные методы

        private int FindDataStartRow(ExcelWorksheet worksheet, int fileNameColumn, int rowCount)
        {
            // Ищем строку, с которой начинаются данные (игнорируем заголовки в столбце B)
            // Проверяем первые 5 строк на наличие заголовков
            int maxRowsToCheck = Math.Min(5, rowCount);
            
            for (int row = 1; row <= maxRowsToCheck; row++)
            {
                string cellValue = GetCellValueAsString(worksheet.Cells[row, fileNameColumn]);
                if (!string.IsNullOrWhiteSpace(cellValue))
                {
                    string trimmed = cellValue.Trim();
                    
                    // Если это не заголовок, начинаем с этой строки
                    if (!IsHeaderOrException(trimmed))
                    {
                        LogMessage($"Первая строка с данными: B{row} = '{trimmed}'");
                        return row;
                    }
                    else
                    {
                        LogMessage($"Строка B{row} содержит заголовок: '{trimmed}', пропускаем");
                    }
                }
            }
            
            // Если не нашли заголовков, начинаем с первой строки
            LogMessage("Заголовки не найдены, начинаем чтение с первой строки");
            return 1;
        }

        private bool IsHeaderOrException(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;
            
            string trimmed = value.Trim();
            
            // Расширенный список стандартных заголовков для столбца B
            HashSet<string> headers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                // Русские заголовки
                "Имя файла", "File Name", "Наименование файла", "Название файла", "Имя", "Наименование", 
                "Название", "Файл", "Файлы", "Документ", "Документы", "Чертеж", "Чертежи", "Модель", "Модели",
                "Схема", "Схемы", "Проект", "Проекты", "Архив", "Архивы", "Папка", "Папки",
                
                // Английские заголовки
                "Filename", "File", "Name", "Document", "Drawing", "Model", "Scheme", "Project", "Folder",
                "File name", "File names", "Document name", "Drawing name",
                
                // Заголовки таблиц (могут быть в столбце B)
                "№ п/п", "№", "Порядковый номер", "Номер", "Позиция", "Код", "Артикул", "Артикулы",
                "Количество", "Кол-во", "Кол.", "Примечание", "Примечания", "Комментарий", "Комментарии",
                "Дата", "Date", "Время", "Time", "Создан", "Создано", "Изменен", "Изменено",
                "Автор", "Author", "Версия", "Version", "Ревизия", "Revision", "Rev.", "Вер.",
                
                // Итоговые строки
                "Итого", "Всего", "Total", "Сумма", "Sum", "Итог", "Итоги", "Total sum",
                "Всего строк", "Всего записей", "Total records", "Итого по странице",
                "Конец", "End", "Конец списка", "End of list", "Продолжение", "Continuation",
                "Далее", "Next", "Следующая страница", "Next page",
                
                // Пустые/служебные значения
                "н/д", "n/a", "N/A", "не указано", "не задано", "не применимо",
                "пусто", "empty", "null", "undefined", "отсутствует", "missing",
                
                // Специальные символы
                "-", "--", "---", "...", "…", "***", "///"
            };

            // Проверяем полное совпадение
            if (headers.Contains(trimmed))
                return true;
            
            // Проверяем, начинается ли с заголовка (например, "Имя файла:" или "File Name:")
            if (headers.Any(h => 
                trimmed.StartsWith(h, StringComparison.OrdinalIgnoreCase) && 
                (trimmed.Length == h.Length || 
                 trimmed[h.Length] == ':' || 
                 trimmed[h.Length] == ' ' ||
                 trimmed[h.Length] == '-' ||
                 trimmed[h.Length] == '(')))
            {
                return true;
            }
            
            // Проверяем на наличие скобок с номером (например, "Итого (10 шт.)")
            if (trimmed.StartsWith("Итого", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("Всего", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("Total", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("Sum", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            
            // Проверяем, не является ли это числом (номера строк)
            if (int.TryParse(trimmed, out _))
                return true;
            
            // Проверяем, не является ли это датой
            if (DateTime.TryParse(trimmed, out _))
                return true;
            
            // Проверяем на наличие служебных слов (только в начале!)
            string lowerTrimmed = trimmed.ToLower();
            if (lowerTrimmed.StartsWith("страница") ||
                lowerTrimmed.StartsWith("page") ||
                lowerTrimmed.StartsWith("таблица") ||
                lowerTrimmed.StartsWith("table") ||
                lowerTrimmed.StartsWith("отчет") ||
                lowerTrimmed.StartsWith("report"))
            {
                return true;
            }

            return false;
        }

        private string GetCellValueAsString(ExcelRange cell)
        {
            if (cell == null || cell.Value == null)
                return string.Empty;
            
            try
            {
                // Используем cell.Text для получения отображаемого текста
                // cell.Text возвращает текст, как он отображается в Excel
                string cellText = cell.Text;
                
                // Если cell.Text пустой (может быть для некоторых типов данных),
                // используем cell.Value.ToString()
                if (string.IsNullOrEmpty(cellText))
                {
                    cellText = cell.Value?.ToString() ?? string.Empty;
                }
                
                // Если все еще пусто, возвращаем пустую строку
                if (string.IsNullOrEmpty(cellText))
                    return string.Empty;

                // Очищаем текст от неразрывных пробелов и других специальных символов
                cellText = cellText.Replace("\u00A0", " ") // Неразрывный пробел
                                   .Replace("\u202F", " ") // Узкий неразрывный пробел
                                   .Replace("\u2000", " ") // En quad
                                   .Replace("\u2001", " ") // Em quad
                                   .Replace("\u2002", " ") // En space
                                   .Replace("\u2003", " ") // Em space
                                   .Replace("\u2004", " ") // Three-per-em space
                                   .Replace("\u2005", " ") // Four-per-em space
                                   .Replace("\u2006", " ") // Six-per-em space
                                   .Replace("\u2007", " ") // Figure space
                                   .Replace("\u2008", " ") // Punctuation space
                                   .Replace("\u2009", " ") // Thin space
                                   .Replace("\u200A", " ") // Hair space
                                   .Replace("\r\n", " ")   // Переносы строк
                                   .Replace("\n", " ")
                                   .Replace("\r", " ")
                                   .Replace("\t", " ");    // Табуляция

                // Удаляем различные типы кавычек
                cellText = cellText.Replace("\"", "")     // Обычная двойная кавычка
                                   .Replace("'", "")      // Обычный апостроф
                                   .Replace("\u201C", "") // Левая двойная кавычка
                                   .Replace("\u201D", "") // Правая двойная кавычка
                                   .Replace("«", "")      // Левая елочка
                                   .Replace("»", "")      // Правая елочка
                                   .Replace("\u2039", "") // Одиночная левая кавычка
                                   .Replace("\u203A", "") // Одиночная правая кавычка
                                   .Replace("\u2018", "") // Левый апостроф
                                   .Replace("\u2019", "") // Правый апостроф
                                   .Replace("`", "")      // Backtick
                                   .Replace("\u00B4", "");// Acute accent

                // Заменяем различные типы дефисов и тире на обычный дефис
                cellText = cellText.Replace("–", "-")     // En dash
                                   .Replace("—", "-")     // Em dash
                                   .Replace("‐", "-")     // Hyphen
                                   .Replace("‑", "-")     // Non-breaking hyphen
                                   .Replace("⁃", "-")     // Hypen bullet
                                   .Replace("−", "-")     // Minus sign
                                   .Replace("─", "-");    // Box drawings

                // Удаляем невидимые символы (zero-width space, -joiner, -non-joiner и т.д.)
                cellText = cellText.Replace("\u200B", "")  // Zero Width Space
                                   .Replace("\u200C", "")  // Zero Width Non-Joiner
                                   .Replace("\u200D", "")  // Zero Width Joiner
                                   .Replace("\u206B", "")  // Activate Symmetric Swapping
                                   .Replace("\uFEFF", ""); // Zero Width No-Break Space (BOM)

                // Удаляем лишние пробелы
                while (cellText.Contains("  "))
                    cellText = cellText.Replace("  ", " ");

                return cellText.Trim();
            }
            catch
            {
                // В случае ошибки возвращаем пустую строку
                return string.Empty;
            }
        }

        private string GetFileNameWithoutExtension(string fileName)
        {
            try
            {
                // Список известных расширений
                HashSet<string> knownExtensions = new HashSet<string>
                {
                    ".pdf", ".dxf", ".dwg", ".doc", ".docx", ".xlsx", ".xls", ".txt",
                    ".jpg", ".png", ".sldprt", ".sldasm", ".slddrw", ".step", ".stp",
                    ".iges", ".igs", ".ipt", ".iam", ".idw", ".prt", ".asm", ".drw",
                    ".catpart", ".catproduct", ".catdrawing", ".par", ".psm", ".dft",
                    ".3dm", ".skp", ".dgn", ".rvt", ".rfa", ".rte", ".ifc", ".sat",
                    ".x_t", ".x_b", ".jt", ".u3d", ".dae", ".fbx", ".obj", ".stl"
                };

                // Проверяем все известные расширения - ищем совпадение в конце имена (нечувствительно к регистру)
                foreach (var ext in knownExtensions)
                {
                    if (fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                    {
                        // Нашли расширение в конце - удаляем его
                        return fileName.Substring(0, fileName.Length - ext.Length).Trim();
                    }
                }

                // Если не нашли известное расширение, возвращаем оригинальное имя
                return fileName;
            }
            catch
            {
                // В случае ошибки возвращаем оригинальное имя
                return fileName;
            }
        }

        private string FindLatestFile(string baseName, string extension)
        {
            string targetFileName = baseName + extension;
            FileInfo latestFile = null;
            bool foundFile = false;

            LogMessage($"ОТЛАДКА FindLatestFile: Ищу файл '{baseName}' + расширение '{extension}' = '{targetFileName}'");

            foreach (string searchFolder in searchFolders)
            {
                try
                {
                    SearchInDirectory(searchFolder, baseName, extension, ref latestFile);
                    if (latestFile != null)
                        foundFile = true;
                }
                catch (Exception ex)
                {
                    LogMessage($"Ошибка доступа к папке {searchFolder}: {ex.Message}");
                }
            }

            // Логируем результат поиска для отладки
            if (!foundFile)
            {
                LogMessage($"ОТЛАДКА: Файл '{targetFileName}' не найден в следующих папках:");
                foreach (string folder in searchFolders)
                {
                    LogMessage($"  - {folder}");
                }
            }
            else
            {
                LogMessage($"ОТЛАДКА: Найден файл: {latestFile?.FullName}");
            }

            return latestFile?.FullName;
        }

        private List<string> EnumerateFilesCaseInsensitive(string directory, string targetFileName)
        {
            try
            {
                var options = new EnumerationOptions
                {
                    RecurseSubdirectories = false,
                    IgnoreInaccessible = true,
                    MatchCasing = MatchCasing.CaseInsensitive,
                    ReturnSpecialDirectories = false
                };

                return Directory.EnumerateFiles(directory, targetFileName, options).ToList();
            }
            catch (PlatformNotSupportedException)
            {
                // Fallback for runtimes that do not support explicit match casing.
            }
            catch (NotSupportedException)
            {
                // Fallback for runtimes that do not support explicit match casing.
            }
            catch (ArgumentException)
            {
                // Fallback if the target mask cannot be used directly.
            }

            return Directory
                .EnumerateFiles(directory, "*", SearchOption.TopDirectoryOnly)
                .Where(path => string.Equals(Path.GetFileName(path), targetFileName, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        private void SearchInDirectory(string directory, string baseName, string extension, ref FileInfo latestFile)
        {
            FileInfo sharedLatest = latestFile;
            string targetFileName = baseName + extension;

            try
            {
                UpdateCurrentFolderLabel($"Сканирование: {directory}");

                if (skipSystemHidden)
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(directory);
                    if ((dirInfo.Attributes & (FileAttributes.System | FileAttributes.Hidden)) != 0)
                        return;
                }

                // Точный поиск по имени файла
                var files = EnumerateFilesCaseInsensitive(directory, targetFileName);

                foreach (string filePath in files)
                {
                    try
                    {
                        FileInfo fileInfo = new FileInfo(filePath);

                        if (skipSystemHidden && 
                            (fileInfo.Attributes & (FileAttributes.System | FileAttributes.Hidden)) != 0)
                            continue;

                        lock (latestFileLock)
                        {
                            if (sharedLatest == null || fileInfo.LastWriteTime > sharedLatest.LastWriteTime)
                            {
                                sharedLatest = fileInfo;
                            }
                        }
                    }
                    catch (UnauthorizedAccessException)
                    {
                        LogMessage($"Нет доступа к файлу: {filePath}");
                    }
                }

                // Рекурсивный поиск в подпапках
                var subdirectories = Directory.EnumerateDirectories(directory).ToList();
                Parallel.ForEach(subdirectories, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, subdirectory =>
                {
                    try
                    {
                        FileInfo localLatest = null;
                        SearchInDirectory(subdirectory, baseName, extension, ref localLatest);

                        if (localLatest != null)
                        {
                            lock (latestFileLock)
                            {
                                if (sharedLatest == null || localLatest.LastWriteTime > sharedLatest.LastWriteTime)
                                {
                                    sharedLatest = localLatest;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"Ошибка при сканировании подпапки {subdirectory}: {ex.Message}");
                    }
                });
            }
            catch (UnauthorizedAccessException)
            {
                LogMessage($"Нет доступа к папке: {directory}");
            }
            catch (Exception ex)
            {
                LogMessage($"Ошибка сканирования папки {directory}: {ex.Message}");
            }

            latestFile = sharedLatest;
        }

        private void Cleanup()
        {
            previewItems.Clear();
            completedOperations.Clear();
            
            // Очистка старых записей кэша
            var oldKeys = fileSystemCache
                .Where(kvp => (DateTime.Now - kvp.Value.LastScanTime).TotalMinutes > 30)
                .Select(kvp => kvp.Key)
                .ToList();
            
            foreach (var key in oldKeys)
            {
                fileSystemCache.TryRemove(key, out _);
            }
            
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        #endregion

        #region Helper Methods

        private void LogMessage(string message)
        {
            if (txtLog.InvokeRequired)
            {
                txtLog.Invoke(new Action(() => LogMessage(message)));
            }
            else
            {
                txtLog.AppendText($"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}");
                txtLog.ScrollToCaret();
            }
        }

        private void UpdateStats()
        {
            if (lblStats.InvokeRequired)
            {
                lblStats.Invoke(new Action(() => UpdateStats()));
            }
            else
            {
                lblStats.Text = $"Статистика: Всего файлов: {totalFiles} | " +
                               $"Скопировано: {copiedFiles} | " +
                               $"Не найдено: {notFoundFiles}";
            }
        }

        private void UpdateProgressBar(int value)
        {
            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action(() => UpdateProgressBar(value)));
            }
            else
            {
                progressBar.Value = Math.Max(0, Math.Min(value, 100));
            }
        }

        private void UpdateProgressLabel(string text)
        {
            if (lblProgress.InvokeRequired)
            {
                lblProgress.Invoke(new Action(() => UpdateProgressLabel(text)));
            }
            else
            {
                lblProgress.Text = text;
            }
        }

        private void UpdateCurrentFolderLabel(string text)
        {
            if (lblCurrentFolder.InvokeRequired)
            {
                lblCurrentFolder.Invoke(new Action(() => UpdateCurrentFolderLabel(text)));
            }
            else
            {
                lblCurrentFolder.Text = text;
            }
        }

        #endregion
    }

    public class PreviewItem
    {
        public string FileName { get; set; }
        public string Extension { get; set; }
        public string SourcePath { get; set; }
        public DateTime FileDate { get; set; }
    }

    public class CopyOperation
    {
        public string FileName { get; set; }
        public string Extension { get; set; }
        public string SourcePath { get; set; }
        public string DestinationPath { get; set; }
        public string Status { get; set; }
        public long FileSize { get; set; }
    }

    public class FileSystemCacheEntry
    {
        public DateTime LastScanTime { get; set; }
        public List<string> Files { get; set; }
        public List<string> Directories { get; set; }
    }


}
