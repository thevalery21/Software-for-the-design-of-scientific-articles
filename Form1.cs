using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using PdfiumViewer;

namespace ПО
{
    public partial class Form1 : Form
    {
        // Словари для кнопок
        private Dictionary<Button, Image> originalImages = new Dictionary<Button, Image>();
        private Dictionary<Button, (int topLeft, int topRight, int bottomRight, int bottomLeft)> buttonCorners =
            new Dictionary<Button, (int, int, int, int)>();
        private Dictionary<Button, bool> transparencyFixed = new Dictionary<Button, bool>();
        private List<Image> pdfPages = new List<Image>();

        private string currentPdfPath;
        private Panel pagesPanel;
        // Элементы управления
        private ToolTip toolTip;
        private Button activeToggleButton = null;
        private double currentScale = 1.0;

        // Масштабирование
        private int currentPageIndex = 0;
        private int totalPages = 0;
        private const double minScale = 0.1;
        private const double maxScale = 5.0;

        // PdfViewer вместо WebBrowser
        private PdfDocument pdfDocument;
        private Timer progressTimer;
        private int progressValue = 0;
        private Panel progressNotificationPanel;
        private ProgressBar notificationProgressBar;
        private Label progressLabel;
        private PictureBox pdfPictureBox;
        private Panel pdfContainerPanel;
        

        private void InitializeProgressTimer()
        {
            progressTimer = new Timer();
            progressTimer.Interval = 50;
            progressTimer.Tick += ProgressTimer_Tick;
        }

        private void ProgressTimer_Tick(object sender, EventArgs e)
        {
            if (progressValue < 100)
            {
                Random rand = new Random();
                int increment = rand.Next(1, 2);
                progressValue += increment;
                if (progressValue > 100) progressValue = 100;

                if (progressValue < 90)
                {
                    UpdateProgressBarInStatus(progressValue);
                }
            }
            else
            {
                progressTimer.Stop();
            }
        }

        private void UpdateProgressBarInStatus(int progress)
        {
            try
            {
                if (notificationProgressBar != null && notificationProgressBar.InvokeRequired)
                {
                    notificationProgressBar.Invoke(new Action<int>(UpdateProgressBarInStatus), progress);
                    return;
                }

                int displayProgress = Math.Min(progress, 99);

                if (notificationProgressBar != null)
                {
                    notificationProgressBar.Value = displayProgress;
                    notificationProgressBar.Refresh();
                }
                if (progressLabel != null)
                {
                    progressLabel.Text = $"Загрузка файла... {displayProgress}%";
                    progressLabel.Refresh();
                }

                if (progress > 0 && progressNotificationPanel != null && !progressNotificationPanel.Visible)
                {
                    progressNotificationPanel.Visible = true;
                    progressNotificationPanel.BringToFront();
                    ForceUIUpdate();
                }

                UpdateProgressInHistory(displayProgress);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обновления прогресса: {ex.Message}");
            }
        }

        private void CompleteProgressFinal()
        {
            try
            {
                if (notificationProgressBar != null && notificationProgressBar.InvokeRequired)
                {
                    notificationProgressBar.Invoke(new Action(CompleteProgressFinal));
                    return;
                }

                if (progressTimer != null && progressTimer.Enabled)
                    progressTimer.Stop();

                if (notificationProgressBar != null)
                {
                    notificationProgressBar.Value = 100;
                }
                if (progressLabel != null)
                {
                    progressLabel.Text = $"Загрузка файла... 100%";
                }

                UpdateProgressInHistory(100);

                Timer hideTimer = new Timer();
                hideTimer.Interval = 800;
                hideTimer.Tick += (s, args) =>
                {
                    if (progressNotificationPanel != null)
                    {
                        progressNotificationPanel.Visible = false;
                    }

                    // ✅ ГАРАНТИРУЕМ ЧТО СТАТУС БАР ВСЕГДА ВИДИМ
                    statusStrip1.Visible = true;
                    statusStrip1.BringToFront();

                    // ОБНОВЛЯЕМ ИНФОРМАЦИЮ В СТАТУС БАРЕ
                    UpdateStatusBarInfo();

                    hideTimer.Stop();
                    hideTimer.Dispose();
                };
                hideTimer.Start();

                progressValue = 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка в CompleteProgressFinal: {ex.Message}");
            }
        }

        private void CompleteProgress()
        {
            try
            {
                if (progressTimer != null && progressTimer.Enabled)
                {
                    progressTimer.Stop();
                }

                progressValue = 0;

                if (progressNotificationPanel != null)
                {
                    progressNotificationPanel.Visible = false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка в CompleteProgress: {ex.Message}");
            }
        }

        private void DrawBlackBorderWithClear(object sender, PaintEventArgs e)
        {
            System.Windows.Forms.Control control = (System.Windows.Forms.Control)sender;

            e.Graphics.Clear(control.BackColor);
            ControlPaint.DrawBorder(e.Graphics, control.ClientRectangle,
                System.Drawing.Color.Gray, 1, ButtonBorderStyle.Solid,
                System.Drawing.Color.Gray, 1, ButtonBorderStyle.Solid,
                System.Drawing.Color.Gray, 1, ButtonBorderStyle.Solid,
                System.Drawing.Color.Gray, 1, ButtonBorderStyle.Solid);
        }

        private void DrawBlackBorder(object sender, PaintEventArgs e)
        {
            ToolStripStatusLabel label = (ToolStripStatusLabel)sender;

            using (Pen blackPen = new Pen(System.Drawing.Color.Gray, 1))
            {
                Rectangle borderRect = new Rectangle(0, 0, label.Width - 1, label.Height - 1);
                e.Graphics.DrawRectangle(blackPen, borderRect);
            }
        }

        public Form1()
        {
            InitializeComponent();
            InitializeToolTips();
            InitializeButtonHoverEffects();

            UpdateButtonsState(false);

            toolStripStatusLabel1.Paint += DrawBlackBorder;
            toolStripStatusLabel2.Paint += DrawBlackBorder;
            toolStrip1.Paint += DrawBlackBorderWithClear;
            panel3.BorderStyle = BorderStyle.FixedSingle;
            panel3.Resize += panel3_Resize;
            statusStrip1.Paint += DrawBlackBorderWithClear;
            panel1.BorderStyle = BorderStyle.FixedSingle;
            panel2.BorderStyle = BorderStyle.FixedSingle;

           
            InitializeHistoryListBox();
            InitializeProgressTimer();
            CreateProgressNotification();
            DrawStatusBox();
            InitializePdfViewer();
            statusStrip1.Dock = DockStyle.Bottom;
            statusStrip1.Visible = true;

            // Убедимся что StatusStrip поверх других элементов
            statusStrip1.BringToFront();
            progressBar1.Visible = false;
        }

        // +++ ИСПРАВЛЕННАЯ ИНИЦИАЛИЗАЦИЯ PDF VIEWER +++
        private void InitializePdfViewer()
        {
            
        }

        // +++ МЕТОД ДЛЯ РЕНДЕРИНГА СТРАНИЦЫ PDF В BITMAP +++
        private Image RenderPageToBitmap(PdfDocument document, int pageIndex, int width, int height)
        {
            try
            {
                return document.Render(pageIndex, width, height, 96, 96, PdfRenderFlags.Annotations);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка рендеринга страницы: {ex.Message}");

                // Создаем пустое изображение в случае ошибки
                Bitmap errorImage = new Bitmap(width, height);
                using (Graphics g = Graphics.FromImage(errorImage))
                {
                    g.Clear(Color.White);
                    using (Font font = new Font("Arial", 12))
                    using (Brush brush = new SolidBrush(Color.Red))
                    {
                        g.DrawString("Ошибка загрузки страницы", font, brush, 10, 10);
                    }
                }
                return errorImage;
            }
        }

        // +++ ИСПРАВЛЕННЫЙ МЕТОД ДЛЯ ОТОБРАЖЕНИЯ PDF +++
        private void DisplayPdfInPdfViewer(string pdfPath)
        {
            try
            {
                if (!File.Exists(pdfPath))
                {
                    ShowErrorMessage($"PDF файл не найден: {pdfPath}");
                    CompleteProgress();
                    return;
                }

                panel3.SuspendLayout();
                RemoveWebBrowserOnly();


                pdfDocument = PdfDocument.Load(pdfPath);
                totalPages = pdfDocument.PageCount;
                currentPageIndex = 0;

                // Очищаем предыдущие страницы
                foreach (var page in pdfPages)
                {
                    page?.Dispose();
                }
                pdfPages.Clear();

                // Рендерим все страницы
                for (int i = 0; i < totalPages; i++)
                {
                    var pageSize = CalculatePageSizeForDisplay();
                    using (var image = pdfDocument.Render(i, pageSize.Width, pageSize.Height, 96, 96, PdfRenderFlags.Annotations))
                    {
                        pdfPages.Add(new Bitmap(image));
                    }
                }

                pdfContainerPanel = new Panel
                {
                    Dock = DockStyle.None,
                    Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                    AutoScroll = true,
                    BackColor = Color.FromArgb(240, 240, 240),
                    Padding = new Padding(40, 20, 40, 20),
                    Location = new Point(0, 0),
                    Size = new Size(panel3.Width, panel3.Height - statusStrip1.Height)
                };

                // Панель, содержащая страницы
                pagesPanel = new Panel
                {
                    AutoSize = true,
                    BackColor = Color.FromArgb(240, 240, 240)
                };

                int yOffset = 0;
                foreach (var pageImage in pdfPages)
                {
                    PictureBox pagePictureBox = new PictureBox
                    {
                        Image = pageImage,
                        Size = pageImage.Size,
                        Location = new Point(0, yOffset),
                        BackColor = Color.White,
                        BorderStyle = BorderStyle.FixedSingle,
                        SizeMode = PictureBoxSizeMode.AutoSize
                    };

                    pagesPanel.Controls.Add(pagePictureBox);
                    yOffset += pagePictureBox.Height + 20;
                }

                pagesPanel.Height = yOffset;

                // Добавляем масштабирование колесиком мыши
                pagesPanel.MouseWheel += PagesPanel_MouseWheel;
                pagesPanel.MouseEnter += (s, e) => pagesPanel.Focus();
             
                pdfContainerPanel.Controls.Add(pagesPanel);
                panel3.Controls.Add(pdfContainerPanel);

                // Центрирование страниц при изменении размера
                pdfContainerPanel.Resize += (s, e) =>
                {
                    foreach (Control ctrl in pagesPanel.Controls)
                    {
                        ctrl.Left = (pdfContainerPanel.ClientSize.Width - ctrl.Width) / 2;
                    }
                };

                // Первичное выравнивание
                foreach (Control ctrl in pagesPanel.Controls)
                {
                    ctrl.Left = (pdfContainerPanel.ClientSize.Width - ctrl.Width) / 2;
                }

                // Обновляем статус бар
                UpdateStatusBarInfo();

                panel3.ResumeLayout(true);
                CompleteProgressFinal();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка загрузки PDF: {ex.Message}");
                UseAlternativePdfViewer(pdfPath);
            }
        }
        private void RemoveWebBrowserOnly()
        {
            try
            {
                // Ищем webBrowser1 на panel3
                var webBrowser = panel3.Controls.OfType<WebBrowser>()
                    .FirstOrDefault(wb => wb.Name == "webBrowser1");

                if (webBrowser != null)
                {
                    panel3.Controls.Remove(webBrowser);
                    webBrowser.Dispose();
                }

                // Также удаляем PDF контент если он есть
                if (pdfContainerPanel != null && panel3.Controls.Contains(pdfContainerPanel))
                {
                    panel3.Controls.Remove(pdfContainerPanel);
                    pdfContainerPanel.Dispose();
                    pdfContainerPanel = null;
                }

                // Удаляем pagesPanel если он есть
                if (pagesPanel != null && panel3.Controls.Contains(pagesPanel))
                {
                    panel3.Controls.Remove(pagesPanel);
                    pagesPanel.Dispose();
                    pagesPanel = null;
                }

                // Удаляем все PictureBox (страницы PDF)
                var pictureBoxes = panel3.Controls.OfType<PictureBox>().ToList();
                foreach (var pb in pictureBoxes)
                {
                    panel3.Controls.Remove(pb);
                    pb.Dispose();
                }

                // Удаляем сообщения об ошибках
                var errorLabels = panel3.Controls.OfType<Label>()
                    .Where(l => l.Text.Contains("❌") || l.Text.Contains("Ошибка") || l.Text.Contains("PDF"))
                    .ToList();

                foreach (var label in errorLabels)
                {
                    panel3.Controls.Remove(label);
                    label.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка очистки webBrowser: {ex.Message}");
            }
        }
        private void PagesPanel_MouseWheel(object sender, MouseEventArgs e)
        {
            if (Control.ModifierKeys == Keys.Control)
            {
                // Zoom with Ctrl + Mouse Wheel
                if (e.Delta > 0)
                {
                    ZoomIn();
                }
                else
                {
                    ZoomOut();
                }
                // Убираем строку с e.Handled = true;
            }
        }

        private void PdfPictureBox_MouseWheel(object sender, MouseEventArgs e)
        {
            if (Control.ModifierKeys == Keys.Control)
            {
                // Zoom with Ctrl + Mouse Wheel
                if (e.Delta > 0)
                {
                    ZoomIn();
                }
                else
                {
                    ZoomOut();
                }
            }
        }
        // +++ ИСПРАВЛЕННЫЕ МЕТОДЫ ДЛЯ УПРАВЛЕНИЯ МАСШТАБОМ +++
        private void ApplyPdfViewerZoom()
        {
             if (pdfDocument != null)
            {
                RenderCurrentPageWithMargins();
            }
        }

        private void RefreshPdfDisplay()
        {
            if (!string.IsNullOrEmpty(currentPdfPath) && File.Exists(currentPdfPath))
            {
                DisplayPdfInPdfViewer(currentPdfPath);
            }
        }

        private void ZoomIn()
        {
            if (currentScale < maxScale)
            {
                currentScale += 0.1;
                RefreshAllPages();
                UpdateZoomStatus();
            }
        }

        private void ZoomOut()
        {
            if (currentScale > minScale)
            {
                currentScale -= 0.1;
                RefreshAllPages();
                UpdateZoomStatus();
            }
        }

        private void RefreshAllPages()
        {
            if (pdfDocument != null && !string.IsNullOrEmpty(currentPdfPath))
            {
                // Перезагружаем PDF с новым масштабом
                DisplayPdfInPdfViewer(currentPdfPath);
            }
        }
        private void UpdateZoomStatus()
        {
            // Можно добавить отображение масштаба в статус бар если нужно
            //toolStripStatusLabel3.Text = $"Масштаб: {(int)(currentScale * 100)}%";
        }
        // +++ ИСПРАВЛЕННЫЕ МЕТОДЫ ДЛЯ НАВИГАЦИИ ПО СТРАНИЦАМ +++
        private void GoToNextPage()
        {
            if (currentPageIndex < totalPages - 1)
            {
                currentPageIndex++;
                RenderCurrentPageWithMargins();
            }
        }

        private void GoToPreviousPage()
        {
            if (currentPageIndex > 0)
            {
                currentPageIndex--;
                RenderCurrentPageWithMargins();
            }
        }

        private void RenderCurrentPageWithMargins()
        {
            if (pdfDocument == null || currentPageIndex < 0 || currentPageIndex >= totalPages)
                return;

            try
            {
                // Рассчитываем размер для отображения с сохранением пропорций A4
                Size pageSize = CalculatePageSizeForDisplay();

                // Рендерим страницу
                using (var image = pdfDocument.Render(currentPageIndex,
                    pageSize.Width, pageSize.Height, 96, 96, PdfRenderFlags.Annotations))
                {
                    // Освобождаем предыдущее изображение
                    if (pdfPictureBox.Image != null)
                    {
                        pdfPictureBox.Image.Dispose();
                        pdfPictureBox.Image = null;
                    }

                    // Создаем новое изображение с белым фоном
                    Bitmap finalImage = new Bitmap(pageSize.Width, pageSize.Height);
                    using (Graphics g = Graphics.FromImage(finalImage))
                    {
                        g.Clear(Color.White);
                        g.DrawImage(image, 0, 0, pageSize.Width, pageSize.Height);
                    }

                    pdfPictureBox.Image = finalImage;
                    pdfPictureBox.Size = pageSize;

                    // Центрируем PictureBox в контейнере
                    CenterPictureBoxInContainer();
                }

                // Обновляем информацию о странице в статус баре
                UpdateStatusBarInfo();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка рендеринга страницы: {ex.Message}");
                ShowErrorMessage($"Ошибка отображения страницы: {ex.Message}");
            }
        }
        private void UpdateStatusBarInfo()
        {
            if (pdfDocument != null && totalPages > 0)
            {
                toolStripStatusLabel1.Text = $"Страниц: {totalPages}";

                // Убедимся что статус бар виден
                statusStrip1.Visible = true;
                statusStrip1.BringToFront();
            }
        }

        private void CenterPictureBoxInContainer()
        {
            if (pdfContainerPanel != null && pdfPictureBox != null)
            {
                // Центрируем по горизонтали
                int x = (pdfContainerPanel.ClientSize.Width - pdfPictureBox.Width) / 2;
                x = Math.Max(pdfContainerPanel.Padding.Left, x); // Не меньше левого отступа

                pdfPictureBox.Location = new Point(x, pdfContainerPanel.Padding.Top);
            }
        }

        private Size CalculatePageSizeForDisplay()
        {
            // Базовый размер A4 в пикселях (при 96 DPI)
            const int baseA4Width = 794;   // 210mm
            const int baseA4Height = 1123; // 297mm

            // Максимальная доступная ширина с учетом отступов (оставляем место для статус бара)
            int maxAvailableWidth = panel3.Width - 80 - 20; // 80px отступы, 20px для полосы прокрутки

            // Сохраняем пропорции A4
            double a4Ratio = (double)baseA4Height / baseA4Width;

            int displayWidth = Math.Min(maxAvailableWidth, baseA4Width);
            int displayHeight = (int)(displayWidth * a4Ratio);

            // Применяем масштаб
            displayWidth = (int)(displayWidth * currentScale);
            displayHeight = (int)(displayHeight * currentScale);

            // Ограничиваем минимальный размер
            displayWidth = Math.Max(400, displayWidth);
            displayHeight = Math.Max(565, displayHeight);

            return new Size(displayWidth, displayHeight);
        }

        private void ShowErrorMessage(string message)
        {
            panel3.Controls.Clear();

            Label errorLabel = new Label();
            errorLabel.Text = $"❌ {message}";
            errorLabel.TextAlign = ContentAlignment.MiddleCenter;
            errorLabel.Dock = DockStyle.Fill;
            errorLabel.Font = new Font("Arial", 10, FontStyle.Bold);
            errorLabel.ForeColor = Color.Red;
            errorLabel.Padding = new Padding(20);

            panel3.Controls.Add(errorLabel);
        }

        private void RefreshPanelLayout()
        {
            try
            {
                panel3.SuspendLayout();
                panel3.ResumeLayout(true);
                panel3.PerformLayout();

                panel3.Invalidate();
                panel3.Update();
                panel3.Refresh();

                Application.DoEvents();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обновления panel3: {ex.Message}");
            }
        }

        private void CleanupPdfViewer()
        {
            try
            {
                // Очищаем страницы
                foreach (var page in pdfPages)
                {
                    page?.Dispose();
                }
                pdfPages.Clear();

                if (pdfDocument != null)
                {
                    pdfDocument.Dispose();
                    pdfDocument = null;
                }

                // Не очищаем panel3 полностью, только контент PDF
                var pdfControls = panel3.Controls.OfType<Panel>()
                    .Where(p => p.Name == "pdfContainerPanel" || p.BackColor == Color.FromArgb(240, 240, 240))
                    .ToList();

                foreach (var control in pdfControls)
                {
                    panel3.Controls.Remove(control);
                    control.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка очистки PdfViewer: {ex.Message}");
            }
        }

        // +++ АЛЬТЕРНАТИВНЫЙ МЕТОД +++
        private void UseAlternativePdfViewer(string pdfPath)
        {
            try
            {
                panel3.Controls.Clear();

                Panel pdfPanel = new Panel();
                pdfPanel.Dock = DockStyle.Fill;
                pdfPanel.BackColor = Color.White;

                Label infoLabel = new Label();
                infoLabel.Text = "📄 PDF файл готов для просмотра\n\n" +
                                $"Файл: {Path.GetFileName(pdfPath)}\n\n" +
                                "Для просмотра нажмите кнопку ниже";
                infoLabel.TextAlign = ContentAlignment.MiddleCenter;
                infoLabel.Dock = DockStyle.Fill;
                infoLabel.Font = new Font("Arial", 10, FontStyle.Regular);

                Button openPdfButton = new Button();
                openPdfButton.Text = "Открыть PDF в программе по умолчанию";
                openPdfButton.Size = new Size(250, 40);
                openPdfButton.Location = new Point(
                    (panel3.Width - openPdfButton.Width) / 2,
                    panel3.Height - 60
                );
                openPdfButton.Anchor = AnchorStyles.Bottom;
                openPdfButton.Click += (s, e) =>
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = pdfPath,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Не удалось открыть PDF: {ex.Message}");
                    }
                };

                pdfPanel.Controls.Add(infoLabel);
                pdfPanel.Controls.Add(openPdfButton);
                panel3.Controls.Add(pdfPanel);

                CompleteProgressFinal();
            }
            catch (Exception ex)
            {
                ShowErrorMessage($"Ошибка отображения PDF: {ex.Message}");
                CompleteProgress();
            }
        }

        // +++ ОБНОВЛЕННЫЙ МЕТОД КОНВЕРТАЦИИ WORD В PDF +++
        private void ConvertWordToPdf(string filePath)
        {
            Word.Application wordApp = null;
            Word.Document wordDoc = null;

            try
            {
                currentPdfPath = Path.Combine(Path.GetTempPath(), $"temp_pdf_{Guid.NewGuid()}.pdf");

                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                UpdateProgressBarInStatus(30);

                wordDoc = wordApp.Documents.Open(
                    FileName: filePath,
                    ReadOnly: false,
                    Visible: false,
                    ConfirmConversions: false,
                    AddToRecentFiles: false
                );

                UpdateProgressBarInStatus(50);

                string fileName = Path.GetFileName(filePath);
                label3.Text = fileName;
                label3.ForeColor = Color.Black;

                int wordCount = wordDoc.ComputeStatistics(Word.WdStatistic.wdStatisticWords);
                int pageCount = wordDoc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);

                UpdateProgressBarInStatus(70);

                wordDoc.SaveAs2(
                    FileName: currentPdfPath,
                    FileFormat: Word.WdSaveFormat.wdFormatPDF,
                    AddToRecentFiles: false
                );

                UpdateProgressBarInStatus(85);

                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                wordDoc.Close(ref saveChanges);
                wordDoc = null;

                wordApp.Quit();
                wordApp = null;

                UpdateProgressBarInStatus(90);

                DisplayPdfInPdfViewer(currentPdfPath);

                // Обновляем статус бар
                ShowFileLoadedStatus(fileName, pageCount, wordCount);
                AddHistoryRecord($"Файл загружен: {fileName}");
            }
            catch (Exception ex)
            {
                CompleteProgress();
                MessageBox.Show($"Ошибка конвертации: {ex.Message}");
                ShowErrorStatus();
                UpdateButtonsState(false);
            }
            finally
            {
                try
                {
                    if (wordDoc != null)
                    {
                        object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                        wordDoc.Close(ref saveChanges);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
                        wordDoc = null;
                    }
                    if (wordApp != null)
                    {
                        wordApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                        wordApp = null;
                    }
                }
                catch (Exception cleanupEx)
                {
                    System.Diagnostics.Debug.WriteLine($"Ошибка очистки Word: {cleanupEx.Message}");
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        // +++ СОЗДАНИЕ ПРОГРЕСС-БАРА УВЕДОМЛЕНИЯ +++
        private void CreateProgressNotification()
        {
            if (progressNotificationPanel != null)
            {
                if (listBox1.Controls.Contains(progressNotificationPanel))
                {
                    listBox1.Controls.Remove(progressNotificationPanel);
                }
                progressNotificationPanel.Dispose();
                progressNotificationPanel = null;
            }

            progressNotificationPanel = new Panel();
            progressNotificationPanel.BackColor = Color.FromArgb(240, 240, 240);
            progressNotificationPanel.Height = 40;
            progressNotificationPanel.Dock = DockStyle.Top;
            progressNotificationPanel.Padding = new Padding(10);
            progressNotificationPanel.Visible = false;

            notificationProgressBar = new ProgressBar();
            notificationProgressBar.Dock = DockStyle.Top;
            notificationProgressBar.Height = 10;
            notificationProgressBar.Style = ProgressBarStyle.Continuous;
            notificationProgressBar.Minimum = 0;
            notificationProgressBar.Maximum = 100;
            notificationProgressBar.Value = 0;

            progressLabel = new Label();
            progressLabel.Dock = DockStyle.Top;
            progressLabel.Height = 15;
            progressLabel.Text = "Загрузка файла... 0%";
            progressLabel.TextAlign = ContentAlignment.MiddleLeft;
            progressLabel.Font = new Font("Arial", 8, FontStyle.Regular);

            progressNotificationPanel.Controls.Add(notificationProgressBar);
            progressNotificationPanel.Controls.Add(progressLabel);

            listBox1.Controls.Add(progressNotificationPanel);
            progressNotificationPanel.BringToFront();
        }

        private void ShowLoadingStatus()
        {
            CleanupPdfViewer();

            var oldLabels = panel3.Controls.OfType<Label>()
                .Where(l => l.Text.Contains("Конвертация") || l.Text.Contains("Ошибка") || l.Text.Contains("❌") || l.Text.Contains("🔄") || l.Text.Contains("✅"))
                .ToList();

            foreach (var label in oldLabels)
            {
                panel3.Controls.Remove(label);
                label.Dispose();
            }

            // Показываем статус что файл загружается
            statusStrip1.Visible = true;
            toolStripStatusLabel1.Text = "Загрузка файла...";
            toolStripStatusLabel2.Text = "";

            if (progressNotificationPanel != null)
            {
                progressNotificationPanel.Visible = true;
                progressNotificationPanel.BringToFront();
            }
        }

        private void ShowErrorStatus()
        {
            CompleteProgress();

            panel3.Controls.Clear();
            Label errorLabel = new Label();
            errorLabel.Text = "❌ Ошибка загрузки файла";
            errorLabel.TextAlign = ContentAlignment.MiddleCenter;
            errorLabel.Dock = DockStyle.Fill;
            errorLabel.Font = new Font("Arial", 12, FontStyle.Bold);
            errorLabel.ForeColor = Color.Red;
            panel3.Controls.Add(errorLabel);
        }

        // +++ НОВЫЙ МЕТОД ДЛЯ ПОКАЗА СТАТУСА "ФАЙЛ ЗАГРУЖЕН" +++
        private void ShowFileLoadedStatus(string fileName, int pageCount, int wordCount)
        {
            try
            {
                // Всегда обновляем статус бар и делаем его видимым
                toolStripStatusLabel1.Text = $"Страниц: {pageCount}";
                toolStripStatusLabel2.Text = $"Слов: {wordCount}";

                // Принудительно показываем статус бар
                statusStrip1.Visible = true;
                statusStrip1.BringToFront();

                // Обновляем информацию о файле
                label3.Text = fileName;
                label3.ForeColor = Color.Black;

                Timer timer = new Timer();
                timer.Interval = 3000;
                timer.Tick += (s, e) =>
                {
                    if (progressNotificationPanel != null && progressNotificationPanel.Visible)
                    {
                        progressNotificationPanel.Visible = false;
                    }
                    timer.Stop();
                    timer.Dispose();
                };
                timer.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка в ShowFileLoadedStatus: {ex.Message}");
            }
        }

        private void UpdateProgressInHistory(int progress)
        {
            try
            {
                string progressRecord = listBox1.Items
                    .OfType<string>()
                    .FirstOrDefault(item => item.Contains("Загрузка файла") );

                if (progressRecord != null)
                {
                    int index = listBox1.Items.IndexOf(progressRecord);
                    string timestamp = DateTime.Now.ToString("HH:mm:ss");
                    string newRecord = $"{timestamp} - Загрузка файла ";
                    listBox1.Items[index] = newRecord;
                }
                else
                {
                    string timestamp = DateTime.Now.ToString("HH:mm:ss");
                    string record = $"{timestamp} - Загрузка файла {progress}%";
                    listBox1.Items.Insert(0, record);

                    if (listBox1.Items.Count > 50)
                    {
                        listBox1.Items.RemoveAt(listBox1.Items.Count - 1);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обновления прогресса в истории: {ex.Message}");
            }
        }

        private void UpdateButtonsState(bool fileLoaded)
        {
            button3.Enabled = fileLoaded;
            button7.Enabled = fileLoaded;
            button6.Enabled = fileLoaded;

            if (fileLoaded)
            {
                ToggleTransparency(button3, EventArgs.Empty);
                button3.Invalidate();
                button7.Invalidate();
                button6.Invalidate();
            }
            else
            {
                if (activeToggleButton != null)
                {
                    activeToggleButton.Image = originalImages[activeToggleButton];
                    transparencyFixed[activeToggleButton] = false;
                    activeToggleButton.Invalidate();
                    activeToggleButton = null;
                }

                button3.Image = originalImages[button3];
                button7.Image = originalImages[button7];
                button6.Image = originalImages[button6];
                transparencyFixed[button3] = false;
                transparencyFixed[button7] = false;
                transparencyFixed[button6] = false;
            }
        }

        private void InitializeToolTips()
        {
            toolTip = new ToolTip();
            toolTip.AutoPopDelay = 5000;
            toolTip.InitialDelay = 500;
            toolTip.ReshowDelay = 100;
            toolTip.ShowAlways = true;

            toolTip.SetToolTip(button8, "Загрузить файл");
            toolTip.SetToolTip(button2, "Обработать файл");
            toolTip.SetToolTip(button1, "Сохранить файл");
            toolTip.SetToolTip(button3, "Исходный документ");
            toolTip.SetToolTip(button7, "Обработанный документ с комментариями");
            toolTip.SetToolTip(button6, "Обработанный документ без комментариев");
            toolTip.SetToolTip(button4, "Настройки");
            toolTip.SetToolTip(button5, "Справка");
        }

        private void InitializeButtonHoverEffects()
        {
            Button[] buttons = { button1, button2, button3, button4, button5, button6, button7, button8 };

            foreach (Button btn in buttons)
            {
                if (btn != null && btn.Image != null)
                {
                    btn.FlatStyle = FlatStyle.Flat;
                    btn.FlatAppearance.BorderSize = 0;
                    btn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
                    btn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
                    btn.BackColor = System.Drawing.Color.Transparent;

                    originalImages[btn] = btn.Image;
                    transparencyFixed[btn] = false;

                    btn.MouseEnter += (sender, e) =>
                    {
                        Button button = sender as Button;
                        if (button.Enabled && !transparencyFixed[button])
                        {
                            button.Image = ApplyAlpha(originalImages[button], 0.3f);
                            button.Invalidate();
                        }
                    };

                    btn.MouseLeave += (sender, e) =>
                    {
                        Button button = sender as Button;
                        if (button.Enabled && !transparencyFixed[button])
                        {
                            button.Image = originalImages[button];
                            button.Invalidate();
                        }
                    };

                    btn.Paint += DrawRoundedBorder;

                    if (btn == button3 || btn == button7 || btn == button6)
                    {
                        btn.Click += ToggleTransparency;
                    }
                }
            }

            ConfigureRoundedCorners();
        }

        private void ToggleTransparency(object sender, EventArgs e)
        {
            Button clickedButton = sender as Button;
            if (clickedButton == null || !transparencyFixed.ContainsKey(clickedButton) || !clickedButton.Enabled)
                return;

            if (activeToggleButton == clickedButton)
            {
                clickedButton.Image = originalImages[clickedButton];
                transparencyFixed[clickedButton] = false;
                activeToggleButton = null;
            }
            else
            {
                if (activeToggleButton != null)
                {
                    activeToggleButton.Image = originalImages[activeToggleButton];
                    transparencyFixed[activeToggleButton] = false;
                    activeToggleButton.Invalidate();
                }

                clickedButton.Image = ApplyAlpha(originalImages[clickedButton], 0.3f);
                transparencyFixed[clickedButton] = true;
                activeToggleButton = clickedButton;
            }

            clickedButton.Invalidate();
        }

        private void ConfigureRoundedCorners()
        {
            buttonCorners[button1] = (0, 0, 8, 8);
            buttonCorners[button3] = (8, 8, 0, 0);
            buttonCorners[button4] = (8, 8, 8, 8);
            buttonCorners[button5] = (8, 8, 8, 8);
            buttonCorners[button6] = (0, 0, 8, 8);
            buttonCorners[button8] = (8, 8, 0, 0);

            if (button2 != null) buttonCorners[button2] = (0, 0, 0, 0);
            if (button7 != null) buttonCorners[button7] = (0, 0, 0, 0);

            ApplyRoundedRegions();
        }

        private void ApplyRoundedRegions()
        {
            foreach (var kvp in buttonCorners)
            {
                Button button = kvp.Key;
                var (topLeft, topRight, bottomRight, bottomLeft) = kvp.Value;

                if (topLeft == 0 && topRight == 0 && bottomRight == 0 && bottomLeft == 0)
                {
                    button.Region = null;
                }
                else
                {
                    GraphicsPath path = new GraphicsPath();
                    int width = button.Width;
                    int height = button.Height;

                    if (topLeft > 0)
                        path.AddArc(0, 0, topLeft * 2, topLeft * 2, 180, 90);
                    else
                        path.AddLine(0, 0, 0, 0);

                    path.AddLine(topLeft, 0, width - topRight, 0);

                    if (topRight > 0)
                        path.AddArc(width - topRight * 2, 0, topRight * 2, topRight * 2, 270, 90);
                    else
                        path.AddLine(width, 0, width, 0);

                    path.AddLine(width, topRight, width, height - bottomRight);

                    if (bottomRight > 0)
                        path.AddArc(width - bottomRight * 2, height - bottomRight * 2, bottomRight * 2, bottomRight * 2, 0, 90);
                    else
                        path.AddLine(width, height, width, height);

                    path.AddLine(width - bottomRight, height, bottomLeft, height);

                    if (bottomLeft > 0)
                        path.AddArc(0, height - bottomLeft * 2, bottomLeft * 2, bottomLeft * 2, 90, 90);
                    else
                        path.AddLine(0, height, 0, height);

                    path.AddLine(0, height - bottomLeft, 0, topLeft);

                    path.CloseFigure();

                    button.Region = new Region(path);
                }
            }
        }

        private void DrawRoundedBorder(object sender, PaintEventArgs e)
        {
            Button button = sender as Button;
            if (button == null || !buttonCorners.ContainsKey(button)) return;
            var (topLeft, topRight, bottomRight, bottomLeft) = buttonCorners[button];

            float borderWidth = 2f;

            using (Pen borderPen = new Pen(System.Drawing.Color.FromArgb(80, 80, 80), borderWidth))
            using (GraphicsPath path = new GraphicsPath())
            {
                int width = button.Width - (int)borderWidth;
                int height = button.Height - (int)borderWidth;

                if (topLeft > 0)
                    path.AddArc(0, 0, topLeft * 2, topLeft * 2, 180, 90);
                else
                    path.AddLine(0, 0, 0, 0);

                path.AddLine(topLeft, 0, width - topRight, 0);

                if (topRight > 0)
                    path.AddArc(width - topRight * 2, 0, topRight * 2, topRight * 2, 270, 90);
                else
                    path.AddLine(width, 0, width, 0);

                path.AddLine(width, topRight, width, height - bottomRight);

                if (bottomRight > 0)
                    path.AddArc(width - bottomRight * 2, height - bottomRight * 2, bottomRight * 2, bottomRight * 2, 0, 90);
                else
                    path.AddLine(width, height, width, height);

                path.AddLine(width - bottomRight, height, bottomLeft, height);

                if (bottomLeft > 0)
                    path.AddArc(0, height - bottomLeft * 2, bottomLeft * 2, bottomLeft * 2, 90, 90);
                else
                    path.AddLine(0, height, 0, height);

                path.AddLine(0, height - bottomLeft, 0, topLeft);

                path.CloseFigure();

                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.DrawPath(borderPen, path);
            }
        }

        private Image ApplyAlpha(Image originalImage, float alpha)
        {
            Bitmap result = new Bitmap(originalImage.Width, originalImage.Height);

            using (Graphics g = Graphics.FromImage(result))
            {
                ColorMatrix matrix = new ColorMatrix();
                matrix.Matrix33 = alpha;

                ImageAttributes attributes = new ImageAttributes();
                attributes.SetColorMatrix(matrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

                g.DrawImage(originalImage,
                           new Rectangle(0, 0, originalImage.Width, originalImage.Height),
                           0, 0, originalImage.Width, originalImage.Height,
                           GraphicsUnit.Pixel, attributes);
            }

            return result;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Word Documents|*.doc;*.docx";
                openFileDialog.Title = "Выберите Word файл";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ShowLoadingStatus();
                        CreateProgressNotification();

                        if (progressNotificationPanel != null)
                        {
                            progressNotificationPanel.Visible = true;
                            progressNotificationPanel.BringToFront();
                            progressNotificationPanel.Refresh();
                            Application.DoEvents();
                        }

                        progressValue = 0;
                        UpdateProgressBarInStatus(1);
                        progressTimer.Start();
                        Application.DoEvents();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Ошибка инициализации прогресс-бара: {ex.Message}");
                    }

                    try
                    {
                        ConvertWordToPdf(openFileDialog.FileName);
                    }
                    catch (Exception ex)
                    {
                        CompleteProgress();
                        MessageBox.Show($"Ошибка загрузки файла: {ex.Message}");
                    }
                }
            }
        }

        private void ForceUIUpdate()
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(ForceUIUpdate));
                    return;
                }

                Application.DoEvents();

                if (progressNotificationPanel != null)
                {
                    progressNotificationPanel.Update();
                    progressNotificationPanel.Refresh();
                }

                listBox1.Update();
                listBox1.Refresh();

                statusStrip1.Update();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка при обновлении UI: {ex.Message}");
            }
        }

        private void InitializeHistoryListBox()
        {
            listBox1.DrawMode = DrawMode.OwnerDrawVariable;
            listBox1.ItemHeight = 40;
            listBox1.DrawItem += ListBox1_DrawItem;
            listBox1.MeasureItem += ListBox1_MeasureItem;
        }

        private void ListBox1_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            e.ItemHeight = 40;
        }

        private void ListBox1_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            e.DrawBackground();

            Rectangle rect = e.Bounds;
            using (System.Drawing.Brush brush = new SolidBrush(System.Drawing.Color.FromArgb(240, 240, 240)))
            {
                e.Graphics.FillRectangle(brush, rect);
            }

            using (System.Drawing.Pen pen = new System.Drawing.Pen(System.Drawing.Color.Gray, 1))
            {
                e.Graphics.DrawRectangle(pen, rect);
            }

            string text = listBox1.Items[e.Index].ToString();
            using (System.Drawing.Brush textBrush = new SolidBrush(e.ForeColor))
            {
                StringFormat format = new StringFormat();
                format.LineAlignment = StringAlignment.Center;
                format.Alignment = StringAlignment.Near;
                e.Graphics.DrawString(text, e.Font, textBrush,
                    new Rectangle(rect.X + 5, rect.Y, rect.Width - 10, rect.Height), format);
            }

            e.DrawFocusRectangle();
        }

        private void AddHistoryRecord(string operation)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string record = $"{timestamp} - {operation}";

            listBox1.Items.Insert(0, record);

            if (listBox1.Items.Count > 50)
            {
                listBox1.Items.RemoveAt(listBox1.Items.Count - 1);
            }
        }

        private void DrawStatusBox()
        {
            System.Drawing.Bitmap statusImage = new System.Drawing.Bitmap(panel3.Width, panel3.Height);
            using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(statusImage))
            {
                g.Clear(System.Drawing.Color.White);

                using (System.Drawing.Pen borderPen = new System.Drawing.Pen(System.Drawing.Color.Gray, 2))
                {
                    System.Drawing.Rectangle borderRect = new System.Drawing.Rectangle(5, 5, statusImage.Width - 10, statusImage.Height - 10);
                    g.DrawRectangle(borderPen, borderRect);
                }

                using (System.Drawing.Font statusFont = new System.Drawing.Font("Arial", 10))
                using (System.Drawing.Brush textBrush = new SolidBrush(System.Drawing.Color.Black))
                {
                    string statusText = "Статус\n\nПриложение готово к работе\nЗапущен процесс загрузки документа\n\n49%";
                    System.Drawing.StringFormat format = new System.Drawing.StringFormat();
                    format.Alignment = System.Drawing.StringAlignment.Near;
                    format.LineAlignment = System.Drawing.StringAlignment.Near;

                    System.Drawing.Rectangle textRect = new System.Drawing.Rectangle(15, 15, statusImage.Width - 30, statusImage.Height - 30);
                    g.DrawString(statusText, statusFont, textBrush, textRect, format);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AddHistoryRecord("Обработка файла");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddHistoryRecord("Сохранить файл");
        }

        // +++ ОБРАБОТЧИК ИЗМЕНЕНИЯ РАЗМЕРА PANEL3 +++
        private void panel3_Resize(object sender, EventArgs e)
        {
            if (pdfDocument != null && pdfPictureBox != null)
            {
                RenderCurrentPageWithMargins();
            }
            RefreshPanelLayout();
        }

        // +++ МЕТОДЫ ДЛЯ ПОЛУЧЕНИЯ ИНФОРМАЦИИ О ДОКУМЕНТЕ +++
        private void UpdateStatusWithDocumentInfo()
        {
            try
            {
                string fileName = label3.Text;
                int wordCount = GetWordCountFromStatus();
                int pageCount = GetPageCountFromStatus();

                toolStripStatusLabel1.Text = $"Страниц: {pageCount}";
                toolStripStatusLabel2.Text = $"Слов: {wordCount}";
            }
            catch
            {
            }
        }

        private int GetPageCountFromStatus()
        {
            try
            {
                string pagesText = toolStripStatusLabel1.Text;
                if (pagesText.Contains("Страниц:"))
                {
                    string numberPart = pagesText.Replace("Страниц:", "").Trim();
                    if (int.TryParse(numberPart, out int pageCount))
                    {
                        return pageCount;
                    }
                }
            }
            catch
            {
            }
            return 1;
        }

        private int GetWordCountFromStatus()
        {
            try
            {
                string wordsText = toolStripStatusLabel2.Text;
                if (wordsText.Contains("Слов:"))
                {
                    string numberPart = wordsText.Replace("Слов:", "").Trim();
                    if (int.TryParse(numberPart, out int wordCount))
                    {
                        return wordCount;
                    }
                }
            }
            catch
            {
            }
            return 0;
        }

       
    }
}