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

namespace ПО
{
    public partial class Form1 : Form
    {
        // Словарь для хранения оригинальных изображений
        private Dictionary<Button, Image> originalImages = new Dictionary<Button, Image>();
        // Словарь для хранения параметров закругления кнопок
        private Dictionary<Button, (int topLeft, int topRight, int bottomRight, int bottomLeft)> buttonCorners =
            new Dictionary<Button, (int, int, int, int)>();
        // Словарь для отслеживания состояния фиксации прозрачности
        private Dictionary<Button, bool> transparencyFixed = new Dictionary<Button, bool>();

        public Form1()
        {
            InitializeComponent();
            InitializeButtonHoverEffects();
        }

        private void InitializeButtonHoverEffects()
        {
            Button[] buttons = { button1, button2, button3, button4,
                button5, button6, button7, button8 };

            // Темно-серая рамка
            Color darkGrayBorder = Color.FromArgb(80, 80, 80);

            foreach (Button btn in buttons)
            {
                if (btn != null && btn.Image != null)
                {
                    // Настраиваем стиль кнопки
                    btn.FlatStyle = FlatStyle.Flat;
                    btn.FlatAppearance.BorderSize = 0; // Убираем стандартную рамку
                    btn.FlatAppearance.MouseOverBackColor = Color.Transparent;
                    btn.FlatAppearance.MouseDownBackColor = Color.Transparent;
                    btn.BackColor = Color.Transparent;

                    // Сохраняем оригинальное изображение
                    originalImages[btn] = btn.Image;

                    // Инициализируем состояние фиксации прозрачности
                    transparencyFixed[btn] = false;

                    btn.MouseEnter += (sender, e) =>
                    {
                        Button button = sender as Button;
                        // Если прозрачность не зафиксирована, применяем эффект наведения
                        if (!transparencyFixed[button])
                        {
                            button.Image = ApplyAlpha(originalImages[button], 0.3f);
                            button.Invalidate(); // Перерисовываем для обновления рамки
                        }
                    };

                    btn.MouseLeave += (sender, e) =>
                    {
                        Button button = sender as Button;
                        // Если прозрачность не зафиксирована, возвращаем оригинальное изображение
                        if (!transparencyFixed[button])
                        {
                            button.Image = originalImages[button];
                            button.Invalidate(); // Перерисовываем для обновления рамки
                        }
                    };

                    // Добавляем обработчик для отрисовки рамки
                    btn.Paint += DrawRoundedBorder;

                    // Добавляем обработчик клика для кнопок button3, button7, button6
                    if (btn == button3 || btn == button7 || btn == button6)
                    {
                        btn.Click += ToggleTransparency;
                    }
                }
            }

            // Настраиваем закругления
            ConfigureRoundedCorners();
        }

        // Обработчик переключения фиксации прозрачности
        private void ToggleTransparency(object sender, EventArgs e)
        {
            Button button = sender as Button;
            if (button == null || !transparencyFixed.ContainsKey(button)) return;

            // Переключаем состояние
            transparencyFixed[button] = !transparencyFixed[button];

            if (transparencyFixed[button])
            {
                // Фиксируем прозрачность
                button.Image = ApplyAlpha(originalImages[button], 0.3f);
            }
            else
            {
                // Возвращаем оригинальное изображение
                button.Image = originalImages[button];
            }

            button.Invalidate(); // Перерисовываем для обновления
        }

        private void ConfigureRoundedCorners()
        {
            // Сохраняем параметры закругления для каждой кнопки
            buttonCorners[button1] = (0, 0, 8, 8);    // Нижние углы
            buttonCorners[button3] = (8, 8, 0, 0);    // Верхние углы
            buttonCorners[button4] = (8, 8, 8, 8);    // Все углы
            buttonCorners[button5] = (8, 8, 8, 8);    // Все углы
            buttonCorners[button6] = (0, 0, 8, 8);    // Нижние углы
            buttonCorners[button8] = (8, 8, 0, 0);    // Верхние углы

            // Для кнопок без закругления тоже добавляем (радиус 0)
            if (button2 != null) buttonCorners[button2] = (0, 0, 0, 0);
            if (button7 != null) buttonCorners[button7] = (0, 0, 0, 0);

            // Применяем закругленные регионы
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
                    // Для кнопок без закругления сбрасываем регион
                    button.Region = null;
                }
                else
                {
                    // Создаем закругленный регион
                    GraphicsPath path = new GraphicsPath();
                    int width = button.Width;
                    int height = button.Height;

                    // Создаем путь по внешнему краю кнопки
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

            float borderWidth = 2f; // ← измените это значение

            using (Pen borderPen = new Pen(Color.FromArgb(80, 80, 80), borderWidth))
            using (GraphicsPath path = new GraphicsPath())
            {
                int width = button.Width - (int)borderWidth;
                int height = button.Height - (int)borderWidth;

                // Создаем путь для рамки с учетом закруглений
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

                // Рисуем рамку
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.DrawPath(borderPen, path);
            }
        }

        // Метод для применения альфа-канала (затемнения)
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
    }
}