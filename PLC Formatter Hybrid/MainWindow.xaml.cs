using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace TagFormatterApp
{
    public partial class MainWindow : Window
    {
        private List<TextBox> columnTextBoxes = new List<TextBox>();
        private TextBox lastFocusedColumnBox = null;

        private readonly string defaultAutoTemplate =
@"    #""0320V005""(Simulation_On := ""General"".Simulation_On,
                Emergency := ""Emr_PST"",
                Reset := ""General"".Reset,
                Sim_Opn_fb_Time := 1.0,
                Sim_Cls_fb_Time := 1.0,
                FeedbackTime := 5,
                CNT_Type := 4,
                ResetType := 1,
                qOpen => ""0320V005"",
                Valve := ""VLV"".""{0320V005}"");".Replace("{0320V005}", "0320V005");

        private readonly string defaultManualTemplate =
@"    #""{0}""(Signal := ""{1}"",
                Emergency := ""Emr_PST"",
                Reset := ""General"".Reset);";

        public MainWindow()
        {
            InitializeComponent();
            TxtTemplate.Text = defaultAutoTemplate;

            // Manuel mod için başlangıçta 2 sütun ({0} ve {1})
            AddInputColumn("0320V001\n0320V002");
            AddInputColumn("Sensor_A\nSensor_B");
        }

        private void Mode_Changed(object sender, RoutedEventArgs e)
        {
            if (GrpAutoInput == null || GrpManualInput == null || TxtTemplate == null || LblTemplateHeader == null) return;

            TxtOutput.Clear();

            if (RbAutoMode.IsChecked == true)
            {
                LblTemplateHeader.Text = "Örnek SCL Blok Çıktısı (İçindeki #\"ETİKET\" otomatik algılanır)  ";
                GrpAutoInput.Visibility = Visibility.Visible;
                GrpManualInput.Visibility = Visibility.Collapsed;
                TxtTemplate.Text = defaultAutoTemplate;
            }
            else
            {
                LblTemplateHeader.Text = "PLC SCL Blok Şablonu (Aşağıdaki kutucuk numaralarına göre {0}, {1}, {2}... kullanın)  ";
                GrpAutoInput.Visibility = Visibility.Collapsed;
                GrpManualInput.Visibility = Visibility.Visible;
                TxtTemplate.Text = defaultManualTemplate;
            }
        }

        private void BtnClearTemplate_Click(object sender, RoutedEventArgs e)
        {
            TxtTemplate.Clear();
            TxtTemplate.Focus();
        }

        private void BtnCopyOutput_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtOutput.Text))
            {
                MessageBox.Show("Kopyalanacak çıktı bulunamadı. Önce 'Dönüştür' butonuna basın.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Clipboard.SetText(TxtOutput.Text);

            // Kullanıcıya kopyalandığını görsel olarak hissettir
            BtnCopyOutput.Content = "Kopyalandı ✓";
            var timer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromSeconds(1.5) };
            timer.Tick += (s, args) =>
            {
                BtnCopyOutput.Content = "Tümünü Kopyala";
                timer.Stop();
            };
            timer.Start();
        }

        #region Excel'den Etiket Çekme İşlemi

        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Dosyaları (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                Title = "Etiketlerin Alınacağı Excel Dosyasını Seçin"
            };

            if (openFileDialog.ShowDialog() != true) return;

            // Sayfa boş bırakıldıysa boş string olarak bırakılır (aşağıda ilk sayfa seçilecek)
            string sheetName = TxtExcelSheet.Text.Trim();

            // Sütun boş bırakıldıysa varsayılan olarak "A" alınır
            string colLetter = TxtExcelCol.Text.Trim().ToUpper();
            if (string.IsNullOrEmpty(colLetter))
            {
                colLetter = "A";
            }

            // Başlangıç satırı kontrolü (Boş bırakıldıysa default 1)
            int startRow = 1;
            string startText = TxtExcelStartRow.Text.Trim();
            if (!string.IsNullOrEmpty(startText))
            {
                if (!int.TryParse(startText, out startRow) || startRow < 1)
                {
                    MessageBox.Show("Başlangıç satırı 1 veya daha büyük bir sayı olmalıdır.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            // Bitiş satırı kontrolü (Boş bırakıldıysa dinamik son satır)
            int? endRow = null;
            string endText = TxtExcelEndRow.Text.Trim();
            if (!string.IsNullOrEmpty(endText))
            {
                if (!int.TryParse(endText, out int parsedEnd) || parsedEnd < 1)
                {
                    MessageBox.Show("Bitiş satırı 1 veya daha büyük bir sayı olmalıdır.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                endRow = parsedEnd;
            }

            try
            {
                var extractedTags = new List<string>();

                using (var stream = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var workbook = new XLWorkbook(stream))
                    {
                        IXLWorksheet worksheet;

                        // Sayfa adı boş bırakıldıysa doğrudan ilk sayfayı al
                        if (string.IsNullOrEmpty(sheetName))
                        {
                            worksheet = workbook.Worksheets.First();
                        }
                        else if (!workbook.TryGetWorksheet(sheetName, out worksheet))
                        {
                            var availableSheets = string.Join(", ", workbook.Worksheets.Select(w => w.Name));
                            MessageBox.Show($"'{sheetName}' adında bir sayfa bulunamadı.\nDosyadaki mevcut sayfalar: {availableSheets}", "Sayfa Bulunamadı", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        var lastRowUsed = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                        if (lastRowUsed == 0 || lastRowUsed < startRow)
                        {
                            MessageBox.Show("Seçilen başlangıç satırından itibaren veri bulunamadı.", "Bilgi", MessageBoxButton.OK, MessageBoxImage.Information);
                            return;
                        }

                        int targetEndRow = endRow.HasValue ? Math.Min(endRow.Value, lastRowUsed) : lastRowUsed;

                        if (startRow > targetEndRow)
                        {
                            MessageBox.Show("Başlangıç satırı bitiş satırından büyük olamaz.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        for (int row = startRow; row <= targetEndRow; row++)
                        {
                            string cellValue = worksheet.Cell($"{colLetter}{row}").GetString().Trim();
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                extractedTags.Add(cellValue);
                            }
                        }
                    }
                }

                if (extractedTags.Count == 0)
                {
                    MessageBox.Show($"Belirtilen aralıkta ('{colLetter}' sütunu) aktarılacak veri bulunamadı.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string joinedData = string.Join(Environment.NewLine, extractedTags);

                if (RbAutoMode.IsChecked == true)
                {
                    TxtAutoInput.Text = joinedData;
                }
                else
                {
                    TextBox targetBox = lastFocusedColumnBox != null && columnTextBoxes.Contains(lastFocusedColumnBox)
                        ? lastFocusedColumnBox
                        : columnTextBoxes.First();

                    targetBox.Text = joinedData;
                }

                MessageBox.Show($"{extractedTags.Count} adet etiket başarıyla yüklendi!", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excel okuma hatası: {ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Manuel Mod Dinamik Sütun Yönetimi

        private void AddInputColumn(string defaultText = "")
        {
            int index = columnTextBoxes.Count;

            var groupBox = new GroupBox
            {
                Header = $"{{{index}}} Etiketleri",
                Margin = new Thickness(3),
                MinWidth = 140,
                FontWeight = FontWeights.SemiBold
            };

            var textBox = new TextBox
            {
                AcceptsReturn = true,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                FontFamily = new FontFamily("Consolas"),
                FontSize = 13,
                FontWeight = FontWeights.Normal,
                Padding = new Thickness(4),
                Text = defaultText
            };

            // Kullanıcı hangi sütun kutusuna tıklarsa onu hatırla (Excel'den çekince oraya dolsun)
            textBox.GotFocus += (s, e) => { lastFocusedColumnBox = textBox; };

            groupBox.Content = textBox;
            ColumnsContainer.Children.Add(groupBox);
            columnTextBoxes.Add(textBox);
        }

        private void BtnAddColumn_Click(object sender, RoutedEventArgs e)
        {
            AddInputColumn();
        }

        private void BtnRemoveColumn_Click(object sender, RoutedEventArgs e)
        {
            if (columnTextBoxes.Count <= 1)
            {
                MessageBox.Show("En az bir girdi sütunu ({0}) kalmalıdır.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int lastIndex = columnTextBoxes.Count - 1;
            if (lastFocusedColumnBox == columnTextBoxes[lastIndex])
            {
                lastFocusedColumnBox = null;
            }
            ColumnsContainer.Children.RemoveAt(lastIndex);
            columnTextBoxes.RemoveAt(lastIndex);
        }

        #endregion

        #region Dönüştürme İşlemleri

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            if (RbAutoMode.IsChecked == true)
            {
                ConvertAutoMode();
            }
            else
            {
                ConvertManualMode();
            }
        }

        private void ConvertAutoMode()
        {
            try
            {
                string sampleText = TxtTemplate.Text;
                if (string.IsNullOrWhiteSpace(sampleText))
                {
                    MessageBox.Show("Lütfen örnek bir SCL blok çıktısı girin.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var inputTags = TxtAutoInput.Text
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(t => t.Trim())
                    .Where(t => !string.IsNullOrEmpty(t))
                    .ToArray();

                if (inputTags.Length == 0)
                {
                    MessageBox.Show("Lütfen dönüştürülecek en az bir etiket girin.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var match = Regex.Match(sampleText, @"#""([^""]+)""");
                if (!match.Success)
                {
                    MessageBox.Show("Örnek blok içinde #\"ETİKET\" formatında ana bir başlık bulunamadı.", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string sampleTag = match.Groups[1].Value;
                string dynamicTemplate = sampleText.Replace(sampleTag, "{0}");

                var formattedBlocks = inputTags.Select(tag => string.Format(dynamicTemplate, tag));
                TxtOutput.Text = string.Join(Environment.NewLine + Environment.NewLine, formattedBlocks);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Otomatik dönüştürme hatası: {ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ConvertManualMode()
        {
            try
            {
                string template = TxtTemplate.Text;
                if (string.IsNullOrWhiteSpace(template))
                {
                    MessageBox.Show("Lütfen bir şablon girin.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var allColumnsData = new List<string[]>();
                int maxRowCount = 0;

                foreach (var tb in columnTextBoxes)
                {
                    var lines = tb.Text
                        .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(l => l.Trim())
                        .ToArray();

                    allColumnsData.Add(lines);
                    if (lines.Length > maxRowCount)
                    {
                        maxRowCount = lines.Length;
                    }
                }

                if (maxRowCount == 0)
                {
                    MessageBox.Show("Lütfen en az bir sütuna veri girin.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var formattedBlocks = new List<string>();

                for (int row = 0; row < maxRowCount; row++)
                {
                    object[] rowArgs = new object[allColumnsData.Count];

                    for (int col = 0; col < allColumnsData.Count; col++)
                    {
                        rowArgs[col] = row < allColumnsData[col].Length ? allColumnsData[col][row] : "";
                    }

                    formattedBlocks.Add(string.Format(template, rowArgs));
                }

                TxtOutput.Text = string.Join(Environment.NewLine + Environment.NewLine, formattedBlocks);
            }
            catch (FormatException)
            {
                MessageBox.Show("Şablondaki en yüksek {indeks} numarası için yeterli sütun eklediğinizden emin olun.", "Şablon Hatası", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Manuel dönüştürme hatası: {ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            TxtOutput.Clear();
            if (RbAutoMode.IsChecked == true)
            {
                TxtAutoInput.Clear();
            }
            else
            {
                foreach (var tb in columnTextBoxes)
                {
                    tb.Clear();
                }
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtOutput.Text))
            {
                MessageBox.Show("Kaydedilecek çıktı bulunamadı. Önce 'Dönüştür' butonuna basın.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Text file (*.txt)|*.txt|All files (*.*)|*.*",
                FileName = "TagFlow_Output.txt"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                File.WriteAllText(saveFileDialog.FileName, TxtOutput.Text);
                MessageBox.Show("Dosya başarıyla kaydedildi!", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}