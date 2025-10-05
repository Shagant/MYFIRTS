using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ExcelMerger
{
    public class MainForm : Form
    {
        private Button btnSelectFiles;
        private Button btnSaveFile;
        private Button btnMerge;
        private ListBox lstFiles;
        private string[] selectedFiles = new string[0];
        private string saveFilePath;

        public MainForm()
        {
            Text = "Excel Merger";
            Width = 600;
            Height = 380;
            StartPosition = FormStartPosition.CenterScreen;

            btnSelectFiles = new Button { Text = "Выбрать файлы...", Left = 20, Top = 20, Width = 140, Height = 30 };
            btnSaveFile    = new Button { Text = "Сохранить как...", Left = 180, Top = 20, Width = 140, Height = 30 };
            btnMerge       = new Button { Text = "Объединить", Left = 340, Top = 20, Width = 140, Height = 30 };

            lstFiles = new ListBox { Left = 20, Top = 70, Width = 540, Height = 220 };

            var lblInfo = new Label { Left = 20, Top = 300, Width = 540, Height = 40, Text = "Выберите Excel (.xlsx) файлы. Пустые строки удаляются.", AutoSize = false };

            btnSelectFiles.Click += BtnSelectFiles_Click;
            btnSaveFile.Click += BtnSaveFile_Click;
            btnMerge.Click += BtnMerge_Click;

            Controls.Add(btnSelectFiles);
            Controls.Add(btnSaveFile);
            Controls.Add(btnMerge);
            Controls.Add(lstFiles);
            Controls.Add(lblInfo);
        }

        private void BtnSelectFiles_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Multiselect = true;
                ofd.Filter = "Excel files|*.xlsx";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    selectedFiles = ofd.FileNames;
                    lstFiles.Items.Clear();
                    lstFiles.Items.AddRange(selectedFiles);
                }
            }
        }

        private void BtnSaveFile_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel files|*.xlsx";
                sfd.FileName = "merged.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    saveFilePath = sfd.FileName;
                    MessageBox.Show($"Файл будет сохранён как:\n{saveFilePath}", "Сохранение");
                }
            }
        }

        private void BtnMerge_Click(object sender, EventArgs e)
        {
            if (selectedFiles == null || selectedFiles.Length == 0)
            {
                MessageBox.Show("Сначала выберите Excel файлы!", "Ошибка");
                return;
            }

            if (string.IsNullOrEmpty(saveFilePath))
            {
                MessageBox.Show("Укажите путь для сохранения (Save as)...", "Ошибка");
                return;
            }

            try
            {
                DataTable mergedTable = null;

                foreach (var file in selectedFiles)
                {
                    using (var wb = new ClosedXML.Excel.XLWorkbook(file))
                    {
                        var ws = wb.Worksheets.First();
                        var rangeUsed = ws.RangeUsed();
                        if (rangeUsed == null) continue;

                        var dt = rangeUsed.AsTable().AsNativeDataTable();

                        var rowsToKeep = dt.Rows.Cast<DataRow>()
                            .Where(r => r.ItemArray.Any(v => v != null && !string.IsNullOrWhiteSpace(v.ToString())))
                            .ToArray();

                        var dtFiltered = dt.Clone();
                        foreach (var r in rowsToKeep) dtFiltered.ImportRow(r);

                        if (mergedTable == null)
                            mergedTable = dtFiltered.Clone();

                        foreach (DataRow row in dtFiltered.Rows)
                            mergedTable.ImportRow(row);
                    }
                }

                using (var wbOut = new ClosedXML.Excel.XLWorkbook())
                {
                    wbOut.Worksheets.Add(mergedTable, "Объединение");
                    wbOut.SaveAs(saveFilePath);
                }

                MessageBox.Show("Файл успешно объединён!", "Готово");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
