using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using ClosedXML.Excel;
using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;

namespace Xlsx_and_Csv_Comparison_Dotnet_Program
{
    public partial class MainWindow : Window
    {
        private DataTable? sourceData;
        private DataTable? destinationData;
        private string? sourceFilePath;
        private string? destinationFilePath;
        private ObservableCollection<ColumnMapping>? columnMappings;
        private bool ignoreCase = true;
        private bool ignoreWhitespace = true;

        public MainWindow()
        {
            InitializeComponent();
            columnMappings = new ObservableCollection<ColumnMapping>();
            dgColumnMapping.ItemsSource = columnMappings;
        }

        private void BtnLoadSource_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                Title = "Select Source File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    sourceFilePath = openFileDialog.FileName;
                    txtSourceFilePath.Text = sourceFilePath;

                    if (IsValidFile(sourceFilePath))
                    {
                        sourceData = LoadFile(sourceFilePath);
                        DisplayFileInfo(sourceFilePath, sourceData, lblSourceFileInfo);
                        UpdateSourceColumnComboBox();
                        UpdateColumnMappings();
                        CheckRunButtonEnabled();
                    }
                    else
                    {
                        MessageBox.Show("Please select a valid CSV or XLSX file.", "Invalid File",
                                      MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading source file: {ex.Message}", "Error",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnLoadDestination_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                Title = "Select Destination File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    destinationFilePath = openFileDialog.FileName;
                    txtDestinationFilePath.Text = destinationFilePath;

                    if (IsValidFile(destinationFilePath))
                    {
                        destinationData = LoadFile(destinationFilePath);
                        DisplayFileInfo(destinationFilePath, destinationData, lblDestinationFileInfo);
                        UpdateDestinationColumnComboBox();
                        UpdateColumnMappings();
                        CheckRunButtonEnabled();
                    }
                    else
                    {
                        MessageBox.Show("Please select a valid CSV or XLSX file.", "Invalid File",
                                      MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading destination file: {ex.Message}", "Error",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private bool IsValidFile(string filePath)
        {
            string extension = Path.GetExtension(filePath).ToLower();
            return extension == ".csv" || extension == ".xlsx";
        }

        private DataTable LoadFile(string filePath)
        {
            string extension = Path.GetExtension(filePath).ToLower();

            if (extension == ".csv")
            {
                return LoadCsvFile(filePath);
            }
            else if (extension == ".xlsx")
            {
                return LoadXlsxFile(filePath);
            }

            throw new ArgumentException("Unsupported file format");
        }

        private DataTable LoadCsvFile(string filePath)
        {
            var dataTable = new DataTable();

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                MissingFieldFound = null
            };

            // Use StreamReader instead of loading entire file into memory
            using (var streamReader = new StreamReader(filePath))
            using (var csv = new CsvReader(streamReader, config))
            {
                csv.Read();
                csv.ReadHeader();

                // Clean and ensure unique column names
                var columnNames = new List<string>();
                foreach (string header in csv.HeaderRecord)
                {
                    string cleanHeader = CleanColumnName(header);
                    string uniqueHeader = MakeColumnNameUnique(cleanHeader, columnNames);
                    columnNames.Add(uniqueHeader);
                    dataTable.Columns.Add(uniqueHeader);
                }

                // Add rows
                while (csv.Read())
                {
                    var row = dataTable.NewRow();
                    for (int i = 0; i < csv.HeaderRecord.Length; i++)
                    {
                        row[i] = csv.GetField(i);
                    }
                    dataTable.Rows.Add(row);
                }
            }

            return dataTable;
        }

        private DataTable LoadXlsxFile(string filePath)
        {
            var dataTable = new DataTable();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var firstRowUsed = worksheet.FirstRowUsed();
                var lastRowUsed = worksheet.LastRowUsed();
                var firstColumnUsed = worksheet.FirstColumnUsed();
                var lastColumnUsed = worksheet.LastColumnUsed();

                // Clean and ensure unique column names
                var columnNames = new List<string>();
                for (int col = firstColumnUsed.ColumnNumber(); col <= lastColumnUsed.ColumnNumber(); col++)
                {
                    string columnName = firstRowUsed.Cell(col).GetString();
                    string cleanHeader = CleanColumnName(columnName);
                    string uniqueHeader = MakeColumnNameUnique(cleanHeader, columnNames);
                    columnNames.Add(uniqueHeader);
                    dataTable.Columns.Add(uniqueHeader);
                }

                // Add data rows
                for (int row = firstRowUsed.RowNumber() + 1; row <= lastRowUsed.RowNumber(); row++)
                {
                    var dataRow = dataTable.NewRow();
                    for (int col = firstColumnUsed.ColumnNumber(); col <= lastColumnUsed.ColumnNumber(); col++)
                    {
                        dataRow[col - firstColumnUsed.ColumnNumber()] = worksheet.Cell(row, col).GetString();
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }

            return dataTable;
        }

        private string CleanColumnName(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName))
                return "Column";
            
            // Remove leading/trailing whitespace and replace multiple spaces with single space
            return string.Join(" ", columnName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
        }

        private string MakeColumnNameUnique(string baseName, List<string> existingNames)
        {
            if (!existingNames.Contains(baseName))
                return baseName;

            int counter = 1;
            string newName = $"{baseName}_{counter}";
            while (existingNames.Contains(newName))
            {
                counter++;
                newName = $"{baseName}_{counter}";
            }
            return newName;
        }

        private void DisplayFileInfo(string filePath, DataTable data, TextBlock label)
        {
            var fileInfo = new FileInfo(filePath);
            string fileName = Path.GetFileName(filePath);
            long fileSize = fileInfo.Length;
            int itemCount = data.Rows.Count;

            label.Text = $"File: {fileName} | Items: {itemCount} | Size: {FormatFileSize(fileSize)}";
        }

        private string FormatFileSize(long bytes)
        {
            string[] suffixes = { "B", "KB", "MB", "GB", "TB" };
            int counter = 0;
            decimal number = bytes;

            while (Math.Round(number / 1024) >= 1)
            {
                number /= 1024;
                counter++;
            }

            return $"{number:n1}{suffixes[counter]}";
        }

        private void UpdateSourceColumnComboBox()
        {
            cmbSourceMainColumn.Items.Clear();
            if (sourceData != null)
            {
                foreach (DataColumn column in sourceData.Columns)
                {
                    cmbSourceMainColumn.Items.Add(column.ColumnName);
                }
            }
        }

        private void UpdateDestinationColumnComboBox()
        {
            cmbDestinationMainColumn.Items.Clear();
            var destinationColumns = new List<string> { "" }; // Empty option

            if (destinationData != null)
            {
                foreach (DataColumn column in destinationData.Columns)
                {
                    cmbDestinationMainColumn.Items.Add(column.ColumnName);
                    destinationColumns.Add(column.ColumnName);
                }
            }

            // Update column mapping destination options
            colDestinationMapping.ItemsSource = destinationColumns;
        }

        private void UpdateColumnMappings()
        {
            columnMappings.Clear();

            if (sourceData != null)
            {
                foreach (DataColumn column in sourceData.Columns)
                {
                    columnMappings.Add(new ColumnMapping { SourceColumn = column.ColumnName });
                }
            }
        }

        private void CmbSourceMainColumn_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CheckRunButtonEnabled();
        }

        private void CmbDestinationMainColumn_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CheckRunButtonEnabled();
        }

        private void CheckRunButtonEnabled()
        {
            btnRun.IsEnabled = sourceData != null &&
                              destinationData != null &&
                              cmbSourceMainColumn.SelectedItem != null &&
                              cmbDestinationMainColumn.SelectedItem != null;
        }

        private void ChkIgnoreCase_Checked(object sender, RoutedEventArgs e)
        {
            ignoreCase = true;
        }

        private void ChkIgnoreCase_Unchecked(object sender, RoutedEventArgs e)
        {
            ignoreCase = false;
        }

        private void ChkIgnoreWhitespace_Checked(object sender, RoutedEventArgs e)
        {
            ignoreWhitespace = true;
        }

        private void ChkIgnoreWhitespace_Unchecked(object sender, RoutedEventArgs e)
        {
            ignoreWhitespace = false;
        }

        private async void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Disable all controls during processing
                btnRun.IsEnabled = false;
                btnLoadSource.IsEnabled = false;
                btnLoadDestination.IsEnabled = false;
                progressBar.Visibility = Visibility.Visible;
                progressBar.IsIndeterminate = true;
                lblProgress.Text = "Processing...";

                // Capture all necessary data on UI thread before starting background work
                var comparisonData = new ComparisonData
                {
                    SourceData = sourceData?.Copy(),
                    DestinationData = destinationData?.Copy(),
                    SourceMainColumn = cmbSourceMainColumn.SelectedItem?.ToString(),
                    DestinationMainColumn = cmbDestinationMainColumn.SelectedItem?.ToString(),
                    ColumnMappings = columnMappings?.ToList(),
                    SourceFilePath = sourceFilePath,
                    DestinationFilePath = destinationFilePath,
                    IgnoreCase = ignoreCase,
                    IgnoreWhitespace = ignoreWhitespace
                };

                await Task.Run(() => GenerateComparisonReport(comparisonData));

                MessageBox.Show("Comparison report generated successfully!", "Success",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating report: {ex.Message}", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // Re-enable all controls
                btnRun.IsEnabled = true;
                btnLoadSource.IsEnabled = true;
                btnLoadDestination.IsEnabled = true;
                progressBar.Visibility = Visibility.Collapsed;
                progressBar.IsIndeterminate = false;
                lblProgress.Text = "";
            }
        }

        private void GenerateComparisonReport(ComparisonData data)
        {
            string sourceFileName = Path.GetFileNameWithoutExtension(data.SourceFilePath);
            string destinationFileName = Path.GetFileNameWithoutExtension(data.DestinationFilePath);
            string reportFileName = $"{sourceFileName}_Compare_{destinationFileName}_Report.xlsx";

            // Show save dialog for report location
            string reportPath = GetReportSavePath(reportFileName);
            if (string.IsNullOrEmpty(reportPath))
                return; // User cancelled

            using var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Comparison Report");

            // Create headers
            var headers = new List<string> { "TimeStamp" };

            // Add all source columns
            foreach (DataColumn column in data.SourceData.Columns)
            {
                headers.Add($"Source_{column.ColumnName}");
            }

            // Add all destination columns
            foreach (DataColumn column in data.DestinationData.Columns)
            {
                headers.Add($"Dest_{column.ColumnName}");
            }

            headers.Add("Status");
            headers.Add("Message");

            // Write headers
            for (int i = 0; i < headers.Count; i++)
            {
                worksheet.Cell(1, i + 1).Value = headers[i];
            }

            string sourceMainColumn = data.SourceMainColumn;
            string destinationMainColumn = data.DestinationMainColumn;

            // Create mapping dictionary
            var mappings = new Dictionary<string, string>();
            foreach (var mapping in data.ColumnMappings)
            {
                if (!string.IsNullOrEmpty(mapping.DestinationColumn))
                {
                    mappings[mapping.SourceColumn] = mapping.DestinationColumn;
                }
            }

            // Create destination lookup dictionary and detect duplicates
            var destinationLookup = new Dictionary<string, DataRow>();
            var duplicateKeys = new List<string>();
            
            foreach (DataRow row in data.DestinationData.Rows)
            {
                string key = row[destinationMainColumn]?.ToString() ?? "";
                if (!string.IsNullOrEmpty(key))
                {
                    if (destinationLookup.ContainsKey(key))
                    {
                        duplicateKeys.Add(key);
                    }
                    destinationLookup[key] = row; // Last occurrence wins
                }
            }

            // Add duplicate key warning to worksheet
            if (duplicateKeys.Any())
            {
                var warningWorksheet = workbook.AddWorksheet("Duplicate Keys Warning");
                warningWorksheet.Cell(1, 1).Value = "Warning: Duplicate keys found in destination file";
                warningWorksheet.Cell(2, 1).Value = "Duplicate Keys:";
                
                for (int i = 0; i < duplicateKeys.Count; i++)
                {
                    warningWorksheet.Cell(3 + i, 1).Value = duplicateKeys[i];
                }
                
                warningWorksheet.Columns().AdjustToContents();
            }

            int currentRow = 2;

            // Process each source row
            foreach (DataRow sourceRow in data.SourceData.Rows)
            {
                string sourceKey = sourceRow[sourceMainColumn]?.ToString() ?? "";
                var rowData = new List<object> { DateTime.Now };

                // Add source data
                foreach (DataColumn column in data.SourceData.Columns)
                {
                    rowData.Add(sourceRow[column] ?? "");
                }

                if (destinationLookup.ContainsKey(sourceKey))
                {
                    // Found - compare columns
                    DataRow destinationRow = destinationLookup[sourceKey];

                    // Add destination data
                    foreach (DataColumn column in data.DestinationData.Columns)
                    {
                        rowData.Add(destinationRow[column] ?? "");
                    }

                    // Check for differences with normalization options
                    var differences = new List<string>();
                    foreach (var mapping in mappings)
                    {
                        string sourceValue = sourceRow[mapping.Key]?.ToString() ?? "";
                        string destValue = destinationRow[mapping.Value]?.ToString() ?? "";

                        if (!AreValuesEqual(sourceValue, destValue, data.IgnoreCase, data.IgnoreWhitespace))
                        {
                            differences.Add(mapping.Key);
                        }
                    }

                    rowData.Add("Existed!");
                    rowData.Add(differences.Count > 0 ? $"Column not match: {string.Join(",", differences)}" : "All columns match");
                }
                else
                {
                    // Not found - add empty destination columns
                    foreach (DataColumn column in data.DestinationData.Columns)
                    {
                        rowData.Add("");
                    }

                    rowData.Add("Not found!");
                    rowData.Add("Not found!");
                }

                // Write row data
                for (int i = 0; i < rowData.Count; i++)
                {
                    worksheet.Cell(currentRow, i + 1).Value = rowData[i]?.ToString() ?? "";
                }
                currentRow++;
            }

            // Auto-fit columns
            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(reportPath);
        }

        private bool AreValuesEqual(string value1, string value2, bool ignoreCase, bool ignoreWhitespace)
        {
            if (ignoreWhitespace)
            {
                value1 = value1?.Trim() ?? "";
                value2 = value2?.Trim() ?? "";
            }

            if (ignoreCase)
            {
                return string.Equals(value1, value2, StringComparison.OrdinalIgnoreCase);
            }

            return string.Equals(value1, value2, StringComparison.Ordinal);
        }

        private string GetReportSavePath(string defaultFileName)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FileName = defaultFileName,
                Title = "Save Comparison Report"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                return saveFileDialog.FileName;
            }

            return string.Empty;
        }
    }

    public class ColumnMapping
    {
        public string? SourceColumn { get; set; }
        public string? DestinationColumn { get; set; }
    }

    public class ComparisonData
    {
        public DataTable? SourceData { get; set; }
        public DataTable? DestinationData { get; set; }
        public string? SourceMainColumn { get; set; }
        public string? DestinationMainColumn { get; set; }
        public List<ColumnMapping>? ColumnMappings { get; set; }
        public string? SourceFilePath { get; set; }
        public string? DestinationFilePath { get; set; }
        public bool IgnoreCase { get; set; }
        public bool IgnoreWhitespace { get; set; }
    }
}