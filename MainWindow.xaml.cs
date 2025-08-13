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
		private ObservableCollection<ColumnMapping> columnMappings;
		private ObservableCollection<ReplacementRule> replacementRules = new ObservableCollection<ReplacementRule>();
        private bool ignoreCase = true;
        private bool ignoreWhitespace = true;
        private bool ignoreSymbols = false;
        // Deprecated single replacement inputs (replaced by rules list)
        private string? replacementFrom;
        private string? replacementTo;

		public MainWindow()
        {
            InitializeComponent();
			columnMappings = new ObservableCollection<ColumnMapping>();
            dgColumnMapping.ItemsSource = columnMappings;
			dgReplacementRules.ItemsSource = replacementRules;
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
                        
                        // Auto-map columns if both files are loaded
                        if (sourceData != null)
                        {
                            AutoMapColumns();
                        }
                        
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
                var headerRecord = csv.HeaderRecord ?? Array.Empty<string>();
                foreach (string header in headerRecord)
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
                    int columnCount = dataTable.Columns.Count;
                    for (int i = 0; i < columnCount; i++)
                    {
                        row[i] = csv.TryGetField(i, out string? field) ? field : string.Empty;
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

        private void AutoMapColumns()
        {
            if (sourceData == null || destinationData == null)
                return;

            // Get all destination column names
            var destinationColumns = destinationData.Columns.Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToList();

            // Auto-map columns with the same names
            foreach (var mapping in columnMappings)
            {
                if (mapping.SourceColumn != null)
                {
                    // Find exact match first (case-insensitive)
                    var exactMatch = destinationColumns.FirstOrDefault(dc => 
                        string.Equals(dc, mapping.SourceColumn, StringComparison.OrdinalIgnoreCase));
                    
                    if (exactMatch != null)
                    {
                        mapping.DestinationColumn = exactMatch;
                        continue;
                    }

                    // Find partial match (one contains the other)
                    var partialMatch = destinationColumns.FirstOrDefault(dc => 
                        dc.Contains(mapping.SourceColumn, StringComparison.OrdinalIgnoreCase) ||
                        mapping.SourceColumn.Contains(dc, StringComparison.OrdinalIgnoreCase));
                    
                    if (partialMatch != null)
                    {
                        mapping.DestinationColumn = partialMatch;
                        continue;
                    }

                    // Find similar names (common patterns)
                    var similarMatch = destinationColumns.FirstOrDefault(dc => 
                        IsSimilarColumnName(mapping.SourceColumn, dc));
                    
                    if (similarMatch != null)
                    {
                        mapping.DestinationColumn = similarMatch;
                    }
                }
            }

            // Refresh the DataGrid to show the changes
            dgColumnMapping.Items.Refresh();
        }

        private bool IsSimilarColumnName(string sourceName, string destinationName)
        {
            if (string.IsNullOrEmpty(sourceName) || string.IsNullOrEmpty(destinationName))
                return false;

            // Normalize names for comparison
            var normalizedSource = NormalizeColumnName(sourceName);
            var normalizedDest = NormalizeColumnName(destinationName);

            // Check if normalized names are similar
            if (string.Equals(normalizedSource, normalizedDest, StringComparison.OrdinalIgnoreCase))
                return true;

            // Check for common abbreviations and variations
            var sourceWords = normalizedSource.Split(new[] { ' ', '_', '-', '.' }, StringSplitOptions.RemoveEmptyEntries);
            var destWords = normalizedDest.Split(new[] { ' ', '_', '-', '.' }, StringSplitOptions.RemoveEmptyEntries);

            // If both have words, check for word overlap
            if (sourceWords.Length > 0 && destWords.Length > 0)
            {
                var commonWords = sourceWords.Intersect(destWords, StringComparer.OrdinalIgnoreCase);
                if (commonWords.Count() >= Math.Min(sourceWords.Length, destWords.Length) * 0.7) // 70% overlap
                {
                    return true;
                }
            }

            return false;
        }

        private string NormalizeColumnName(string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
                return string.Empty;

            // Remove common prefixes/suffixes and normalize
            var normalized = columnName
                .Replace("Column", "")
                .Replace("Col", "")
                .Replace("Field", "")
                .Replace("Fld", "")
                .Replace("_", " ")
                .Replace("-", " ")
                .Replace(".", " ");

            // Remove extra spaces and trim
            return string.Join(" ", normalized.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
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

        private void ChkIgnoreSymbols_Checked(object sender, RoutedEventArgs e)
        {
            ignoreSymbols = true;
        }

        private void ChkIgnoreSymbols_Unchecked(object sender, RoutedEventArgs e)
        {
            ignoreSymbols = false;
        }

        private void BtnAutoMap_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                AutoMapColumns();
                MessageBox.Show("Auto-mapping completed! Please review and adjust if needed.", "Auto-Map Complete",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during auto-mapping: {ex.Message}", "Auto-Map Error",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
				// Keep backward compatibility if any value is present later; rules list is primary
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
                    IgnoreWhitespace = ignoreWhitespace,
					IgnoreSymbols = ignoreSymbols,
					ReplacementFrom = replacementFrom,
					ReplacementTo = replacementTo,
					ReplacementRules = replacementRules.ToList()
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
			// Validate required inputs
			if (data.SourceData == null || data.DestinationData == null ||
				string.IsNullOrEmpty(data.SourceMainColumn) || string.IsNullOrEmpty(data.DestinationMainColumn))
			{
				return;
			}

			var sourceTable = data.SourceData;
			var destinationTable = data.DestinationData;
			var sourceMainColumn = data.SourceMainColumn;
			var destinationMainColumn = data.DestinationMainColumn;

			string sourceFileName = Path.GetFileNameWithoutExtension(data.SourceFilePath ?? string.Empty);
			string destinationFileName = Path.GetFileNameWithoutExtension(data.DestinationFilePath ?? string.Empty);
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
			foreach (DataColumn column in sourceTable.Columns)
            {
                headers.Add($"Source_{column.ColumnName}");
            }

            // Add all destination columns
			foreach (DataColumn column in destinationTable.Columns)
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

            // Create mapping dictionary
            var mappings = new Dictionary<string, string>();
			var columnMappingsLocal = data.ColumnMappings ?? new List<ColumnMapping>();
			foreach (var mapping in columnMappingsLocal)
            {
				if (!string.IsNullOrEmpty(mapping.SourceColumn) && !string.IsNullOrEmpty(mapping.DestinationColumn))
                {
					mappings[mapping.SourceColumn] = mapping.DestinationColumn;
                }
            }

            // Create destination lookup dictionary and detect duplicates
            var destinationLookup = new Dictionary<string, DataRow>();
            var duplicateKeys = new List<string>();
            
			foreach (DataRow row in destinationTable.Rows)
            {
				string key = ApplyReplacements(row[destinationMainColumn!]?.ToString() ?? string.Empty, data);
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
			foreach (DataRow sourceRow in sourceTable.Rows)
            {
				string sourceKey = ApplyReplacements(sourceRow[sourceMainColumn!]?.ToString() ?? string.Empty, data);
                var rowData = new List<object> { DateTime.Now };

                // Add source data
				foreach (DataColumn column in sourceTable.Columns)
                {
					rowData.Add(sourceRow[column] ?? "");
                }

                if (destinationLookup.ContainsKey(sourceKey))
                {
                    // Found - compare columns
                    DataRow destinationRow = destinationLookup[sourceKey];

                    // Add destination data
					foreach (DataColumn column in destinationTable.Columns)
                    {
						rowData.Add(destinationRow[column] ?? "");
                    }

                    // Check for differences with normalization options
                    var differences = new List<string>();
					foreach (var mapping in mappings)
                    {
						string sourceValue = sourceRow[mapping.Key]?.ToString() ?? "";
						string destValue = destinationRow[mapping.Value]?.ToString() ?? "";

                        if (!AreValuesEqual(sourceValue, destValue, data.IgnoreCase, data.IgnoreWhitespace, data.IgnoreSymbols))
                        {
                            differences.Add(mapping.Key);
                        }
                    }

                    rowData.Add(differences.Count > 0 ? "Half Match" : "Full Match");
                    rowData.Add(differences.Count > 0 ? $"Column not match: {string.Join(",", differences)}" : "All columns match");
                }
                else
                {
                    // Not found - add empty destination columns
                    foreach (DataColumn column in destinationTable.Columns)
                    {
                        rowData.Add("");
                    }

                    rowData.Add("Not found");
                    rowData.Add("Can not found the match column data!");
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

		private bool AreValuesEqual(string value1, string value2, bool ignoreCase, bool ignoreWhitespace, bool ignoreSymbols)
        {
            value1 ??= string.Empty;
            value2 ??= string.Empty;

            if (ignoreSymbols)
            {
                value1 = RemoveSymbols(value1, preserveWhitespace: true);
                value2 = RemoveSymbols(value2, preserveWhitespace: true);
            }

			if (ignoreWhitespace)
            {
				value1 = RemoveAllWhitespace(value1);
				value2 = RemoveAllWhitespace(value2);
            }

            if (ignoreCase)
            {
                return string.Equals(value1, value2, StringComparison.OrdinalIgnoreCase);
            }

            return string.Equals(value1, value2, StringComparison.Ordinal);
        }

		private static string RemoveSymbols(string input, bool preserveWhitespace)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            var filtered = new System.Text.StringBuilder(input.Length);
            foreach (var ch in input)
            {
                if (char.IsLetterOrDigit(ch))
                {
                    filtered.Append(ch);
                }
                else if (preserveWhitespace && char.IsWhiteSpace(ch))
                {
                    filtered.Append(ch);
                }
                // skip other symbols/punctuation
            }
			return filtered.ToString();
        }

		private static string RemoveAllWhitespace(string input)
		{
			if (string.IsNullOrEmpty(input)) return string.Empty;
			var filtered = new System.Text.StringBuilder(input.Length);
			foreach (var ch in input)
			{
				if (!char.IsWhiteSpace(ch))
				{
					filtered.Append(ch);
				}
			}
			return filtered.ToString();
		}

		private void BtnAddReplacement_Click(object sender, RoutedEventArgs e)
		{
			replacementRules.Add(new ReplacementRule { From = string.Empty, To = string.Empty });
		}

		private void BtnRemoveReplacement_Click(object sender, RoutedEventArgs e)
		{
			if (dgReplacementRules.SelectedItem is ReplacementRule rule)
			{
				replacementRules.Remove(rule);
			}
		}

		private void BtnMoveReplacementUp_Click(object sender, RoutedEventArgs e)
		{
			var index = dgReplacementRules.SelectedIndex;
			if (index > 0 && index < replacementRules.Count)
			{
				var item = replacementRules[index];
				replacementRules.RemoveAt(index);
				replacementRules.Insert(index - 1, item);
				dgReplacementRules.SelectedIndex = index - 1;
			}
		}

		private void BtnMoveReplacementDown_Click(object sender, RoutedEventArgs e)
		{
			var index = dgReplacementRules.SelectedIndex;
			if (index >= 0 && index < replacementRules.Count - 1)
			{
				var item = replacementRules[index];
				replacementRules.RemoveAt(index);
				replacementRules.Insert(index + 1, item);
				dgReplacementRules.SelectedIndex = index + 1;
			}
		}

		private static string ApplyReplacement(string input, string? replaceFrom, string? replaceTo)
		{
			if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(replaceFrom))
			{
				return input ?? string.Empty;
			}
			return input.Replace(replaceFrom, replaceTo ?? string.Empty, StringComparison.OrdinalIgnoreCase);
		}

		private static string ApplyReplacements(string input, ComparisonData data)
		{
			string result = input ?? string.Empty;
			if (data.ReplacementRules != null && data.ReplacementRules.Count > 0)
			{
				foreach (var rule in data.ReplacementRules)
				{
					if (!string.IsNullOrEmpty(rule?.From))
					{
						result = result.Replace(rule.From, rule?.To ?? string.Empty, StringComparison.OrdinalIgnoreCase);
					}
				}
			}
			else if (!string.IsNullOrEmpty(data.ReplacementFrom))
			{
				result = ApplyReplacement(result, data.ReplacementFrom, data.ReplacementTo);
			}
			return result;
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
        public bool IgnoreSymbols { get; set; }
			public string? ReplacementFrom { get; set; }
			public string? ReplacementTo { get; set; }
			public List<ReplacementRule>? ReplacementRules { get; set; }
    }

		public class ReplacementRule
		{
			public string? From { get; set; }
			public string? To { get; set; }
		}
}