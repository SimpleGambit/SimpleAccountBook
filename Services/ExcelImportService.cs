using ClosedXML.Excel;
using ExcelDataReader;
using SimpleAccountBook.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleAccountBook.Services;

public class ExcelImportService : IExcelImportService
{
    private static readonly string[] DateHeaders = { "거래일시" };
    private static readonly string[] TypeHeaders = { "구분", "입출금구분" };
    private static readonly string[] AmountHeaders = { "거래금액", "금액" };
    private static readonly string[] WithdrawAmountHeaders = { "출금(원)" };
    private static readonly string[] DepositAmountHeaders = { "입금(원)" };
    private static readonly string[] CategoryHeaders = { "거래구분" };
    private static readonly string[] DescriptionHeaders = { "내용", "적요" };

    public Task<IList<TransactionRecord>> LoadAsync(string filePath)
    {
        return Task.Run(() => LoadInternal(filePath));
    }

    private IList<TransactionRecord> LoadInternal(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("파일을 찾을 수 없습니다.", filePath);
        }

        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension switch
        {
            ".xls" => LoadLegacyExcel(filePath),
            _ => LoadWithClosedXml(filePath)
        };
    }

    private static bool TryGetDateTime(object? rawValue, string text, out DateTime value)
    {
        switch (rawValue)
        {
            case DateTime dateTime:
                value = dateTime;
                return true;
            case double oaDate:
                value = DateTime.FromOADate(oaDate);
                return true;
        }

        text = text.Trim();
        return DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out value) ||
               DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out value);
    }

    private static bool TryGetDecimal(object? rawValue, string text, out decimal value)
    {
        switch (rawValue)
        {
            case decimal decimalValue:
                value = decimalValue;
                return true;
            case double doubleValue:
                value = Convert.ToDecimal(doubleValue);
                return true;
            case float floatValue:
                value = Convert.ToDecimal(floatValue);
                return true;
            case int intValue:
                value = intValue;
                return true;
            case long longValue:
                value = longValue;
                return true;
        }

        text = text.Replace(",", string.Empty).Replace("원", string.Empty).Trim();
        text = text.Replace("(", "-").Replace(")", string.Empty);
        return decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value) ||
               decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryGetNonZeroDecimal(object? rawValue, string text, out decimal value)
    {
        if (!TryGetDecimal(rawValue, text, out value))
        {
            return false;
        }

        return value != 0;
    }

    private IList<TransactionRecord> LoadWithClosedXml(string filePath)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheets.First();

        var (headerMap, headerRowNumber) = FindHeaderRow(worksheet);

        int dateColumn = FindColumn(headerMap, DateHeaders);
        int typeColumn = FindColumn(headerMap, TypeHeaders);
        bool hasAmountColumn = TryFindColumn(headerMap, AmountHeaders, out var amountColumn);
        bool hasWithdrawColumn = TryFindColumn(headerMap, WithdrawAmountHeaders, out var withdrawColumn);
        bool hasDepositColumn = TryFindColumn(headerMap, DepositAmountHeaders, out var depositColumn);

        if (!hasAmountColumn && !hasWithdrawColumn && !hasDepositColumn)
        {
            throw new InvalidDataException($"필수 컬럼을 찾을 수 없습니다: {string.Join(", ", AmountHeaders.Concat(WithdrawAmountHeaders).Concat(DepositAmountHeaders))}");
        }

        int categoryColumn = FindColumn(headerMap, CategoryHeaders);
        int descriptionColumn = FindColumn(headerMap, DescriptionHeaders);

        var records = new List<TransactionRecord>();
        foreach (var row in worksheet.RowsUsed().Where(r => r.RowNumber() > headerRowNumber))
        {
            var dateCell = row.Cell(dateColumn);
            if (dateCell.IsEmpty())
            {
                continue;
            }

            var dateText = dateCell.GetString();
            if (!TryGetDateTime(dateCell.Value, dateText, out var transactionTime))
            {
                continue;
            }

            var type = row.Cell(typeColumn).GetString().Trim();
            
            // 콜 바인딩
            var withdrawCell = hasWithdrawColumn ? row.Cell(withdrawColumn) : null;
            var depositCell = hasDepositColumn ? row.Cell(depositColumn) : null;
            
            if (!TryReadAmount(hasAmountColumn ? row.Cell(amountColumn) : null,
                      withdrawCell,
                      depositCell,
                      out var amount))
            {
                continue;
            }

            var category = row.Cell(categoryColumn).GetString().Trim();
            var description = row.Cell(descriptionColumn).GetString().Trim();

            // 입금/출금 구분 개선
            decimal withdrawAmt = 0, depositAmt = 0;
            if (hasWithdrawColumn && withdrawCell != null)
                TryGetDecimal(withdrawCell.Value, withdrawCell.GetString(), out withdrawAmt);
            if (hasDepositColumn && depositCell != null)
                TryGetDecimal(depositCell.Value, depositCell.GetString(), out depositAmt);
            
            var transactionType = TransactionTypeHelper.DetermineType(type, withdrawAmt, depositAmt);
            
            records.Add(new TransactionRecord
            {
                TransactionTime = transactionTime,
                TransactionType = transactionType,
                Amount = Math.Abs(amount),
                Category = category,
                Description = description
            });
        }

        return records.OrderBy(r => r.TransactionTime).ToList();
    }

    private IList<TransactionRecord> LoadLegacyExcel(string filePath)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration
            {
                UseHeaderRow = false
            }
        });

        if (dataSet.Tables.Count == 0)
        {
            return new List<TransactionRecord>();
        }

        var table = dataSet.Tables[0];
        if (table.Rows.Count == 0)
        {
            return new List<TransactionRecord>();
        }

        var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var headerRow = table.Rows[0];
        for (int column = 0; column < table.Columns.Count; column++)
        {
            var header = NormalizeHeader(Convert.ToString(headerRow[column]));
            if (string.IsNullOrEmpty(header) || headerMap.ContainsKey(header))
            {
                continue;
            }

            headerMap[header] = column + 1;
        }

        int dateColumn = FindColumn(headerMap, DateHeaders);
        int typeColumn = FindColumn(headerMap, TypeHeaders);
        bool hasAmountColumn = TryFindColumn(headerMap, AmountHeaders, out var amountColumn);
        bool hasWithdrawColumn = TryFindColumn(headerMap, WithdrawAmountHeaders, out var withdrawColumn);
        bool hasDepositColumn = TryFindColumn(headerMap, DepositAmountHeaders, out var depositColumn);

        if (!hasAmountColumn && !hasWithdrawColumn && !hasDepositColumn)
        {
            throw new InvalidDataException($"필수 컬럼을 찾을 수 없습니다: {string.Join(", ", AmountHeaders.Concat(WithdrawAmountHeaders).Concat(DepositAmountHeaders))}");
        }

        int categoryColumn = FindColumn(headerMap, CategoryHeaders);
        int descriptionColumn = FindColumn(headerMap, DescriptionHeaders);

        var records = new List<TransactionRecord>();
        for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
        {
            var row = table.Rows[rowIndex];
            var dateValue = row[dateColumn - 1];
            var dateText = Convert.ToString(dateValue) ?? string.Empty;
            if (IsNullOrEmpty(dateValue, dateText))
            {
                continue;
            }

            if (!TryGetDateTime(dateValue, dateText, out var transactionTime))
            {
                continue;
            }

            var type = Convert.ToString(row[typeColumn - 1])?.Trim() ?? string.Empty;
            if (!TryReadAmount(row, hasAmountColumn ? amountColumn : (int?)null,
           hasWithdrawColumn ? withdrawColumn : (int?)null,
           hasDepositColumn ? depositColumn : (int?)null,
           out var amount))
            {
                continue;
            }

            var category = Convert.ToString(row[categoryColumn - 1])?.Trim() ?? string.Empty;
            var description = Convert.ToString(row[descriptionColumn - 1])?.Trim() ?? string.Empty;

            // 입금/출금 구분 개선
            decimal withdrawAmt = 0, depositAmt = 0;
            if (hasWithdrawColumn)
            {
                var withdrawValue = row[withdrawColumn - 1];
                TryGetDecimal(withdrawValue, Convert.ToString(withdrawValue) ?? string.Empty, out withdrawAmt);
            }
            if (hasDepositColumn)
            {
                var depositValue = row[depositColumn - 1];
                TryGetDecimal(depositValue, Convert.ToString(depositValue) ?? string.Empty, out depositAmt);
            }
            
            var transactionType = TransactionTypeHelper.DetermineType(type, withdrawAmt, depositAmt);
            
            records.Add(new TransactionRecord
            {
                TransactionTime = transactionTime,
                TransactionType = transactionType,
                Amount = Math.Abs(amount),
                Category = category,
                Description = description
            });
        }

        return records.OrderBy(r => r.TransactionTime).ToList();
    }

    private static bool IsNullOrEmpty(object? value, string text)
    {
        if (value is DBNull)
        {
            return true;
        }

        return string.IsNullOrWhiteSpace(text);
    }

    private static bool TryReadAmount(IXLCell? amountCell, IXLCell? withdrawCell, IXLCell? depositCell, out decimal amount)
    {
        amount = 0;

        if (amountCell != null && TryGetNonZeroDecimal(amountCell.Value, amountCell.GetString(), out amount))
        {
            return true;
        }

        if (withdrawCell != null && TryGetNonZeroDecimal(withdrawCell.Value, withdrawCell.GetString(), out amount))
        {
            return true;
        }

        if (depositCell != null && TryGetNonZeroDecimal(depositCell.Value, depositCell.GetString(), out amount))
        {
            return true;
        }

        return false;
    }

    private static bool TryReadAmount(DataRow row, int? amountColumn, int? withdrawColumn, int? depositColumn, out decimal amount)
    {
        amount = 0;

        if (amountColumn.HasValue)
        {
            var value = row[amountColumn.Value - 1];
            var text = Convert.ToString(value) ?? string.Empty;
            if (TryGetNonZeroDecimal(value, text, out amount))
            {
                return true;
            }
        }

        if (withdrawColumn.HasValue)
        {
            var value = row[withdrawColumn.Value - 1];
            var text = Convert.ToString(value) ?? string.Empty;
            if (TryGetNonZeroDecimal(value, text, out amount))
            {
                return true;
            }
        }

        if (depositColumn.HasValue)
        {
            var value = row[depositColumn.Value - 1];
            var text = Convert.ToString(value) ?? string.Empty;
            if (TryGetNonZeroDecimal(value, text, out amount))
            {
                return true;
            }
        }

        return false;
    }

    private static (Dictionary<string, int> HeaderMap, int HeaderRowNumber) FindHeaderRow(IXLWorksheet worksheet)
    {
        foreach (var row in worksheet.RowsUsed())
        {
            var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var cell in row.CellsUsed())
            {
                var normalized = NormalizeHeader(cell.GetString());
                if (string.IsNullOrEmpty(normalized) || headerMap.ContainsKey(normalized))
                {
                    continue;
                }

                headerMap[normalized] = cell.Address.ColumnNumber;
            }

            if (ContainsAllRequiredHeaders(headerMap))
            {
                return (headerMap, row.RowNumber());
            }
        }

        throw new InvalidDataException("필수 컬럼을 포함한 헤더 행을 찾을 수 없습니다.");
    }

    private static bool ContainsAllRequiredHeaders(Dictionary<string, int> headerMap)
    {
        return HasCandidate(headerMap, DateHeaders) &&
               HasCandidate(headerMap, TypeHeaders) &&
               (HasCandidate(headerMap, AmountHeaders) ||
                HasCandidate(headerMap, WithdrawAmountHeaders) ||
                HasCandidate(headerMap, DepositAmountHeaders)) &&
               HasCandidate(headerMap, CategoryHeaders) &&
               HasCandidate(headerMap, DescriptionHeaders);
    }

    private static bool HasCandidate(Dictionary<string, int> headerMap, IEnumerable<string> candidates)
    {
        foreach (var candidate in candidates)
        {
            var normalized = NormalizeHeader(candidate);
            if (headerMap.ContainsKey(normalized))
            {
                return true;
            }
        }

        return false;
    }

    private static string NormalizeHeader(string? header)
    {
        if (string.IsNullOrWhiteSpace(header))
        {
            return string.Empty;
        }

        var builder = new StringBuilder(header.Length);
        foreach (var ch in header)
        {
            if (!char.IsWhiteSpace(ch))
            {
                builder.Append(ch);
            }
        }

        return builder.ToString();
    }

    private static bool TryFindColumn(Dictionary<string, int> headers, IEnumerable<string> candidates, out int column)
    {
        foreach (var candidate in candidates)
        {
            var normalized = NormalizeHeader(candidate);
            if (headers.TryGetValue(normalized, out column))
            {
                return true;
            }
        }

        column = default;
        return false;
    }

    private static int FindColumn(Dictionary<string, int> headers, IEnumerable<string> candidates)
    {
        if (TryFindColumn(headers, candidates, out var column))
        {
            return column;
        }

        throw new InvalidDataException($"필수 컬럼을 찾을 수 없습니다: {string.Join(", ", candidates)}");
    }
}