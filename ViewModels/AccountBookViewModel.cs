using SimpleAccountBook.Models;
using SimpleAccountBook.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace SimpleAccountBook.ViewModels;

public class AccountBookViewModel : INotifyPropertyChanged
{
    private readonly IExcelImportService _importService;
    private readonly Dictionary<string, List<TransactionRecord>> _fileTransactions = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<DateOnly, (decimal income, decimal expense)> _dailyTotals = new();
    private Dictionary<DateOnly, DailySummary> _dailySummaries = new();
    private ObservableCollection<TransactionRecord> _transactions = new();
    private readonly ObservableCollection<TransactionRecord> _selectedDateTransactions = new();
    private ObservableCollection<TransactionWithMemo> _selectedDateTransactionsWithMemo = new();
    private DateTime _currentMonth;
    private DateTime? _selectedDate;
    private string _statusMessage = "엑셀 파일을 드래그하여 불러옵니다.";
    private bool _isBusy;
    private decimal _monthlyIncome;
    private decimal _monthlyExpense;
    private decimal _assetBalance;
    private int _dailySummaryVersion;

    public event PropertyChangedEventHandler? PropertyChanged;

    public AccountBookViewModel(IExcelImportService importService)
    {
        _importService = importService;
        var today = DateTime.Today;
        _currentMonth = new DateTime(today.Year, today.Month, 1);
    }
    public ObservableCollection<LoadedFileInfo> LoadedFiles { get; } = new();

    public ObservableCollection<TransactionRecord> Transactions
    {
        get => _transactions;
        private set
        {
            if (!ReferenceEquals(_transactions, value))
            {
                _transactions = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasTransactions));
            }
        }
    }

    public ObservableCollection<TransactionRecord> SelectedDateTransactions => _selectedDateTransactions;

    // 선택된 날짜의 거래 목록 (메모 기능 포함)
    public ObservableCollection<TransactionWithMemo> SelectedDateTransactionsWithMemo
    {
        get => _selectedDateTransactionsWithMemo;
        set
        {
            _selectedDateTransactionsWithMemo = value;
            OnPropertyChanged();
        }
    }
    
    // 메모가 포함된 거래 목록 설정
    public void SetSelectedDateTransactionsWithMemo(ObservableCollection<TransactionWithMemo> transactions)
    {
        SelectedDateTransactionsWithMemo = transactions;
    }

    public DateTime CurrentMonth
    {
        get => _currentMonth;
        set
        {
            var normalized = new DateTime(value.Year, value.Month, 1);
            if (_currentMonth != normalized)
            {
                _currentMonth = normalized;
                OnPropertyChanged();
                UpdateMonthlyTotals();
            }
        }
    }

    public DateTime? SelectedDate
    {
        get => _selectedDate;
        set
        {
            if (_selectedDate != value)
            {
                _selectedDate = value;
                OnPropertyChanged();
                UpdateSelectedDateTransactions();
                if (value.HasValue)
                {
                    CurrentMonth = value.Value;
                }
            }
        }
    }

    public decimal MonthlyIncome
    {
        get => _monthlyIncome;
        private set
        {
            if (_monthlyIncome != value)
            {
                _monthlyIncome = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(MonthlyNet));
            }
        }
    }

    public decimal MonthlyExpense
    {
        get => _monthlyExpense;
        private set
        {
            if (_monthlyExpense != value)
            {
                _monthlyExpense = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(MonthlyNet));
            }
        }
    }

    public decimal MonthlyNet => MonthlyIncome - MonthlyExpense;

    public decimal AssetBalance
    {
        get => _assetBalance;
        private set
        {
            if (_assetBalance != value)
            {
                _assetBalance = value;
                OnPropertyChanged();
            }
        }
    }

    public string StatusMessage
    {
        get => _statusMessage;
        private set
        {
            if (_statusMessage != value)
            {
                _statusMessage = value;
                OnPropertyChanged();
            }
        }
    }

    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            if (_isBusy != value)
            {
                _isBusy = value;
                OnPropertyChanged();
            }
        }
    }

    public bool HasTransactions => Transactions.Any();

    public IReadOnlyDictionary<DateOnly, DailySummary> DailySummaries => _dailySummaries;

    public int DailySummaryVersion
    {
        get => _dailySummaryVersion;
        private set
        {
            if (_dailySummaryVersion != value)
            {
                _dailySummaryVersion = value;
                OnPropertyChanged();
            }
        }
    }

    public async Task LoadFromFileAsync(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return;
        }

        try
        {
            IsBusy = true;
            StatusMessage = "파일을 불러오는 중입니다...";
            var normalizedPath = Path.GetFullPath(filePath);
            var records = (await _importService.LoadAsync(normalizedPath)).OrderBy(r => r.TransactionTime).ToList();

            bool isUpdate = _fileTransactions.ContainsKey(normalizedPath);
            _fileTransactions[normalizedPath] = records;
            UpdateLoadedFile(normalizedPath, records.Count);

            RebuildTransactionsFromFiles();

            StatusMessage = isUpdate
                ? $"{Path.GetFileName(normalizedPath)} 새로고침 완료 - {records.Count}건 (총 {Transactions.Count}건)"
                : $"{Path.GetFileName(normalizedPath)} 불러오기 완료 - {records.Count}건 (총 {Transactions.Count}건)";
        }
        catch (Exception ex)
        {
            StatusMessage = $"파일을 불러오지 못했습니다: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }
    public bool RemoveFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return false;
        }

        var normalizedPath = Path.GetFullPath(filePath);
        if (!_fileTransactions.Remove(normalizedPath))
        {
            return false;
        }

        var existing = LoadedFiles.FirstOrDefault(f =>
            string.Equals(f.FilePath, normalizedPath, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            LoadedFiles.Remove(existing);
        }

        RebuildTransactionsFromFiles();

        StatusMessage = LoadedFiles.Any()
            ? $"{Path.GetFileName(normalizedPath)} 제거 완료 (총 {Transactions.Count}건)"
            : "모든 파일이 제거되었습니다. 새로운 엑셀 파일을 불러오세요.";

        return true;
    }

    public DailySummary GetDailySummary(DateTime date)
    {
        var key = DateOnly.FromDateTime(date);
        if (_dailySummaries.TryGetValue(key, out var summary))
        {
            return summary;
        }

        return DailySummary.CreateEmpty(date);
    }

    public AppState CreateAppState(Dictionary<string, MemoSnapshot>? memoSnapshots)
    {
        var state = new AppState
        {
            CurrentMonth = CurrentMonth,
            SelectedDate = SelectedDate
        };

        foreach (var (filePath, records) in _fileTransactions)
        {
            var normalizedPath = Path.GetFullPath(filePath);
            var loadedInfo = LoadedFiles.FirstOrDefault(f =>
                string.Equals(f.FilePath, normalizedPath, StringComparison.OrdinalIgnoreCase));

            var clonedRecords = records.Select(r => new TransactionRecord
            {
                TransactionTime = r.TransactionTime,
                TransactionType = r.TransactionType,
                Amount = r.Amount,
                Category = r.Category,
                Description = r.Description
            }).ToList();

            state.Files.Add(new AppStateFile
            {
                FilePath = normalizedPath,
                LoadedAt = loadedInfo?.LoadedAt ?? DateTime.Now,
                Transactions = clonedRecords
            });
        }

        state.Memos = memoSnapshots is null
            ? new Dictionary<string, MemoSnapshot>()
            : memoSnapshots.ToDictionary(
                kvp => kvp.Key,
                kvp => new MemoSnapshot
                {
                    LastModified = kvp.Value.LastModified,
                    Details = kvp.Value.Details
                        .Select(detail => new MemoDetailSnapshot
                        {
                            ItemName = detail.ItemName,
                            Amount = detail.Amount,
                            Note = detail.Note
                        }).ToList()
                },
                StringComparer.OrdinalIgnoreCase);

        return state;
    }

    public void RestoreState(AppState? state)
    {
        if (state is null)
        {
            return;
        }

        _fileTransactions.Clear();
        LoadedFiles.Clear();

        if (state.Files is not null)
        {
            foreach (var file in state.Files)
            {
                if (string.IsNullOrWhiteSpace(file.FilePath))
                {
                    continue;
                }

                var normalizedPath = Path.GetFullPath(file.FilePath);
                var orderedRecords = (file.Transactions ?? new List<TransactionRecord>())
                    .Select(r => new TransactionRecord
                    {
                        TransactionTime = r.TransactionTime,
                        TransactionType = r.TransactionType,
                        Amount = r.Amount,
                        Category = r.Category,
                        Description = r.Description
                    })
                    .OrderBy(r => r.TransactionTime)
                    .ToList();

                _fileTransactions[normalizedPath] = orderedRecords;
                var loadedAt = file.LoadedAt == default ? DateTime.Now : file.LoadedAt;
                LoadedFiles.Add(new LoadedFileInfo(normalizedPath, orderedRecords.Count, loadedAt));
            }
        }

        RebuildTransactionsFromFiles();

        if (state.CurrentMonth != default)
        {
            CurrentMonth = state.CurrentMonth;
        }

        SelectedDate = state.SelectedDate;

        StatusMessage = state.Files?.Any() == true
            ? "저장된 상태를 불러왔습니다."
            : "저장된 상태가 비어 있습니다.";
    }

    private void RebuildDailyTotals()
    {
        var newDailyTotals = new Dictionary<DateOnly, (decimal income, decimal expense)>();
        var newDailySummaries = new Dictionary<DateOnly, DailySummary>();
        foreach (var record in Transactions)
        {
            var date = DateOnly.FromDateTime(record.TransactionTime);
            var totals = newDailyTotals.TryGetValue(date, out var value)
                ? value
                : (income: 0m, expense: 0m);
            if (record.IsIncome)
            {
                totals.income += record.Amount;
            }
            else
            {
                totals.expense += record.Amount;
            }

            newDailyTotals[date] = totals;
        }

        foreach (var (date, totals) in newDailyTotals)
        {
            newDailySummaries[date] = DailySummary.FromTotals(date, totals.income, totals.expense);
        }

        _dailyTotals = newDailyTotals;
        _dailySummaries = newDailySummaries;
        OnPropertyChanged(nameof(DailySummaries));
        DailySummaryVersion++;

        Debug.WriteLine($"[RebuildDailyTotals] days={newDailySummaries.Count}");
        foreach (var kv in newDailyTotals.Take(5))
        {
            Debug.WriteLine($"  {kv.Key}: income={kv.Value.income}, expense={kv.Value.expense}");
        }
    }

    private void UpdateMonthlyTotals()
    {
        var start = new DateTime(CurrentMonth.Year, CurrentMonth.Month, 1);
        var end = start.AddMonths(1);

        MonthlyIncome = Transactions
            .Where(t => t.TransactionTime >= start && t.TransactionTime < end && t.IsIncome)
            .Sum(t => t.Amount);
        MonthlyExpense = Transactions
            .Where(t => t.TransactionTime >= start && t.TransactionTime < end && !t.IsIncome)
            .Sum(t => t.Amount);
    }

    private void UpdateAssetBalance()
    {
        AssetBalance = Transactions.Sum(t => t.SignedAmount);
    }

    private void UpdateSelectedDateTransactions()
    {
        _selectedDateTransactions.Clear();
        if (SelectedDate is not DateTime date)
        {
            return;
        }

        var key = DateOnly.FromDateTime(date);
        foreach (var record in Transactions
                     .Where(t => DateOnly.FromDateTime(t.TransactionTime) == key)
                     .OrderBy(t => t.TransactionTime))
        {
            _selectedDateTransactions.Add(record);
        }
    }
    private void RebuildTransactionsFromFiles()
    {
        var combined = _fileTransactions.Values
            .SelectMany(list => list)
            .OrderBy(t => t.TransactionTime)
            .ToList();

        Transactions = new ObservableCollection<TransactionRecord>(combined);

        RebuildDailyTotals();
        UpdateMonthlyTotals();
        UpdateAssetBalance();

        if (!Transactions.Any())
        {
            SelectedDate = null;
            UpdateSelectedDateTransactions();
            return;
        }

        if (SelectedDate is not DateTime selectedDate ||
            !Transactions.Any(t => DateOnly.FromDateTime(t.TransactionTime) == DateOnly.FromDateTime(selectedDate)))
        {
            SelectedDate = Transactions.Last().TransactionTime.Date;
        }
        else
        {
            UpdateSelectedDateTransactions();
        }
    }

    private void UpdateLoadedFile(string filePath, int recordCount)
    {
        var existing = LoadedFiles.FirstOrDefault(f =>
            string.Equals(f.FilePath, filePath, StringComparison.OrdinalIgnoreCase));

        if (existing is null)
        {
            LoadedFiles.Add(new LoadedFileInfo(filePath, recordCount, DateTime.Now));
        }
        else
        {
            existing.Update(recordCount, DateTime.Now);
        }
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}