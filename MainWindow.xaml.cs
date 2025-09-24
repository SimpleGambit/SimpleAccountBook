using Microsoft.Win32;
using SimpleAccountBook.Controls.Calendar;
using SimpleAccountBook.Models;
using SimpleAccountBook.Services;
using SimpleAccountBook.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;

namespace SimpleAccountBook;

public partial class MainWindow : Window
{
    private readonly AccountBookViewModel _viewModel;
    private readonly MemoService _memoService;
    private readonly AppStateService _appStateService;
    private bool _isRestoringState;

    public MainWindow()
    {
        InitializeComponent();
        _viewModel = new AccountBookViewModel(new ExcelImportService());
        _memoService = new MemoService();
        _appStateService = new AppStateService();
        DataContext = _viewModel;
        
        // ViewModel 속성 변경 이벤트 핸들러
        _viewModel.PropertyChanged += ViewModel_PropertyChanged;

        _memoService.MemosChanged += MemoService_MemosChanged;
        Loaded += MainWindow_Loaded;

        // 커스텀 달력 이벤트 핸들러
        AccountCalendar.SelectedDateChanged += AccountCalendar_SelectedDateChanged;
        AccountCalendar.DisplayDateChanged += AccountCalendar_DisplayDateChanged;
    }
    
    private void AccountCalendar_SelectedDateChanged(object sender, DateTime? selectedDate)
    {
        if (selectedDate.HasValue && _viewModel != null)
        {
            _viewModel.SelectedDate = selectedDate.Value;
            UpdateSelectedDateTransactionsWithMemo();
        }
    }
    
    private void AccountCalendar_DisplayDateChanged(object sender, DateTime displayDate)
    {
        if (_viewModel != null)
        {
            _viewModel.CurrentMonth = displayDate;
        }
    }

    private void ViewModel_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(AccountBookViewModel.DailySummaries) ||
            e.PropertyName == nameof(AccountBookViewModel.DailySummaryVersion))
        {
            System.Diagnostics.Debug.WriteLine("[MainWindow] DailySummaries 업데이트됨");
            UpdateCalendarData();
        }
        else if (e.PropertyName == nameof(AccountBookViewModel.SelectedDate))
        {
            if (_viewModel.SelectedDate != AccountCalendar.SelectedDate)
            {
                AccountCalendar.SelectedDate = _viewModel.SelectedDate;
                UpdateSelectedDateTransactionsWithMemo();
            }
        }
        else if (e.PropertyName == nameof(AccountBookViewModel.CurrentMonth))
        {
            if (_viewModel.CurrentMonth != AccountCalendar.DisplayDate)
            {
                AccountCalendar.DisplayDate = _viewModel.CurrentMonth;
            }
        }
        else if (e.PropertyName == nameof(AccountBookViewModel.SelectedDateTransactions))
        {
            UpdateSelectedDateTransactionsWithMemo();
        }
    }
    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        await LoadStateFromDiskAsync(false);
    }

    private void MemoService_MemosChanged(object? sender, EventArgs e)
    {
        if (_isRestoringState)
        {
            return;
        }

        _ = SaveCurrentStateAsync();
    }

    /// 선택된 날짜의 거래 목록을 메모 기능이 포함된 모델로 변환
    private void UpdateSelectedDateTransactionsWithMemo()
    {
        var transactionsWithMemo = new ObservableCollection<TransactionWithMemo>();
        
        foreach (var transaction in _viewModel.SelectedDateTransactions)
        {
            var withMemo = new TransactionWithMemo(transaction);
            
            // 저장된 메모가 있으면 로드
            var savedMemo = _memoService.GetMemo(transaction);
            if (savedMemo != null)
            {
                withMemo.Memo = savedMemo;
            }
            
            // 메모 변경 이벤트 구독
            withMemo.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(TransactionWithMemo.Memo) || e.PropertyName == nameof(TransactionWithMemo.IsExpanded))
                {
                    // 메모가 변경되면 자동 저장
                    if (withMemo.HasMemo)
                    {
                        _memoService.SaveMemo(withMemo.Transaction, withMemo.Memo);
                    }
                }
            };
            
            transactionsWithMemo.Add(withMemo);
        }
        
        // ViewModel의 SelectedDateTransactionsWithMemo 속성 업데이트
        _viewModel.SetSelectedDateTransactionsWithMemo(transactionsWithMemo);
    }
    
    private void UpdateCalendarData()
    {
        // 커스텀 달력에 데이터 전달
        Dispatcher.BeginInvoke(DispatcherPriority.Loaded, new Action(() =>
        {
            System.Diagnostics.Debug.WriteLine("[MainWindow] UpdateCalendarData 시작");
            System.Diagnostics.Debug.WriteLine($"[MainWindow] 일별 요약 수: {_viewModel.DailySummaries.Count}");
            
            // 커스텀 달력에 데이터 설정
            AccountCalendar.SetDailySummaries(_viewModel.DailySummaries);
            
            System.Diagnostics.Debug.WriteLine("[MainWindow] 달력 데이터 업데이트 완료");
        }));
    }
    
    /// 메모 상세 항목 추가
    private void AddMemoDetail_Click(object sender, RoutedEventArgs e)
    {
        var button = sender as Button;
        var transaction = button?.Tag as TransactionWithMemo;
        
        if (transaction != null)
        {
            var newDetail = new MemoDetail
            {
                ItemName = "",
                Amount = 0,
                Note = ""
            };
            
            transaction.Memo.Details.Add(newDetail);
            
            // 메모 저장
            _memoService.SaveMemo(transaction.Transaction, transaction.Memo);
        }
    }
    
    /// 메모 상세 항목 삭제
    private void RemoveMemoDetail_Click(object sender, RoutedEventArgs e)
    {
        var button = sender as Button;
        var detail = button?.Tag as MemoDetail;
        
        if (detail != null)
        {
            // 상위 TransactionWithMemo 찾기
            var itemsControl = FindParent<ItemsControl>(button);
            if (itemsControl != null)
            {
                var transaction = itemsControl.DataContext as TransactionWithMemo;
                if (transaction != null)
                {
                    transaction.Memo.Details.Remove(detail);
                    
                    // 메모 저장
                    _memoService.SaveMemo(transaction.Transaction, transaction.Memo);
                }
            }
        }
    }
    
    /// Visual Tree에서 부모 요소 찾기
    private T FindParent<T>(DependencyObject child) where T : DependencyObject
    {
        DependencyObject parentObject = VisualTreeHelper.GetParent(child);
        
        if (parentObject == null) return null;
        
        if (parentObject is T parent)
            return parent;
        else
            return FindParent<T>(parentObject);
    }

    private async void Window_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            return;
        }

        var files = (string[])e.Data.GetData(DataFormats.FileDrop);
        var excelFiles = files.Where(IsExcelFile).ToArray();
        if (excelFiles.Length == 0)
        {
            MessageBox.Show("엑셀(.xlsx) 파일만 지원합니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        await LoadFilesAsync(excelFiles);
    }

    private void Window_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            e.Effects = files.Any(IsExcelFile) ? DragDropEffects.Copy : DragDropEffects.None;
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }
    }

    private async void SelectFileButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls",
            Multiselect = true
        };

        if (dialog.ShowDialog() == true)
        {
            await LoadFilesAsync(dialog.FileNames);
        }
    }

    private async Task LoadFilesAsync(IEnumerable<string> filePaths)
    {
        bool hasChanges = false;

        foreach (var path in filePaths.Select(Path.GetFullPath))
        {
            if (!File.Exists(path))
            {
                MessageBox.Show($"파일을 찾을 수 없습니다: {path}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                continue;
            }

            await _viewModel.LoadFromFileAsync(path);
            hasChanges = true;
        }

        if (hasChanges)
        {
            RefreshDataViews();
            await SaveCurrentStateAsync();
        }
    }

    private void RefreshDataViews()
    {
        System.Diagnostics.Debug.WriteLine($"전체 거래 수: {_viewModel.Transactions.Count}");
        System.Diagnostics.Debug.WriteLine($"일별 요약 수: {_viewModel.DailySummaries.Count}");
        
        // 처음 5개의 일별 요약 출력
        foreach (var summary in _viewModel.DailySummaries.Take(5))
        {
            System.Diagnostics.Debug.WriteLine($"  {summary.Key}: 입금={summary.Value.IncomeAmount:#,0}, 출금={summary.Value.ExpenseAmount:#,0}");
            System.Diagnostics.Debug.WriteLine($"    IncomeText='{summary.Value.IncomeText}', ExpenseText='{summary.Value.ExpenseText}'");
        }
        
        // 달력 데이터 업데이트
        UpdateCalendarData();
        
        // 선택된 날짜의 거래 목록 업데이트
        UpdateSelectedDateTransactionsWithMemo();
    }

    private void OpenSettings_Click(object sender, RoutedEventArgs e)
    {
        var settingsWindow = new SettingsWindow
        {
            Owner = this,
            DataContext = _viewModel
        };

        settingsWindow.ShowDialog();
    }

    // 엑셀 데이터 제거
    private async void RemoveLoadedFile_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button || button.Tag is not string filePath)
        {
            return;
        }

        if (_viewModel.RemoveFile(filePath))
        {
            RefreshDataViews();
            await SaveCurrentStateAsync();
        }
        else
        {
            MessageBox.Show("파일을 제거하지 못했습니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static bool IsExcelFile(string path)
    {
        var extension = Path.GetExtension(path).ToLowerInvariant();
        return extension is ".xlsx" or ".xlsm" or ".xls";
    }
    private async void SaveStateButton_Click(object sender, RoutedEventArgs e)
    {
        await SaveCurrentStateAsync(true);
    }

    private async void LoadStateButton_Click(object sender, RoutedEventArgs e)
    {
        await LoadStateFromDiskAsync(true);
    }

    private async Task SaveCurrentStateAsync(bool showMessage = false)
    {
        if (_isRestoringState)
        {
            return;
        }

        try
        {
            var state = _viewModel.CreateAppState(_memoService.CreateSnapshot());
            await _appStateService.SaveAsync(state);

            if (showMessage)
            {
                MessageBox.Show("현재 상태를 저장했습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"상태를 저장하지 못했습니다: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task LoadStateFromDiskAsync(bool showMessage)
    {
        AppState? state;
        try
        {
            state = await _appStateService.LoadAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"상태를 불러오지 못했습니다: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (state is null)
        {
            if (showMessage)
            {
                MessageBox.Show("저장된 상태가 없습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            return;
        }

        try
        {
            _isRestoringState = true;
            _memoService.LoadFromSnapshot(state.Memos);
            _viewModel.RestoreState(state);
            RefreshDataViews();

            if (showMessage)
            {
                MessageBox.Show("저장된 상태를 불러왔습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        finally
        {
            _isRestoringState = false;
        }
    }
}