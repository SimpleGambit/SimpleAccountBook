using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace SimpleAccountBook.Controls
{
    public class CustomCalendarControl : UserControl
    {
        private readonly Grid _calendarGrid;
        private DateTime _displayDate = DateTime.Today;
        private Dictionary<DateOnly, (decimal income, decimal expense)> _dailyData = new Dictionary<DateOnly, (decimal income, decimal expense)>();

        public CustomCalendarControl()
        {
            _calendarGrid = new Grid();
            InitializeCalendar();
        }
        
        public DateTime DisplayDate
        {
            get => _displayDate;
            set
            {
                _displayDate = new DateTime(value.Year, value.Month, 1);
                UpdateCalendar();
            }
        }
        
        public void SetDailyData(Dictionary<DateOnly, (decimal income, decimal expense)> data)
        {
            _dailyData = data;
            UpdateCalendar();
        }
        
        private void InitializeCalendar()
        {
            
            // 8개 열 정의 (일-토 + 주간합계)
            for (int i = 0; i < 8; i++)
            {
                _calendarGrid.ColumnDefinitions.Add(new ColumnDefinition 
                { 
                    Width = i < 7 ? new GridLength(1, GridUnitType.Star) : new GridLength(1.2, GridUnitType.Star)
                });
            }
            
            // 7개 행 정의 (헤더 + 최대 6주)
            for (int i = 0; i < 7; i++)
            {
                _calendarGrid.RowDefinitions.Add(new RowDefinition 
                { 
                    Height = i == 0 ? GridLength.Auto : new GridLength(1, GridUnitType.Star)
                });
            }
            
            // 헤더 추가
            CreateHeaders();
            
            Content = _calendarGrid;
        }
        
        private void CreateHeaders()
        {
            string[] headers = { "일", "월", "화", "수", "목", "금", "토", "주간합계" };
            
            for (int i = 0; i < headers.Length; i++)
            {
                var headerBorder = new Border
                {
                    Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
                    BorderBrush = Brushes.LightGray,
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(5)
                };
                
                var headerText = new TextBlock
                {
                    Text = headers[i],
                    HorizontalAlignment = HorizontalAlignment.Center,
                    FontWeight = i == 7 ? FontWeights.Bold : FontWeights.Normal,
                    FontSize = i == 7 ? 11 : 12
                };
                
                headerBorder.Child = headerText;
                Grid.SetColumn(headerBorder, i);
                Grid.SetRow(headerBorder, 0);
                _calendarGrid.Children.Add(headerBorder);
            }
        }
        
        public void UpdateCalendar()
        {
            // 헤더 제외한 모든 셀 제거
            for (int i = _calendarGrid.Children.Count - 1; i >= 0; i--)
            {
                var element = _calendarGrid.Children[i];
                if (Grid.GetRow(element) > 0)
                {
                    _calendarGrid.Children.RemoveAt(i);
                }
            }
            
            var firstDay = new DateTime(_displayDate.Year, _displayDate.Month, 1);
            var lastDay = firstDay.AddMonths(1).AddDays(-1);
            var startDayOfWeek = (int)firstDay.DayOfWeek;
            
            int currentDay = 1;
            int totalDays = DateTime.DaysInMonth(_displayDate.Year, _displayDate.Month);
            
            // 주별 데이터 저장
            var weeklyTotals = new List<(decimal income, decimal expense)>();
            for (int i = 0; i < 6; i++)
                weeklyTotals.Add((0, 0));
            
            // 6주 × 7일 그리드 채우기
            for (int week = 0; week < 6; week++)
            {
                decimal weekIncome = 0;
                decimal weekExpense = 0;
                
                for (int dayOfWeek = 0; dayOfWeek < 7; dayOfWeek++)
                {
                    var cellBorder = new Border
                    {
                        BorderBrush = Brushes.LightGray,
                        BorderThickness = new Thickness(1),
                        Background = Brushes.White,
                        MinHeight = 60
                    };
                    
                    var cellPanel = new StackPanel
                    {
                        Margin = new Thickness(3)
                    };
                    
                    // 첫 주의 시작 요일 이전이거나 마지막 날 이후
                    if ((week == 0 && dayOfWeek < startDayOfWeek) || currentDay > totalDays)
                    {
                        // 이전/다음 달 날짜 표시 (회색)
                        cellBorder.Opacity = 0.5;
                    }
                    else
                    {
                        // 현재 월의 날짜
                        var dayNumber = new TextBlock
                        {
                            Text = currentDay.ToString(),
                            FontWeight = FontWeights.Bold,
                            HorizontalAlignment = HorizontalAlignment.Center
                        };
                        
                        // 오늘 날짜 강조
                        if (_displayDate.Year == DateTime.Today.Year &&
                            _displayDate.Month == DateTime.Today.Month &&
                            currentDay == DateTime.Today.Day)
                        {
                            cellBorder.BorderBrush = Brushes.Orange;
                            cellBorder.BorderThickness = new Thickness(2);
                            dayNumber.Foreground = Brushes.OrangeRed;
                        }
                        
                        cellPanel.Children.Add(dayNumber);
                        
                        // 입금/출금 데이터 표시
                        if (_dailyData != null)
                        {
                            var dateKey = new DateOnly(_displayDate.Year, _displayDate.Month, currentDay);
                            if (_dailyData.TryGetValue(dateKey, out var dayData))
                            {
                                weekIncome += dayData.income;
                                weekExpense += dayData.expense;
                                
                                if (dayData.income > 0)
                                {
                                    cellPanel.Children.Add(new TextBlock
                                    {
                                        Text = $"{dayData.income:#,0}",
                                        Foreground = Brushes.Green,
                                        FontSize = 10,
                                        HorizontalAlignment = HorizontalAlignment.Center
                                    });
                                }
                                
                                if (dayData.expense > 0)
                                {
                                    cellPanel.Children.Add(new TextBlock
                                    {
                                        Text = $"{dayData.expense:#,0}",
                                        Foreground = Brushes.Red,
                                        FontSize = 10,
                                        HorizontalAlignment = HorizontalAlignment.Center
                                    });
                                }
                            }
                        }
                        
                        currentDay++;
                    }
                    
                    cellBorder.Child = cellPanel;
                    Grid.SetColumn(cellBorder, dayOfWeek);
                    Grid.SetRow(cellBorder, week + 1);
                    _calendarGrid.Children.Add(cellBorder);
                }
                
                // 주간 합계 칸 추가
                CreateWeeklySummaryCell(week + 1, weekIncome, weekExpense);
                
                // 마지막 날 이후면 종료
                if (currentDay > totalDays)
                    break;
            }
        }
        
        private void CreateWeeklySummaryCell(int row, decimal income, decimal expense)
        {
            var summaryBorder = new Border
            {
                BorderBrush = Brushes.Gray,
                BorderThickness = new Thickness(1),
                Background = new SolidColorBrush(Color.FromRgb(245, 245, 245)),
                MinHeight = 60
            };
            
            var summaryPanel = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(3)
            };
            
            if (income > 0 || expense > 0)
            {
                // 수입
                if (income > 0)
                {
                    summaryPanel.Children.Add(new TextBlock
                    {
                        Text = $"{income:#,0}",
                        Foreground = Brushes.Green,
                        FontSize = 10,
                        HorizontalAlignment = HorizontalAlignment.Center
                    });
                }
                
                // 지출
                if (expense > 0)
                {
                    summaryPanel.Children.Add(new TextBlock
                    {
                        Text = $"{expense:#,0}",
                        Foreground = Brushes.Red,
                        FontSize = 10,
                        HorizontalAlignment = HorizontalAlignment.Center
                    });
                }
                
                // 구분선
                summaryPanel.Children.Add(new Separator
                {
                    Margin = new Thickness(0, 2, 0, 2)
                });
                
                // 합계
                var total = income - expense;
                summaryPanel.Children.Add(new TextBlock
                {
                    Text = total >= 0 ? $"+{total:#,0}" : $"{total:#,0}",
                    Foreground = total >= 0 ? Brushes.Green : Brushes.Red,
                    FontSize = 10,
                    FontWeight = FontWeights.Bold,
                    HorizontalAlignment = HorizontalAlignment.Center
                });
            }
            else
            {
                // 데이터 없음
                summaryPanel.Children.Add(new TextBlock
                {
                    Text = "-",
                    Foreground = Brushes.Gray,
                    FontSize = 10,
                    HorizontalAlignment = HorizontalAlignment.Center
                });
            }
            
            summaryBorder.Child = summaryPanel;
            Grid.SetColumn(summaryBorder, 7);
            Grid.SetRow(summaryBorder, row);
            _calendarGrid.Children.Add(summaryBorder);
        }
        
        // 이전/다음 월 이동 메서드
        public void GoToPreviousMonth()
        {
            DisplayDate = _displayDate.AddMonths(-1);
        }
        
        public void GoToNextMonth()
        {
            DisplayDate = _displayDate.AddMonths(1);
        }
    }
}