using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using SimpleAccountBook.Services;

namespace SimpleAccountBook.Models;

public class TransactionRecord : INotifyPropertyChanged
{
    private DateTime _transactionTime;
    private decimal _amount;
    private string _transactionType = string.Empty;
    private string _category = string.Empty;
    private string _description = string.Empty;
    private bool _isExcludedFromTotal = false;
    
    public DateTime TransactionTime
    {
        get => _transactionTime;
        set
        {
            if (_transactionTime != value)
            {
                _transactionTime = value;
                OnPropertyChanged();
            }
        }
    }
    public string TransactionType
    {
        get => _transactionType;
        set
        {
            var normalized = TransactionTypeHelper.NormalizeTypeText(value);

            if (!string.Equals(_transactionType, normalized, StringComparison.Ordinal))
            {
                _transactionType = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsIncome));
                OnPropertyChanged(nameof(SignedAmount));
            }
        }
    }
    public decimal Amount
    {
        get => _amount;
        set
        {
            var normalized = Math.Abs(value);
            if (_amount != normalized)
            {
                _amount = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(SignedAmount));
            }
        }
    }
    public string Category
    {
        get => _category;
        set
        {
            if (!string.Equals(_category, value, StringComparison.Ordinal))
            {
                _category = value ?? string.Empty;
                OnPropertyChanged();
            }
        }
    }

    public string Description
    {
        get => _description;
        set
        {
            if (!string.Equals(_description, value, StringComparison.Ordinal))
            {
                _description = value ?? string.Empty;
                OnPropertyChanged();
            }
        }
    }

    public bool IsExcludedFromTotal
    {
        get => _isExcludedFromTotal;
        set
        {
            if (_isExcludedFromTotal != value)
            {
                _isExcludedFromTotal = value;
                OnPropertyChanged();
            }
        }
    }

    public bool IsIncome => string.Equals(TransactionType, "입금", StringComparison.OrdinalIgnoreCase);

    public decimal SignedAmount => IsIncome ? Amount : -Amount;
    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}