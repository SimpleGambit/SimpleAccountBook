using System;

namespace SimpleAccountBook.Models;

public class TransactionRecord
{
    public DateTime TransactionTime { get; set; }
    public string TransactionType { get; set; } = string.Empty; // 입금, 출금
    public decimal Amount { get; set; }
    public string Category { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;

    public bool IsIncome => string.Equals(TransactionType, "입금", StringComparison.OrdinalIgnoreCase);

    public decimal SignedAmount => IsIncome ? Amount : -Amount;
}