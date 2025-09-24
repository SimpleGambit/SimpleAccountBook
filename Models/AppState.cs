using System;
using System.Collections.Generic;

namespace SimpleAccountBook.Models;

public class AppState
{
    public DateTime CurrentMonth { get; set; }
    public DateTime? SelectedDate { get; set; }
    public List<AppStateFile> Files { get; set; } = new List<AppStateFile>();
    public Dictionary<string, MemoSnapshot> Memos { get; set; } = new Dictionary<string, MemoSnapshot>();
}

public class AppStateFile
{
    public string FilePath { get; set; } = string.Empty;
    public DateTime LoadedAt { get; set; }
    public List<TransactionRecord> Transactions { get; set; } = new List<TransactionRecord>();
    public byte[]? ExcelContent { get; set; }
}

public class MemoSnapshot
{
    public DateTime LastModified { get; set; }
    public List<MemoDetailSnapshot> Details { get; set; } = new List<MemoDetailSnapshot>();
}

public class MemoDetailSnapshot
{
    public string ItemName { get; set; } = string.Empty;
    public decimal Amount { get; set; }
    public string Note { get; set; } = string.Empty;
}