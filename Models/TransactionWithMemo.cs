using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace SimpleAccountBook.Models
{
    /// <summary>
    /// UI에서 사용할 메모 기능이 포함된 거래 모델
    /// </summary>
    public class TransactionWithMemo : INotifyPropertyChanged
    {
        private bool _isExpanded;
        private TransactionMemo _memo;
        
        public TransactionRecord Transaction { get; set; }
        public TransactionMemo Memo 
        { 
            get => _memo ??= new TransactionMemo();
            set
            {
                _memo = value;
                OnPropertyChanged();
            }
        }
        
        // UI에서 펼침/접기 상태
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                _isExpanded = value;
                OnPropertyChanged();
            }
        }
        
        // 상세 내역이 있는지 여부
        public bool HasMemo => Memo != null && Memo.Details.Count > 0;
        
        // 거래 정보를 직접 노출 (바인딩 편의를 위해)
        public DateTime TransactionTime => Transaction.TransactionTime;
        public string TransactionType => Transaction.TransactionType;
        public decimal Amount => Transaction.Amount;
        public string Category => Transaction.Category;
        public string Description => Transaction.Description;
        
        public TransactionWithMemo(TransactionRecord transaction)
        {
            Transaction = transaction;
        }
        
        public event PropertyChangedEventHandler PropertyChanged;
        
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}