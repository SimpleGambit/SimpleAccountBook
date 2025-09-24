using System;

namespace SimpleAccountBook.Services;

/// <summary>
/// 거래 타입 결정을 위한 헬퍼 클래스
/// </summary>
public static class TransactionTypeHelper
{
    /// <summary>
    /// 거래 타입을 결정합니다 (입금/출금)
    /// </summary>
    /// <param name="typeText">구분 컬럼의 텍스트</param>
    /// <param name="withdrawAmount">출금 금액 (있는 경우)</param>
    /// <param name="depositAmount">입금 금액 (있는 경우)</param>
    /// <returns>"입금" 또는 "출금"</returns>
    public static string DetermineType(string typeText, decimal withdrawAmount = 0, decimal depositAmount = 0)
    {
        // 1. 명시적으로 "입금" 또는 "출금"이 있으면 그대로 사용
        if (!string.IsNullOrWhiteSpace(typeText))
        {
            var normalized = typeText.Trim();
            if (normalized.Contains("입금") || normalized.Contains("입금액"))
                return "입금";
            if (normalized.Contains("출금") || normalized.Contains("출금액"))
                return "출금";
        }
        
        // 2. 입금/출금 금액으로 판단
        if (depositAmount > 0 && withdrawAmount == 0)
            return "입금";
        if (withdrawAmount > 0 && depositAmount == 0)
            return "출금";
        
        // 3. 기본값은 타입 텍스트 또는 "출금"
        return string.IsNullOrWhiteSpace(typeText) ? "출금" : typeText;
    }
}