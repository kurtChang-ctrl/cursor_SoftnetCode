using System;

namespace Base.Enums
{
    /// <summary>
    /// 星期幾（可為 Flags 組合，例如 Weekdays = 週一～五）
    /// </summary>
    [Flags]
    public enum DaysOfWeek
    {
        None = 0,
        Sunday = 1 << 0,    // 1
        Monday = 1 << 1,    // 2
        Tuesday = 1 << 2,   // 4
        Wednesday = 1 << 3, // 8
        Thursday = 1 << 4,  // 16
        Friday = 1 << 5,    // 32
        Saturday = 1 << 6,  // 64
        Weekdays = Monday | Tuesday | Wednesday | Thursday | Friday,
        Weekend = Saturday | Sunday
    }
}
