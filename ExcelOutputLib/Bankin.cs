namespace ExcelOutputLib
{
    /// <summary>
    /// 板金クラス
    /// </summary>
    class Bankin
    {
        /// <summary>
        /// 種別（盤名・架台有無）
        /// </summary>
        public string Type { get; set; }
        /// <summary>
        /// 数量
        /// </summary>
        public int Quantity { get; set; }
        /// <summary>
        /// 単価
        /// </summary>
        public int UnitPrice { get; set; }
    }
}
