using System;
using System.Collections.Generic;

namespace ExcelOutputLib
{
    public static class ExcelOutput
    {
        /// <summary>
        /// エクセル出力
        /// </summary>
        /// <param name="item">C++フォームアプリ
        /// 板金リストボックスのアイテム
        /// （ほんとはJasonが格好良い。）</param>
        public static bool Output(string[] list)
        {
            bool ret = false;
            //アイテムをシリアル化
            List<Bankin> bList = new List<Bankin>();
            Bankin bItem;
            foreach (string item in list)
            {
                string[] phraze = item.Split(' ');
                if (phraze.Length > 1)
                {
                    //アイテムに中身があれば取得
                    bItem = new Bankin();
                    bItem.Quantity = Convert.ToInt32(phraze[1].Substring(phraze[1].IndexOf(':') + 1));
                    bItem.Type = phraze[2] + phraze[3];
                    bItem.UnitPrice = Convert.ToInt32(phraze[8].Substring(phraze[8].IndexOf(':') + 1));
                    bList.Add(bItem);
                }
            }
            if (bList.Count > 0)
            {
                //エクセルに出力
                DataExport de = new DataExport();
                ret = de.OpenExcel(bList);
            }
            return ret;
        }
    }
}
