using System;
using System.Collections.Generic;

namespace ExcelAddIn1.Core
{
    /// <summary>
    /// Human-readable fallbacks used when a workbook has not yet created a
    /// PriceProfile. A saved workbook price/profile always has priority.
    /// </summary>
    public static class EstimateV2ResourceNames
    {
        private static readonly IReadOnlyDictionary<string, string> Names =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "LAB-QNCN-5", "Nhân công thợ bậc 5/10" },
                { "LAB-QNCN-7", "Nhân công thợ bậc 7/10" },
                { "LAB-QNCN-8", "Nhân công thợ bậc 8/10" },

                { "MAT-CONCRETE-STAKE", "Cọc bê tông cốt thép (12x12x120)cm" },
                { "MAT-WOOD-STAKE", "Cọc gỗ d = 3cm, L= 50 cm" },
                { "MAT-RED-FLAG", "Cờ đỏ đuôi nheo" },
                { "MAT-RED-FLAG-LARGE", "Cờ đỏ to (40 x 60) cm" },
                { "MAT-ROPE-10MM", "Dây thừng 10 mm" },
                { "MAT-BAMBOO-STAKE", "Cọc tre d=8, L=200 cm" },
                { "MAT-WOOD-BOARD", "Ván gỗ dày 3 cm" },
                { "MAT-NAIL-10CM", "Đinh 10 cm" },
                { "MAT-NAIL-TWO-PRONG", "Đinh 2 mỏ" },
                { "MAT-WIRE-2MM", "Dây thép buộc 2 ly" },
                { "MAT-BAMBOO-POLE", "Sào tre d=7, L=500 cm" },
                { "MAT-ANCHOR-50KG", "Mỏ neo loại 50 kg" },
                { "MAT-ANCHOR-20KG", "Mỏ neo loại 20 kg" },
                { "MAT-SPECIAL-ANCHOR", "Mỏ neo đặc biệt loại 20 kg" },
                { "MAT-LARGE-FLOAT", "Phao lớn 1m3" },
                { "MAT-PLASTIC-FLOAT", "Phao nhựa tròn phi 30 cm" },
                { "MAT-SMALL-FLOAT", "Phao nhỏ phi 12 cm" },
                { "MAT-ROPE-12MM", "Dây Nilon phi 12mm" },
                { "MAT-ROPE-18MM", "Dây Nilon phi 18mm" },
                { "MAT-NYLON-ROPE-14MM", "Dây Nilon phi 14mm" },
                { "MAT-STEEL-EXCAVATION-FRAME", "Khung gia công bằng tôn 3mm và sắt góc (45x45)mm" },
                { "MAT-BARBED-WIRE", "Dây thép gai" },
                { "MAT-CLOTH", "Vải" },
                { "MAT-ELECTRIC-DETONATOR", "Kíp điện" },
                { "MAT-ELECTRIC-WIRE", "Dây điện" },
                { "MAT-FLOAT-40L", "Phao 40 lít" },
                { "MAT-GASOLINE", "Xăng" },
                { "MAT-DIESEL", "Dầu Diesel" },
                { "MAT-GASOLINE-OR-DIESEL", "Xăng hoặc dầu" },
                { "MAT-SIGN", "Biển báo" },
                { "MAT-SLOW-FUSE", "Dây cháy chậm" },
                { "MAT-SMOOTH-PLASTIC-TUBE", "Ống nhựa trơn" },
                { "MAT-TNT", "Thuốc nổ TNT" },
                { "MAT-TWINE", "Dây buộc" },
                { "MAT-WATERPROOF-DETONATING-CORD", "Dây nổ chịu nước" },
                { "MAT-WHITE-FLAG", "Cờ trắng" },
                { "MAT-WOOD-STRIP", "Nẹp gỗ" }
            };

        public static string Get(string code)
        {
            string key = (code ?? string.Empty).Trim();
            string value;
            return Names.TryGetValue(key, out value) ? value : key;
        }
    }
}
