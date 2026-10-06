using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn1.Funtion
{
    internal class Tools
    {
        private static readonly string[] CharVND = new string[10];
       
        private const string DonViTien = "đồng";
        private const string DonViLe = "xu";

        public static string Vnd(decimal numCurrency, string Tiento ="", string Hauto = "")
        {
            string bangChu = "";
            int soLe, soDoi, nghinTy, ty, trieu, nghin, dong;
            int tram, muoi, donVi;
            string phanChan, ten;

            numCurrency = Math.Round(numCurrency, 0);

            if (numCurrency == 0)
            {
                return "Không " + DonViTien;
            }

            if (numCurrency > 922337203685477m) // Số lớn nhất của loại CURRENCY
            {
                return "Không đổi được số lớn hơn 922.337.203.685.477";
            }

            CharVND[1] = "một";
            CharVND[2] = "hai";
            CharVND[3] = "ba";
            CharVND[4] = "bốn";
            CharVND[5] = "năm";
            CharVND[6] = "sáu";
            CharVND[7] = "bảy";
            CharVND[8] = "tám";
            CharVND[9] = "chín";


            soLe = (int)((numCurrency - Math.Floor(numCurrency)) * 100); // 2 ký số
            phanChan = Math.Floor(numCurrency).ToString().PadLeft(15);
            // Khai báo biến cho các giá trị chuyển đổi
           

            // Sử dụng int.TryParse để chuyển đổi và gán giá trị cho các biến
            // Chuyển đổi 3 ký tự đầu tiên
            if (!int.TryParse(phanChan.Substring(0, 3), out nghinTy))
            {
                nghinTy = 0;
            }

            // Chuyển đổi 3 ký tự tiếp theo
            if (!int.TryParse(phanChan.Substring(3, 3), out ty))
            {
                ty = 0;
            }

            // Chuyển đổi 3 ký tự tiếp theo
            if (!int.TryParse(phanChan.Substring(6, 3), out trieu))
            {
                trieu = 0;
            }

            // Chuyển đổi 3 ký tự tiếp theo
            if (!int.TryParse(phanChan.Substring(9, 3), out nghin))
            {
                nghin = 0;
            }

            // Chuyển đổi 3 ký tự cuối cùng
            if (!int.TryParse(phanChan.Substring(12, 3), out dong))
            {
                dong = 0;
            }

            if (nghinTy == 0 && ty == 0 && trieu == 0 && nghin == 0 && dong == 0)
            {
                bangChu = "không " + DonViTien;
            }
            else
            {
                bangChu = "";
            }

            for (int i = 0; i <= 5; i++)
            {
                switch (i)
                {
                    case 0:
                        soDoi = nghinTy;
                        ten = "nghìn tỷ";
                        break;
                    case 1:
                        soDoi = ty;
                        ten = "tỷ";
                        break;
                    case 2:
                        soDoi = trieu;
                        ten = "triệu";
                        break;
                    case 3:
                        soDoi = nghin;
                        ten = "nghìn";
                        break;
                    case 4:
                        soDoi = dong;
                        ten = DonViTien;
                        break;
                    case 5:
                        soDoi = soLe;
                        ten = DonViLe;
                        break;
                    default:
                        soDoi = 0;
                        ten = "";
                        break;
                }

                if (soDoi != 0)
                {
                    tram = soDoi / 100;
                    muoi = (soDoi - tram * 100) / 10;
                    donVi = soDoi - tram * 100 - muoi * 10;

                    if (bangChu.EndsWith(" "))
                    {
                        bangChu = bangChu.TrimEnd();
                    }

                    bangChu += (bangChu.Length == 0 ? "" : ", ") +
                               (tram != 0 ? CharVND[tram] + " trăm " : "");

                    if (muoi == 0 && tram != 0 && donVi != 0)
                    {
                        bangChu += "linh ";
                    }
                    else if (muoi != 0)
                    {
                        bangChu += (muoi != 1 ? CharVND[muoi] + " mươi " : "mười ");
                    }

                    if (muoi != 0 && donVi == 5)
                    {
                        bangChu += "lăm " + ten + " ";
                    }
                    else if (muoi > 1 && donVi == 1)
                    {
                        bangChu += "mốt " + ten + " ";
                    }
                    else
                    {
                        bangChu += (donVi != 0 ? CharVND[donVi] + " " + ten : ten) + " ";
                    }
                }
                else
                {
                    if (i == 4)
                    {
                        bangChu += DonViTien + " ";
                    }
                }
            }

            //if (soLe == 0)
            //{
            //    bangChu += "chẵn";
            //}

            bangChu = bangChu.Trim();
            bangChu = char.ToUpper(bangChu[0]) + bangChu.Substring(1);

            if (Tiento != "" && Hauto != "")
            {
                bangChu = Tiento + bangChu + Hauto;
            }
            else if (Tiento != "" && Hauto == "")
            {
                bangChu = Tiento + bangChu;
            }
            else if (Tiento == "" && Hauto != "")
            {
                bangChu = bangChu + Hauto;
            }

            return bangChu;
        }
      
     }

}

