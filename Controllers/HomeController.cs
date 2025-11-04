using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using System.Xml;
using System.Xml.Linq;
using Taxweb.Models;

namespace Taxweb.Controllers
{
    public class HomeController : Controller
    {

        public static int GetLastDayOfMonth(int year, int month)
        {
            // Kiểm tra tính hợp lệ của tháng
            if (month < 1 || month > 12)
            {
                throw new ArgumentOutOfRangeException(nameof(month), "Tháng phải trong khoảng từ 1 đến 12.");
            }

            // Sử dụng DateTime để lấy ngày cuối cùng của tháng
            return DateTime.DaysInMonth(year, month);
        }
        string password, connectionString;
        public int ExecuteQueryResult(string query, params OleDbParameter[] parameters)
        {
            string password = "1@35^7*9)1";
            connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
            System.Data.DataTable dataTable = new System.Data.DataTable();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                Console.WriteLine("Kết nối đến cơ sở dữ liệu thành công!");

                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    // Thêm các tham số vào command
                    if (parameters != null)
                    {
                        command.Parameters.AddRange(parameters);
                    }

                    int rowsAffected = command.ExecuteNonQuery(); // Thực thi câu lệnh
                    return rowsAffected;
                }
            }

            return -1;
        }
        public static DateTime GetStartDateOfQuarter(int year, int quarter)
        {
            switch (quarter)
            {
                case 1: return new DateTime(year, 1, 1);   // Quý 1
                case 2: return new DateTime(year, 4, 1);   // Quý 2
                case 3: return new DateTime(year, 7, 1);   // Quý 3
                case 4: return new DateTime(year, 10, 1);  // Quý 4
                default: throw new ArgumentOutOfRangeException(nameof(quarter), "Quý phải trong khoảng từ 1 đến 4.");
            }
        }

        public static DateTime GetEndDateOfQuarter(int year, int quarter)
        {
            switch (quarter)
            {
                case 1: return new DateTime(year, 3, 31);  // Quý 1
                case 2: return new DateTime(year, 6, 30);  // Quý 2
                case 3: return new DateTime(year, 9, 30);  // Quý 3
                case 4: return new DateTime(year, 12, 31); // Quý 4
                default: throw new ArgumentOutOfRangeException(nameof(quarter), "Quý phải trong khoảng từ 1 đến 4.");
            }
        }
        public class YourDataModel
        {
            public string TenHH { get; set; }
            public double  GT1 { get; set; }
            public double GT2 { get; set; }
        }
        [HttpPost]
        public ActionResult SaveDataPL(string path,string tableData)
        {
            dbPath = path;  
            // Giải mã chuỗi JSON thành danh sách các đối tượng
            var data = JsonConvert.DeserializeObject<List<YourDataModel>>(tableData);

            // Xử lý dữ liệu và xuất XML
            // Tạo file XML và trả về cho người dùng
            string qr = @"Delete from tbPL1";
            var rowsAffected = ExecuteQueryResult(qr);
            foreach (var item in data)
            {
                if (!string.IsNullOrEmpty(item.TenHH))
                {
                    string query = @"INSERT INTO tbPL1 (TenHH,GT1,GT2) 
                           VALUES (?,?,?)";

                    var parameters = new OleDbParameter[]
                    {
                        new OleDbParameter("@TenHH", item.TenHH),
                        new OleDbParameter("@GT1", item.GT1),
                        new OleDbParameter("@GT2", item.GT2),
                    };

                     rowsAffected = ExecuteQueryResult(query, parameters);
                }

              }

            return Json(new { success = true, message = "Dữ liệu đã được xử lý" });
        }
        [HttpPost]
        public ActionResult SaveDataPL2(string path,string tableData)
        {
            dbPath = path;
            // Giải mã chuỗi JSON thành danh sách các đối tượng
            var data = JsonConvert.DeserializeObject<List<YourDataModel>>(tableData);
            string qr = @"Delete from tbPL2";
            var rowsAffected = ExecuteQueryResult(qr);
            foreach (var item in data)
            {
                if (!string.IsNullOrEmpty(item.TenHH))
                {
                    string query = @"INSERT INTO tbPL2 (TenHH,GT1,GT2) 
                           VALUES (?,?,?)";

                    var parameters = new OleDbParameter[]
                    {
                        new OleDbParameter("@TenHH", item.TenHH),
                        new OleDbParameter("@GT1", item.GT1),
                        new OleDbParameter("@GT2", item.GT2),
                    };

                     rowsAffected = ExecuteQueryResult(query, parameters);
                }
                   
            }

            return Json(new { success = true, message = "Dữ liệu đã được xử lý" });
        }
        [HttpGet] // THAY ĐỔI THÀNH GET
        public ActionResult CreateTaxXMLFull(string path,int khoasl,string ky,DateTime ngayky,string tencty,string tendaily, double N22, double N23, double N24, double N23a, double N24a, double N25,
    double N26, double N27, double N28, double N29, double N30, double N31, double N32, double N33, double N32a, double N34,
    double N35, double N36, double N37, double N38, double N39a, double N40, double N40a, double N40b, double N41, double N42, double N43,string tableData)
        {
            dbPath = path;
            string querys = "SELECT * FROM License";
            DataTable tbLicence = ExecuteQuery(querys, null);

            string qrtbPL1= "SELECT * FROM tbPL1";
            DataTable tbPL1 = ExecuteQuery(qrtbPL1, null);
            string qrtbPL2 = "SELECT * FROM tbPL2";
            DataTable tbPL2 = ExecuteQuery(qrtbPL2, null);
            //Kiểm tra tbThongTinToKhai năm hiện tại có chưa

            string query = @"SELECT * FROM tbThongTinToKhai   WHERE  Nam= ?";

            var parameterss = new OleDbParameter[]
            {
                 new OleDbParameter("?", DateTime.Now.Year.ToString())
            };
            var kq = ExecuteQuery(query, parameterss);
            double Quy1 = 0, Quy2 = 0, Quy3 = 0, Quy4 = 0;
            double T1 = 0, T2 = 0, T3 = 0, T4 = 0, T5 = 0, T6 = 0;
            double T7 = 0, T8 = 0, T9 = 0, T10 = 0, T11 = 0, T12 = 0;
            double[] Thang = { T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12 };
             
             tendaily = string.IsNullOrEmpty(tendaily) ? "Không xác định" : tendaily;
            int k1 = 0, k2 = 0, k3 = 0, k4 = 0,k5=0,k6=0,k7=0,k8=0,k9=0,k10=0,k11=0,k12=0;
            if (kq.Rows.Count == 0)
            {
                // Kiểm tra và khởi tạo tendaily
               

                query = @"INSERT INTO tbThongTinToKhai (Quy1,Quy2,Quy3,Quy4,T1,T2,T3,T4,T5,T6,T7,T8,T9,T10,T11,T12,NguoiKy,Nam,k1,k2,k3,k4,k5,k6,k7,k8,k9,k10,k11,k12) 
              VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

             
                try
                {
                    if (ky.Contains("Quý"))
                    {
                        Match match = Regex.Match(ky, @"Quý\s+(\d+)");
                        if (match.Success)
                        {
                            int soQuy = int.Parse(match.Groups[1].Value); // 4
                            if (soQuy == 1)
                            {
                                Quy1 = N43;
                                if (khoasl == 1)
                                    k1 = 1;
                            }
                            if (soQuy == 2)
                            {
                                Quy2 = N43;
                                if (khoasl == 1)
                                    k2 = 1;
                            }
                            if (soQuy == 3)
                            {
                                Quy3 = N43;
                                if (khoasl == 1)
                                    k3 = 1;
                            }
                            if (soQuy == 4)
                            {
                                Quy4 = N43;
                                if (khoasl == 1)
                                    k4 = 1;
                            }
                        }
                    }
                    else
                    {
                        if (ky.Contains("Tháng"))
                        {
                            Match match = Regex.Match(ky, @"Tháng\s+(\d+)");
                            if (match.Success)
                            {
                                int soThang = int.Parse(match.Groups[1].Value);
                                if (soThang >= 1 && soThang <= 12)
                                    Thang[soThang - 1] = N43;
                                if (soThang == 1)
                                    k1 = 1;
                                if (soThang == 2)
                                    k2 = 1;
                                if (soThang == 3)
                                    k3 = 1;
                                if (soThang == 4)
                                    k4 = 1;
                                if (soThang == 5)
                                    k5 = 1;
                                if (soThang == 6)
                                    k6 = 1;
                                if (soThang == 7)
                                    k7 = 1;
                                if (soThang == 8)
                                    k8 = 1;
                                if (soThang == 9)
                                    k9 = 1;
                                if (soThang == 10)
                                    k10 = 1;
                                if (soThang == 11)
                                    k11 = 1;
                                if (soThang == 12)
                                    k12 = 1;
                            }
                            T1 = Thang[0]; T2 = Thang[1]; T3 = Thang[2]; T4 = Thang[3];
                            T5 = Thang[4]; T6 = Thang[5]; T7 = Thang[6]; T8 = Thang[7];
                            T9 = Thang[8]; T10 = Thang[9]; T11 = Thang[10]; T12 = Thang[11];
                        }
                    }
                    var parameters = new OleDbParameter[]
                    {
                        new OleDbParameter("@Quy1", Quy1),
                        new OleDbParameter("@Quy2", Quy2),
                        new OleDbParameter("@Quy3", Quy3),
                        new OleDbParameter("@Quy4", Quy4),
                        new OleDbParameter("@T1", T1),
                        new OleDbParameter("@T2", T2),
                        new OleDbParameter("@T3", T3),
                        new OleDbParameter("@T4", T4),
                        new OleDbParameter("@T5", T5),
                        new OleDbParameter("@T6", T6),
                        new OleDbParameter("@T7", T7),
                        new OleDbParameter("@T8", T8),
                        new OleDbParameter("@T9", T9),
                        new OleDbParameter("@T10", T10),
                        new OleDbParameter("@T11", T11),
                        new OleDbParameter("@T12", T12),
                        new OleDbParameter("@NguoiKy", tendaily),
                        new OleDbParameter("@Nam", DateTime.Now.Year),
                        new OleDbParameter("@k1", k1),
                        new OleDbParameter("@k2", k2),
                        new OleDbParameter("@k3", k3),
                        new OleDbParameter("@k4", k4),
                        new OleDbParameter("@k5", k5),
                        new OleDbParameter("@k6", k6),
                        new OleDbParameter("@k7", k7),
                        new OleDbParameter("@k8", k8),
                        new OleDbParameter("@k9", k9),
                        new OleDbParameter("@k10", k10),
                        new OleDbParameter("@k11", k11),
                        new OleDbParameter("@k12", k12),
                    };

                    var rowsAffected = ExecuteQueryResult(query, parameters);
                }
                catch (Exception ex)
                {
                }
            }
            else
            {


                // UPDATE khi đã có dữ liệu
                query = @"UPDATE tbThongTinToKhai 
              SET Quy1 = ?, Quy2 = ?, Quy3 = ?, Quy4 = ?,
                  T1 = ?, T2 = ?, T3 = ?, T4 = ?, T5 = ?, T6 = ?,
                  T7 = ?, T8 = ?, T9 = ?, T10 = ?, T11 = ?, T12 = ?,
                  NguoiKy = ?, Nam = ?,k1=?,k2=?,k3=?,k4=?,k5=?,k6=?,k7=?,k8=?,k9=?,k10=?,k11=?,k12=?
              WHERE Id = ?"; // Giả sử có trường Id làm khóa chính

                DataRow row = kq.Rows[0];

                // Gán giá trị cho các biến Quý
                Quy1 = row.Field<double>("Quy1");
                Quy2 = row.Field<double>("Quy2");
                Quy3 = row.Field<double>("Quy3");
                Quy4 = row.Field<double>("Quy4");

                // Gán giá trị cho các biến Tháng
                T1 = row.Field<double>("T1");
                T2 = row.Field<double>("T2");
                T3 = row.Field<double>("T3");
                T4 = row.Field<double>("T4");
                T5 = row.Field<double>("T5");
                T6 = row.Field<double>("T6");
                T7 = row.Field<double>("T7");
                T8 = row.Field<double>("T8");
                T9 = row.Field<double>("T9");
                T10 = row.Field<double>("T10");
                T11 = row.Field<double>("T11");
                T12 = row.Field<double>("T12");
                k1 = int.Parse(row.Field<string>("k1"));
                k2 = int.Parse(row.Field<string>("k2"));
                k3 = int.Parse(row.Field<string>("k3"));
                k4 = int.Parse(row.Field<string>("k4"));
                k5 = int.Parse(row.Field<string>("k5"));
                k6 = int.Parse(row.Field<string>("k6"));
                k7 = int.Parse(row.Field<string>("k7"));
                k8 = int.Parse(row.Field<string>("k8"));
                k9 = int.Parse(row.Field<string>("k9"));
                k10 = int.Parse(row.Field<string>("k10"));
                k11 = int.Parse(row.Field<string>("k11"));
                k12 = int.Parse(row.Field<string>("k12"));
                // Gán giá trị cho NguoiKy (string)
                tendaily = tendaily;


                if (ky.Contains("Quý"))
                {
                    Match match = Regex.Match(ky, @"Quý\s+(\d+)");
                    if (match.Success)
                    {
                        int soQuy = int.Parse(match.Groups[1].Value); // 4
                        if (soQuy == 1)
                        {
                            Quy1 = N43;
                            if (khoasl == 1)
                                k1 = 1;
                            else
                                k1 = 0;
                        }
                           
                        if (soQuy == 2)
                        {
                            Quy2 = N43;
                            if (khoasl == 1)
                                k2 = 1;
                            else
                                k2 = 0;
                        }
                        if (soQuy == 3)
                        {
                            Quy3 = N43;
                            if (khoasl == 1)
                                k3 = 1;
                            else
                                k3 = 0;
                        }
                        if (soQuy == 4)
                        {
                            Quy4 = N43;
                            if (khoasl == 1)
                                k4 = 1;
                            else
                                k4 = 0;
                        }
                    }
                }
                else
                {
                    if (ky.Contains("Tháng"))
                    {
                        Match match = Regex.Match(ky, @"Tháng\s+(\d+)");
                        if (match.Success)
                        {
                            int soThang = int.Parse(match.Groups[1].Value);
                            if (soThang >= 1 && soThang <= 12)
                                Thang[soThang - 1] = N43;
                            if (soThang == 1)
                                k1 = 1;
                            if (soThang == 2)
                                k2 = 1;
                            if (soThang == 3)
                                k3 = 1;
                            if (soThang == 4)
                                k4 = 1;
                            if (soThang == 5)
                                k5 = 1;
                            if (soThang == 6)
                                k6 = 1;
                            if (soThang == 7)
                                k7 = 1;
                            if (soThang == 8)
                                k8 = 1;
                            if (soThang == 9)
                                k9 = 1;
                            if (soThang == 10)
                                k10 = 1;
                            if (soThang == 11)
                                k11 = 1;
                            if (soThang == 12)
                                k12 = 1;
                        }
                        T1 = Thang[0]; T2 = Thang[1]; T3 = Thang[2]; T4 = Thang[3];
                        T5 = Thang[4]; T6 = Thang[5]; T7 = Thang[6]; T8 = Thang[7];
                        T9 = Thang[8]; T10 = Thang[9]; T11 = Thang[10]; T12 = Thang[11];


                    }
                }
                var parameters = new OleDbParameter[]
                {
        new OleDbParameter("?", Quy1),
        new OleDbParameter("?", Quy2),
        new OleDbParameter("?", Quy3),
        new OleDbParameter("?", Quy4),
        new OleDbParameter("?", T1),
        new OleDbParameter("?", T2),
        new OleDbParameter("?", T3),
        new OleDbParameter("?", T4),
        new OleDbParameter("?", T5),
        new OleDbParameter("?", T6),
        new OleDbParameter("?", T7),
        new OleDbParameter("?", T8),
        new OleDbParameter("?", T9),
        new OleDbParameter("?", T10),
        new OleDbParameter("?", T11),
        new OleDbParameter("?", T12),
        new OleDbParameter("?", tendaily),
        new OleDbParameter("?", DateTime.Now.Year),
        new OleDbParameter("?", k1),
        new OleDbParameter("?", k2),
        new OleDbParameter("?", k3),
        new OleDbParameter("?", k4),
        new OleDbParameter("?", k5),
        new OleDbParameter("?", k6),
        new OleDbParameter("?", k7),
        new OleDbParameter("?", k8),
        new OleDbParameter("?", k9),
        new OleDbParameter("?", k10),
        new OleDbParameter("?", k11),
        new OleDbParameter("?", k12),
        new OleDbParameter("?", kq.Rows[0]["Id"]) // Lấy Id từ dòng đầu tiên
                };

                var rowsAffected = ExecuteQueryResult(query, parameters);
            }


                var sb = new StringBuilder();
            string ct = "";
            if (ky.Contains("Q"))
            {
                int soQuy = 0;
                Match match = Regex.Match(ky, @"Quý\s+(\d+)");
                if (match.Success)
                {
                    soQuy = int.Parse(match.Groups[1].Value);

                }
                int year = DateTime.Now.Year;
                  

                DateTime startDate = GetStartDateOfQuarter(year, soQuy);
                DateTime endDate = GetEndDateOfQuarter(year, soQuy);

                ct = $@"<KyKKhaiThue>
<kieuKy>Q</kieuKy> 
<kyKKhai>{soQuy}/{year}</kyKKhai>
<kyKKhaiTuNgay>{startDate.ToString("dd/MM/yyyy")}</kyKKhaiTuNgay>
<kyKKhaiDenNgay>{endDate.ToString("dd/MM/yyyy")}</kyKKhaiDenNgay>
<kyKKhaiTuThang/>
<kyKKhaiDenThang/>
</KyKKhaiThue>";
            }
            else
            {
                int soThang = 0;
                Match match = Regex.Match(ky, @"Tháng\s+(\d+)");
                if (match.Success)
                {
                     soThang = int.Parse(match.Groups[1].Value);
                  
                }
                int lastDay = GetLastDayOfMonth(DateTime.Now.Month, soThang);
                int year = DateTime.Now.Year;
                DateTime fd = new DateTime(year, soThang, 1);
                DateTime td = new DateTime(year, soThang, lastDay);
                ct = $@"<KyKKhaiThue>
<kieuKy>M</kieuKy>
<kyKKhai>{soThang.ToString("D2") + "/"+year.ToString()}</kyKKhai>
<kyKKhaiTuNgay>{fd.ToString("dd/MM/yyyy")}</kyKKhaiTuNgay>
<kyKKhaiDenNgay>{td.ToString("dd/MM/yyyy")}</kyKKhaiDenNgay>
<kyKKhaiTuThang/>
<kyKKhaiDenThang/>
</KyKKhaiThue>";
            }
            var sbdv = new StringBuilder();
            string dv = "";
            int i = 1;
            foreach (DataRow item in tbPL1.Rows)
            {
                dv += $@"<BangKeTenHHDV ID=""ID_{i}"">
<tenHHDVMuaVao>{item.Field<string>("TenHH")}</tenHHDVMuaVao>
<giaTriHHDVMuaVao>{item.Field<double>("GT1")}</giaTriHHDVMuaVao>
<thueGTGTHHDV>{item.Field<double>("GT2")}</thueGTGTHHDV>
</BangKeTenHHDV>";
                i++;
            }
            dv = "";
            dv += $@"<BangKeTenHHDV ID=""ID_{i}"">
<tenHHDVMuaVao>Hàng hóa, dịch vụ mua vào trong kỳ được áp dụng mức thuế suất thuế giá trị gia tăng 8%</tenHHDVMuaVao>
<giaTriHHDVMuaVao>{tbPL1.AsEnumerable().Sum(m => m.Field<double>("GT1"))}</giaTriHHDVMuaVao>
<thueGTGTHHDV>{tbPL1.AsEnumerable().Sum(m => m.Field<double>("GT2"))}</thueGTGTHHDV>
</BangKeTenHHDV>";

            dv += $@"<tongCongGiaTriHHDVMuaVao>{tbPL1.AsEnumerable().Sum(m=>m.Field<double>("GT1"))}</tongCongGiaTriHHDVMuaVao>";
            dv += $@"<tongCongThueGTGTHHDV>{tbPL1.AsEnumerable().Sum(m => m.Field<double>("GT2"))}</tongCongThueGTGTHHDV>";
            string dt1 = @"
<BangKeTenHHDV ID=""ID_1"">
<tenHHDVMuaVao>Sơn, bột trét</tenHHDVMuaVao>
<giaTriHHDVMuaVao>2357914513</giaTriHHDVMuaVao>
<thueGTGTHHDV>188633164</thueGTGTHHDV>
</BangKeTenHHDV>";


            var sbdr = new StringBuilder();
            string dv2 = "";
            i = 1;
            foreach (DataRow item in tbPL2.Rows)
            {
                dv2 += $@"<BangKeTenHHDV ID=""ID_1"">
<tenHHDV>{item.Field<string>("TenHH")}</tenHHDV>
<giaTriHHDV>{item.Field<double>("GT1")}</giaTriHHDV>
<thueSuatTheoQuyDinh>10</thueSuatTheoQuyDinh>
<thueSuatSauGiam>8</thueSuatSauGiam>
<thueGTGTDuocGiam>{item.Field<double>("GT2")}</thueGTGTDuocGiam>
</BangKeTenHHDV>";
            }
            dv2 = "";
            decimal tong3 = decimal.Parse(tbPL2.AsEnumerable().Sum(m => m.Field<double>("GT1")).ToString());
            var tong4 = Math.Round(tong3 * 0.02m);
            dv2 += $@"<BangKeTenHHDV ID=""ID_1"">
<tenHHDV>Hàng hóa, dịch vụ bán ra trong kỳ được áp dụng mức thuế suất thuế giá trị gia tăng 8%</tenHHDV>
<giaTriHHDV>{tbPL2.AsEnumerable().Sum(m => m.Field<double>("GT1"))}</giaTriHHDV>
<thueSuatTheoQuyDinh>10</thueSuatTheoQuyDinh>
<thueSuatSauGiam>8</thueSuatSauGiam>
<thueGTGTDuocGiam>{tong4}</thueGTGTDuocGiam>
</BangKeTenHHDV>";
         
            dv2 += $@"<tongCongGiaTriHHDV>{tong3}</tongCongGiaTriHHDV>";
            dv2 += $@"<tongCongThueGTGTDuocGiam>{tong4}</tongCongThueGTGTDuocGiam>";
                // Tạo XML string đúng y hệt mẫu
                string xmlContent = $@"<HSoThueDTu xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://kekhaithue.gdt.gov.vn/TKhaiThue"" >
<HSoKhaiThue id=""ID_1"">
<TTinChung>
<TTinDVu>
<maDVu>HTKK</maDVu>
<tenDVu>HỖ TRỢ KÊ KHAI THUẾ</tenDVu>
<pbanDVu>5.4.6</pbanDVu>
<ttinNhaCCapDVu>8DB3440999D239B5C635D46EF864F0FB</ttinNhaCCapDVu>
</TTinDVu>  
<TTinTKhaiThue>
<TKhaiThue>
<maTKhai>842</maTKhai>
<tenTKhai>TỜ KHAI THUẾ GIÁ TRỊ GIA TĂNG (Mẫu số 01/GTGT)</tenTKhai>
<moTaBMau>(Ban hành kèm theo Thông tư số 80/2021/TT-BTC ngày 29 tháng 9 năm 2021 của Bộ trưởng Bộ Tài chính)</moTaBMau>
<pbanTKhaiXML>2.8.3</pbanTKhaiXML>
<loaiTKhai>C</loaiTKhai>
<soLan>0</soLan>
{sb.Append(ct)}
<maCQTNoiNop>71701</maCQTNoiNop>
<tenCQTNoiNop>Thuế cơ sở 24 Thành phố Hồ Chí Minh</tenCQTNoiNop>
<ngayLapTKhai>2025-10-10</ngayLapTKhai>
<GiaHan>
<maLyDoGiaHan/>
<lyDoGiaHan/>
</GiaHan>
<nguoiKy>{tendaily}</nguoiKy>
<ngayKy>{ngayky.ToString("yyyy-MM-dd")}</ngayKy>
<nganhNgheKD/>
</TKhaiThue>
<NNT>
<mst>{tbLicence.Rows[0].Field<string>("MaSoThue")}</mst>
<tenNNT>{tencty}</tenNNT>
<dchiNNT>{Helpers.ConvertVniToUnicode(tbLicence.Rows[0].Field<string>("DiaChi"))}</dchiNNT>
<phuongXa/>
<maHuyenNNT>71701</maHuyenNNT>
<tenHuyenNNT/>
<maTinhNNT>701</maTinhNNT>
<tenTinhNNT>Thành phố Hồ Chí Minh</tenTinhNNT>
<dthoaiNNT/>
<faxNNT/>
<emailNNT/>
</NNT>
</TTinTKhaiThue>
</TTinChung>
<CTieuTKhaiChinh>
<ma_NganhNghe>00</ma_NganhNghe>
<ten_NganhNghe>Hoạt động sản xuất kinh doanh thông thường</ten_NganhNghe>
<tieuMucHachToan>1701</tieuMucHachToan>
<Header>
<ct09/>
<ct10/>
<DiaChiHDSXKDKhacTinhNDTSC>
<ct11a_phuongXa_ma/>
<ct11a_phuongXa_ten/>
<ct11b_quanHuyen_ma/>
<ct11b_quanHuyen_ten/>
<ct11c_tinhTP_ma/>
<ct11c_tinhTP_ten/>
</DiaChiHDSXKDKhacTinhNDTSC>
</Header>
<ct21>0</ct21>
<ct22>{N22}</ct22>
<GiaTriVaThueGTGTHHDVMuaVao>
<ct23>{N23}</ct23>
<ct24>{N24}</ct24>
</GiaTriVaThueGTGTHHDVMuaVao>
<HangHoaDichVuNhapKhau>
<ct23a>{N23a}</ct23a>
<ct24a>{N24a}</ct24a>
</HangHoaDichVuNhapKhau>
<ct25>{N25}</ct25>
<ct26>{N26}</ct26>
<HHDVBRaChiuThueGTGT>
<ct27>{N27}</ct27>
<ct28>{N28}</ct28>
</HHDVBRaChiuThueGTGT>
<ct29>{N29}</ct29>
<HHDVBRaChiuTSuat5>
<ct30>{N30}</ct30>
<ct31>{N31}</ct31>
</HHDVBRaChiuTSuat5>
<HHDVBRaChiuTSuat10>
<ct32>{N32}</ct32>
<ct33>{N33}</ct33>
</HHDVBRaChiuTSuat10>
<ct32a>{N32a}</ct32a>
<TongDThuVaThueGTGTHHDVBRa>
<ct34>{N34}</ct34>
<ct35>{N35}</ct35>
</TongDThuVaThueGTGTHHDVBRa>
<ct36>{N36}</ct36>
<ct37>{N37}</ct37>
<ct38>{N38}</ct38>
<ct39a>{N39a}</ct39a>
<ct40a>{N40a}</ct40a>
<ct40b>{N40b}</ct40b>
<ct40>{N40}</ct40>
<ct41>{N41}</ct41>
<ct42>{N42}</ct42>
<ct43>{N43}</ct43>
</CTieuTKhaiChinh>
<PLuc>
<PL_NQ142_GTGT>
<HH_DV_MuaVaoTrongKy> 
{sbdv.Append(dv)} 
</HH_DV_MuaVaoTrongKy>
<HH_DV_BanRaTrongKy>
{sbdr.Append(dv2)} 
</HH_DV_BanRaTrongKy>
<ChenhLech>
<ct9>-138263964</ct9>
</ChenhLech>
</PL_NQ142_GTGT>
</PLuc>
</HSoKhaiThue>
</HSoThueDTu>";

            // Trả về file XML để download
            byte[] fileBytes = Encoding.UTF8.GetBytes(xmlContent);
            //Lưu xml contents
            if (ky.Contains("Q"))
            {
                int soQuy = 0;
                string xml1="";string xml2="";string xml3="";string xml4 = "";
                if (kq.Rows.Count > 0)
                {
                    DataRow row = kq.Rows[0];

                    xml1 = row.Field<string>("xml1");
                    xml2 = row.Field<string>("xml2");
                    xml3 = row.Field<string>("xml3");
                    xml4 = row.Field<string>("xml4");
                }
               
                Match match = Regex.Match(ky, @"Quý\s+(\d+)");
                if (match.Success)
                {
                    soQuy = int.Parse(match.Groups[1].Value);
                    if (soQuy == 1)
                    {
                        xml1 = xmlContent;
                    }
                    if (soQuy == 2)
                    {
                        xml2 = xmlContent;
                    }
                    if (soQuy == 3)
                    {
                        xml3 = xmlContent;
                    }
                    if (soQuy == 4)
                    {
                        xml4 = xmlContent;
                    }
                    query = @"UPDATE tbThongTinToKhai 
              SET xml1 = ?,xml2=?,xml3=?,xml4=?
              WHERE Id = ?"; // Giả sử có trường Id làm khóa chính
                    var parameters = new OleDbParameter[]
               {
        new OleDbParameter("?", xml1!=null?xml1:""),
        new OleDbParameter("?", xml2!=null?xml2:""),
        new OleDbParameter("?", xml3 != null ? xml3 : ""),
        new OleDbParameter("?", xml4 != null ? xml4 : ""),
        new OleDbParameter("?", kq.Rows[0]["Id"]) // Lấy Id từ dòng đầu tiên
               };

                    var rowsAffected = ExecuteQueryResult(query, parameters);
                }
            }
            else
            {
                string xml1 = ""; string xml2 = ""; string xml3 = ""; string xml4 = "";
                string xml5 = ""; string xml6 = ""; string xml7 = ""; string xml8 = "";
                string xml9 = ""; string xml10 = ""; string xml11 = ""; string xml12 = "";

                if (kq.Rows.Count > 0)
                {
                    DataRow row = kq.Rows[0];

                    xml1 = row.Field<string>("xml1");
                    xml2 = row.Field<string>("xml2");
                    xml3 = row.Field<string>("xml3");
                    xml4 = row.Field<string>("xml4");
                    xml5 = row.Field<string>("xml5");
                    xml6 = row.Field<string>("xml6");
                    xml7 = row.Field<string>("xml7");
                    xml8 = row.Field<string>("xml8");
                    xml9 = row.Field<string>("xml9");
                    xml10 = row.Field<string>("xml10");
                    xml11 = row.Field<string>("xml11");
                    xml12 = row.Field<string>("xml12");
                }
                int soThang = 0;
                Match match = Regex.Match(ky, @"Tháng\s+(\d+)");
                if (match.Success)
                {
                    soThang = int.Parse(match.Groups[1].Value);
                    if(soThang==1)
                        xml1 = xmlContent;
                    if (soThang == 2)
                        xml2 = xmlContent;
                    if (soThang == 3)
                        xml3 = xmlContent;
                    if (soThang == 4)
                        xml4 = xmlContent;
                    if (soThang == 5)
                        xml5 = xmlContent;
                    if (soThang == 6)
                        xml6 = xmlContent;
                    if (soThang == 7)
                        xml7 = xmlContent;
                    if (soThang == 8)
                        xml8 = xmlContent;
                    if (soThang == 9)
                        xml9 = xmlContent;
                    if (soThang == 10)
                        xml10 = xmlContent;
                    if (soThang == 11)
                        xml11 = xmlContent;
                    if (soThang == 12)
                        xml12 = xmlContent;

                }
                query = @"UPDATE tbThongTinToKhai 
              SET xml1 = ?,xml2=?,xml3=?,xml4=?,xml5=?,xml6=?,xml7=?,xml8=?,xml9=?,xml10=?,xml11=?,xml12=?
              WHERE Id = ?"; // Giả sử có trường Id làm khóa chính
                var parameters = new OleDbParameter[]
           {
                    new OleDbParameter("?", xml1!=null?xml1:""),
                    new OleDbParameter("?", xml2!=null?xml2:""),
                    new OleDbParameter("?", xml3 != null ? xml3 : ""),
                    new OleDbParameter("?", xml4 != null ? xml4 : ""),
                    new OleDbParameter("?", xml5 != null ? xml5 : ""),
                    new OleDbParameter("?", xml6 != null ? xml6 : ""),
                    new OleDbParameter("?", xml7 != null ? xml7 : ""),
                    new OleDbParameter("?", xml8 != null ? xml8 : ""),
                    new OleDbParameter("?", xml9 != null ? xml9 : ""),
                    new OleDbParameter("?", xml10 != null ? xml10 : ""),
                    new OleDbParameter("?", xml11!= null ? xml11 : ""),
                    new OleDbParameter("?", xml12!= null ? xml12 : ""),
                    new OleDbParameter("?", kq.Rows[0]["Id"]) // Lấy Id từ dòng đầu tiên
           };

                var rowsAffected = ExecuteQueryResult(query, parameters);
            }
                return File(fileBytes, "application/xml", "HSoThueDTu.xml");
        }

        private void WriteXMLContent(XmlWriter writer)
        {
            writer.WriteStartDocument();

            // Root element với namespaces - SỬA LẠI CÁCH KHAI BÁO NAMESPACE
            writer.WriteStartElement("HSoThueDTu", "http://kekhaithue.gdt.gov.vn/TKhaiThue");
            writer.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
            writer.WriteAttributeString("xmlns", "ds", null, "http://www.w3.org/2000/09/xmldsig#");

            // HSoKhaiThue
            writer.WriteStartElement("HSoKhaiThue");
            writer.WriteAttributeString("id", "ID_1");

            // TTinChung
            writer.WriteStartElement("TTinChung");

            // Thông tin dịch vụ
            AddTTinDVu(writer);

            // Thông tin tờ khai thuế
            AddTTinTKhaiThue(writer);

            writer.WriteEndElement(); // TTinChung

            // Chi tiết tờ khai chính
            AddCTieuTKhaiChinh(writer);

            // Phụ lục
            AddPLuc(writer);

            writer.WriteEndElement(); // HSoKhaiThue

            // Chữ ký số
            AddCKyDTu(writer);

            writer.WriteEndElement(); // HSoThueDTu
            writer.WriteEndDocument();
        }

        private void AddTTinDVu(XmlWriter writer)
        {
            writer.WriteStartElement("TTinDVu");

            writer.WriteElementString("maDVu", "HTKK");
            writer.WriteElementString("tenDVu", "HỖ TRỢ KÊ KHAI THUẾ");
            writer.WriteElementString("pbanDVu", "5.4.5");
            writer.WriteElementString("ttinNhaCCapDVu", "33A52A87ECC4AD58652D8FF252B604F5");

            writer.WriteEndElement();
        }

        private void AddTTinTKhaiThue(XmlWriter writer)
        {


            writer.WriteStartElement("TTinTKhaiThue");

            // TKhaiThue
            writer.WriteStartElement("TKhaiThue");
            writer.WriteElementString("maTKhai", "842");
            writer.WriteElementString("tenTKhai", "TỜ KHAI THUẾ GIÁ TRỊ GIA TĂNG (Mẫu số 01/GTGT)");
            writer.WriteElementString("moTaBMau", "(Ban hành kèm theo Thông tư số 80/2021/TT-BTC ngày 29 tháng 9 năm 2021 của Bộ trưởng Bộ Tài chính)");
            writer.WriteElementString("pbanTKhaiXML", "2.8.3");
            writer.WriteElementString("loaiTKhai", "C");
            writer.WriteElementString("soLan", "0");

            // KyKKhaiThue
            writer.WriteStartElement("KyKKhaiThue");
            writer.WriteElementString("kieuKy", "Q");
            writer.WriteElementString("kyKKhai", "3/2025");
            writer.WriteElementString("kyKKhaiTuNgay", "01/07/2025");
            writer.WriteElementString("kyKKhaiDenNgay", "30/09/2025");
            writer.WriteElementString("kyKKhaiTuThang", "");
            writer.WriteElementString("kyKKhaiDenThang", "");
            writer.WriteEndElement(); // KyKKhaiThue

            writer.WriteElementString("maCQTNoiNop", "71701");
            writer.WriteElementString("tenCQTNoiNop", "Thuế cơ sở 24 Thành phố Hồ Chí Minh");
            writer.WriteElementString("ngayLapTKhai", "2025-10-10");

            // GiaHan
            writer.WriteStartElement("GiaHan");
            writer.WriteElementString("maLyDoGiaHan", "");
            writer.WriteElementString("lyDoGiaHan", "");
            writer.WriteEndElement(); // GiaHan

            writer.WriteElementString("nguoiKy", "Vũ Đình Dân");
            writer.WriteElementString("ngayKy", "2025-10-10");
            writer.WriteElementString("nganhNgheKD", "");
            writer.WriteEndElement(); // TKhaiThue

            // NNT
            writer.WriteStartElement("NNT");
            writer.WriteElementString("mst", "3500779171");
            writer.WriteElementString("tenNNT", "axc");
            writer.WriteElementString("dchiNNT", "asdd");
            writer.WriteElementString("phuongXa", "");
            writer.WriteElementString("maHuyenNNT", "71701");
            writer.WriteElementString("tenHuyenNNT", "");
            writer.WriteElementString("maTinhNNT", "701");
            writer.WriteElementString("tenTinhNNT", "Thành phố Hồ Chí Minh");
            writer.WriteElementString("dthoaiNNT", "");
            writer.WriteElementString("faxNNT", "");
            writer.WriteElementString("emailNNT", "");
            writer.WriteEndElement(); // NNT

            writer.WriteEndElement(); // TTinTKhaiThue
        }
        double N23, N24, N24a, N25, N26, N27, N28, N29, N30, N31, N32, N33, N34, N35, N36 = 0;
        public void TinhMuavao()
        {
            string query0 = @"SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 
                      FROM HoaDon 
                      INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo  
                      WHERE HoaDon.Loai=-1 AND DC=0  AND (ThangCT>= ? AND ThangCT<= ?) ";
            var parameters = new OleDbParameter[]
            {
                new OleDbParameter("@FromMonth", 10), // Thay bằng giá trị thực tế
                new OleDbParameter("@ToMonth", 10)    // Thay bằng giá trị thực tế
           };

            var data = ExecuteQuery(query0, parameters);
            N23 = data.Rows[0].Field<double>("F1");
            N24 = data.Rows[0].Field<double>("F2");
            N25 = N24 + N24a;
        }
        public void TinhN()
        {
            string query0 = @"
                SELECT 
                    Sum(ThanhTien) AS F1, 
                    SUM(IIF(TK_ID=3007, SoPS, -SoPS)) AS F2  
                FROM 
                    (HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) 
                    LEFT JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo  
                WHERE 
                    HoaDon.Loai=1 
                    AND DC=0 
                    AND (ThangCT>= ? AND ThangCT<= ?) 
                    AND KCT=0 
                    AND RIGHT(HethongTK.SoHieu,0) = '' 
                    AND TyLe=0";

            // Tạo parameters
            var parameters = new OleDbParameter[]
{
                new OleDbParameter("@FromMonth", 10), // Thay bằng giá trị thực tế
                new OleDbParameter("@ToMonth", 10)    // Thay bằng giá trị thực tế
            };

            var data = ExecuteQuery(query0, parameters);
            N29 = data.Rows[0].Field<double>("F1");

            string query5 = @"
                SELECT 
                    Sum(ThanhTien) AS F1, 
                    SUM(IIF(TK_ID=3007, SoPS, -SoPS)) AS F2  
                FROM 
                    (HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) 
                    LEFT JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo  
                WHERE 
                    HoaDon.Loai=1 
                    AND DC=0 
                    AND (ThangCT>= ? AND ThangCT<= ?) 
                    AND KCT=0 
                    AND RIGHT(HethongTK.SoHieu,0) = '' 
                    AND TyLe=5";

            // Tạo parameters
            parameters = new OleDbParameter[]
{
                new OleDbParameter("@FromMonth", 10), // Thay bằng giá trị thực tế
                new OleDbParameter("@ToMonth", 10)    // Thay bằng giá trị thực tế
           };

            data = ExecuteQuery(query5, parameters);
            N30 = data.Rows[0].Field<double>("F1");
            N31 = data.Rows[0].Field<double>("F2");

            // HHDVBRaChiuTSuat10
            string query10 = @"
                SELECT 
                    Sum(ThanhTien) AS F1, 
                    SUM(IIF(TK_ID=3007, SoPS, -SoPS)) AS F2  
                FROM 
                    (HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) 
                    LEFT JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo  
                WHERE 
                    HoaDon.Loai=1 
                    AND DC=0 
                    AND (ThangCT>= ? AND ThangCT<= ?) 
                    AND KCT=0 
                    AND RIGHT(HethongTK.SoHieu,0) = '' 
                    AND TyLe=10";

            // Tạo parameters
            parameters = new OleDbParameter[]
{
                new OleDbParameter("@FromMonth", 10), // Thay bằng giá trị thực tế
                new OleDbParameter("@ToMonth", 10)    // Thay bằng giá trị thực tế
           };

            data = ExecuteQuery(query10, parameters);
            N32 = data.Rows[0].Field<double>("F1");
            N33 = data.Rows[0].Field<double>("F2");

            N27 = N29 + N30 + N32;
            N28 = N31 + N33;
        }
        private void AddCTieuTKhaiChinh(XmlWriter writer)
        {

            writer.WriteStartElement("CTieuTKhaiChinh");

            writer.WriteElementString("ma_NganhNghe", "00");
            writer.WriteElementString("ten_NganhNghe", "Hoạt động sản xuất kinh doanh thông thường");
            writer.WriteElementString("tieuMucHachToan", "1701");

            // Header
            writer.WriteStartElement("Header");
            writer.WriteElementString("ct09", "");
            writer.WriteElementString("ct10", "");

            // DiaChiHDSXKDKhacTinhNDTSC
            writer.WriteStartElement("DiaChiHDSXKDKhacTinhNDTSC");
            writer.WriteElementString("ct11a_phuongXa_ma", "");
            writer.WriteElementString("ct11a_phuongXa_ten", "");
            writer.WriteElementString("ct11b_quanHuyen_ma", "");
            writer.WriteElementString("ct11b_quanHuyen_ten", "");
            writer.WriteElementString("ct11c_tinhTP_ma", "");
            writer.WriteElementString("ct11c_tinhTP_ten", "");
            writer.WriteEndElement(); // DiaChiHDSXKDKhacTinhNDTSC

            writer.WriteEndElement(); // Header

            writer.WriteElementString("ct21", "0");
            writer.WriteElementString("ct22", "0");

            // GiaTriVaThueGTGTHHDVMuaVao
            writer.WriteStartElement("GiaTriVaThueGTGTHHDVMuaVao");
            writer.WriteElementString("ct23", N23.ToString());
            writer.WriteElementString("ct24", N24.ToString());
            writer.WriteEndElement();

            // HangHoaDichVuNhapKhau
            writer.WriteStartElement("HangHoaDichVuNhapKhau");
            writer.WriteElementString("ct23a", "0");
            writer.WriteElementString("ct24a", "0");
            writer.WriteEndElement();

            writer.WriteElementString("ct25", N25.ToString());
            writer.WriteElementString("ct26", "0");

            // HHDVBRaChiuThueGTGT

            writer.WriteStartElement("HHDVBRaChiuThueGTGT");
            writer.WriteElementString("ct27", N27.ToString());
            writer.WriteElementString("ct28", N28.ToString());
            writer.WriteEndElement();


            writer.WriteElementString("ct29", N29.ToString());

            // HHDVBRaChiuTSuat5 
            writer.WriteStartElement("HHDVBRaChiuTSuat5");
            writer.WriteElementString("ct30", N30.ToString());
            writer.WriteElementString("ct31", N31.ToString());
            writer.WriteEndElement();



            writer.WriteStartElement("HHDVBRaChiuTSuat10");
            writer.WriteElementString("ct32", N32.ToString());

            writer.WriteElementString("ct33", N33.ToString());
            writer.WriteEndElement();

            writer.WriteElementString("ct32a", "0");

            // TongDThuVaThueGTGTHHDVBRa
            N34 = N26 + N27;
            N35 = N28;
            writer.WriteStartElement("TongDThuVaThueGTGTHHDVBRa");
            writer.WriteElementString("ct34", N34.ToString());
            writer.WriteElementString("ct35", N35.ToString());
            writer.WriteEndElement();

            writer.WriteElementString("ct36", "31100874");
            writer.WriteElementString("ct37", "0");
            writer.WriteElementString("ct38", "0");
            writer.WriteElementString("ct39a", "0");
            writer.WriteElementString("ct40a", "31100874");
            writer.WriteElementString("ct40b", "0");
            writer.WriteElementString("ct40", "31100874");
            writer.WriteElementString("ct41", "0");
            writer.WriteElementString("ct42", "0");
            writer.WriteElementString("ct43", "0");

            writer.WriteEndElement(); // CTieuTKhaiChinh
        }

        private void AddPLuc(XmlWriter writer)
        {
            writer.WriteStartElement("PLuc");

            writer.WriteStartElement("PL_NQ142_GTGT");

            // HH_DV_MuaVaoTrongKy
            writer.WriteStartElement("HH_DV_MuaVaoTrongKy");

            AddBangKeHHDV(writer, "ID_1", "Sơn, bột trét", "2357914513", "188633164");
            AddBangKeHHDV(writer, "ID_2", "Xăng", "5359199", "428736");
            AddBangKeHHDV(writer, "ID_3", "Sửa chữa, bảo dưỡng xe ô tô", "16370370", "1309630");
            AddBangKeHHDV(writer, "ID_4", "Gia hạn chữ ký số", "4744444", "379556");
            AddBangKeHHDV(writer, "ID_5", "Cọ lăn", "23446346", "1875708");

            writer.WriteElementString("tongCongGiaTriHHDVMuaVao", "2407834872");
            writer.WriteElementString("tongCongThueGTGTHHDV", "192626794");
            writer.WriteEndElement(); // HH_DV_MuaVaoTrongKy

            // HH_DV_BanRaTrongKy
            writer.WriteStartElement("HH_DV_BanRaTrongKy");

            AddBangKeHHDVBanRa(writer, "ID_1", "Sơn, bột trét", "2718141480", "10", "8", "54362830");

            writer.WriteElementString("tongCongGiaTriHHDV", "2718141480");
            writer.WriteElementString("tongCongThueGTGTDuocGiam", "54362830");
            writer.WriteEndElement(); // HH_DV_BanRaTrongKy

            // ChenhLech
            writer.WriteStartElement("ChenhLech");
            writer.WriteElementString("ct9", "-138263964");
            writer.WriteEndElement(); // ChenhLech

            writer.WriteEndElement(); // PL_NQ142_GTGT
            writer.WriteEndElement(); // PLuc
        }

        private void AddBangKeHHDV(XmlWriter writer, string id, string ten, string giaTri, string thueGTGT)
        {
            writer.WriteStartElement("BangKeTenHHDV");
            writer.WriteAttributeString("ID", id);
            writer.WriteElementString("tenHHDVMuaVao", ten);
            writer.WriteElementString("giaTriHHDVMuaVao", giaTri);
            writer.WriteElementString("thueGTGTHHDV", thueGTGT);
            writer.WriteEndElement();
        }

        private void AddBangKeHHDVBanRa(XmlWriter writer, string id, string ten, string giaTri, string thueSuatQuyDinh, string thueSuatSauGiam, string thueGTGTDuocGiam)
        {
            writer.WriteStartElement("BangKeTenHHDV");
            writer.WriteAttributeString("ID", id);
            writer.WriteElementString("tenHHDV", ten);
            writer.WriteElementString("giaTriHHDV", giaTri);
            writer.WriteElementString("thueSuatTheoQuyDinh", thueSuatQuyDinh);
            writer.WriteElementString("thueSuatSauGiam", thueSuatSauGiam);
            writer.WriteElementString("thueGTGTDuocGiam", thueGTGTDuocGiam);
            writer.WriteEndElement();
        }

        private void AddCKyDTu(XmlWriter writer)
        {
            writer.WriteStartElement("CKyDTu");

            // Chữ ký số - SỬA LẠI PHẦN NAMESPACE CHO CHỮ KÝ
            writer.WriteStartElement("ds", "Signature", "http://www.w3.org/2000/09/xmldsig#");

            writer.WriteStartElement("ds", "SignedInfo", "http://www.w3.org/2000/09/xmldsig#");

            writer.WriteStartElement("ds", "CanonicalizationMethod", "http://www.w3.org/2000/09/xmldsig#");
            writer.WriteAttributeString("Algorithm", "http://www.w3.org/TR/2001/REC-xml-c14n-20010315#WithComments");
            writer.WriteEndElement();

            writer.WriteStartElement("ds", "SignatureMethod", "http://www.w3.org/2000/09/xmldsig#");
            writer.WriteAttributeString("Algorithm", "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256");
            writer.WriteEndElement();

            writer.WriteStartElement("ds", "Reference", "http://www.w3.org/2000/09/xmldsig#");
            writer.WriteAttributeString("URI", "#ID_1");

            writer.WriteStartElement("ds", "Transforms", "http://www.w3.org/2000/09/xmldsig#");
            writer.WriteStartElement("ds", "Transform", "http://www.w3.org/2000/09/xmldsig#");
            writer.WriteAttributeString("Algorithm", "http://www.w3.org/2000/09/xmldsig#enveloped-signature");
            writer.WriteEndElement();
            writer.WriteEndElement(); // Transforms

            writer.WriteStartElement("ds", "DigestMethod", "http://www.w3.org/2000/09/xmldsig#");
            writer.WriteAttributeString("Algorithm", "http://www.w3.org/2001/04/xmlenc#sha256");
            writer.WriteEndElement();

            writer.WriteElementString("ds", "DigestValue", "http://www.w3.org/2000/09/xmldsig#", "ousU0zaI5UKWzUItwzX+VsRdKVvz7R/AU5MEfW01VPY=");
            writer.WriteEndElement(); // Reference
            writer.WriteEndElement(); // SignedInfo

            writer.WriteElementString("ds", "SignatureValue", "http://www.w3.org/2000/09/xmldsig#",
                "XjvL594mUkw7CEAkADjU2lgHA/JumGDvsidHFGgpTVHFwqtuUqkpqVhrLzNMxSH4BNK/md0DppEe\r\n" +
                "9TND0/TGHPQSmwu6nSrJtvmdWOU8uQzFVlV1kseRrLWSraQnXxDl5KanQA8Auc+m6pgGkNcXQCn/\r\n" +
                "tivWUULKu8LexSuPbZE=");

            writer.WriteStartElement("ds", "KeyInfo", "http://www.w3.org/2000/09/xmldsig#");

            writer.WriteStartElement("ds", "KeyValue", "http://www.w3.org/2000/09/xmldsig#");
            writer.WriteStartElement("ds", "RSAKeyValue", "http://www.w3.org/2000/09/xmldsig#");
            writer.WriteElementString("ds", "Modulus", "http://www.w3.org/2000/09/xmldsig#",
                "yuHYtFejkKHvgv6dCjXqcHniEBOX+GPihrDsL5uMxnUeYYb29KMQz9wv+oi0xkm6kuGFJYgEOTCw\r\n" +
                "0FVA4N3ejg8BnxzmKQ506IdbNiHy7L1rD69w90oL+o3lHM5/DrhFY+CDJMUNNGjQ9rEbCinwLHMu\r\n" +
                "Ml8NE4rmgZCSgMO/REs=");
            writer.WriteElementString("ds", "Exponent", "http://www.w3.org/2000/09/xmldsig#", "AQAB");
            writer.WriteEndElement(); // RSAKeyValue
            writer.WriteEndElement(); // KeyValue

            writer.WriteStartElement("ds", "X509Data", "http://www.w3.org/2000/09/xmldsig#");
            writer.WriteElementString("ds", "X509SubjectName", "http://www.w3.org/2000/09/xmldsig#",
                "UID=MST:3500779171,CN=CÔNG TY TRÁCH NHIỆM HỮU HẠN THƯƠNG MẠI  VÀ XÂY DỰNG ĐẠI THÀNH CÔNG,L=BÀ RỊA VŨNG TÀU,C=VN");
            writer.WriteElementString("ds", "X509Certificate", "http://www.w3.org/2000/09/xmldsig#",
                "MIIEjzCCA3egAwIBAgIQVAT//rcDP7MktyjTwuka9jANBgkqhkiG9w0BAQsFADA/MRgwFgYDVQQD\r\n" +
                "DA9WaWV0dGVsLUNBIFNIQTIxFjAUBgNVBAoMDVZpZXR0ZWwgR3JvdXAxCzAJBgNVBAYTAlZOMB4X\r\n" +
                "DTI1MDgxOTAyMDAwMFoXDTI8MDgxOTAyMDAwMFowga4xCzAJBgNVBAYTAlZOMR0wGwYDVQQHDBRC\r\n" +
                "w4AgUuG7ikEgVsWoTkcgVMOAVTFgMF4GA1UEAwxXQ8OUTkcgVFkgVFLDgUNIIE5ISeG7hk0gSOG7\r\n" +
                "rlUgSOG6oE4gVEjGr8agTkcgTeG6oEkgIFbDgCBYw4JZIEThu7BORyDEkOG6oEkgVEjDgE5IIEPD\r\n" +
                "lE5HMR4wHAYKCZImiZPyLGQBAQwOTVNUOjM1MDA3NzkxNzEwgZ8wDQYJKoZIhvcNAQEBBQADgY0A\r\n" +
                "MIGJAoGBAMrh2LRXo5Ch74L+nQo16nB54hATl/hj4oaw7C+bjMZ1HmGG9vSjEM/cL/qItMZJupLh\r\n" +
                "hSWIBDkwsNBVQODd3o4PAZ8c5ikOdOiHWzYh8uy9aw+vcPdKC/qN5RzOfw64RWPggyTFDTRo0Pax\r\n" +
                "Gwop8CxzLjJfDROK5oGQkoDDv0RLAgMBAAGjggGZMIIBlTAMBgNVHRMBAf8EAjAAMB8GA1UdIwQY\r\n" +
                "MBaAFEPVNQCLvge6403mHiRZVohbvsxKMHkGCCsGAQUFBwEBBG0wazBCBggrBgEFBQcwAoY2aHR0\r\n" +
                "cDovL3ZpZXR0ZWwtY2Eudm4vZG93bmxvYWRsL3N1Yi9WaWV0dGVsLUNBX1NIQTIuY3J0MCUGCCsG\r\n" +
                "AQUFBzABhhlodHRwOi8vb2NzcC52aWV0dGVsLWNhLnZuMDMGA1UdJQQsMCoGCCsGAQUFBwMCBggr\r\n" +
                "BgEFBQcDBAYKKwYBBAGCNwoDDAYIKwYBBQUHAyQwgYQGA1UdHwR9MHsweaAyoDCGLmh0dHA6Ly9j\r\n" +
                "cmwudmlldHRlbC1jYS52bi9WaWV0dGVsLUNBLVNIQTItMi5jcmyiQ6RBMD8xGDAWBgNVBAMMD1Zp\r\n" +
                "ZXR0ZWwtQ0EgU0hBMjEWMBQGA1UECgwNVmlldHRlbCBHcm91cDELMAkGA1UEBhMCVk4wHQYDVR0O\r\n" +
                "BBYEFPHzoANwykPcDgH/MI4AZnvNvpOfMA4GA1UdDwEB/wQEAwIF4DANBgkqhkiG9w0BAQsFAAOC\r\n" +
                "AQEAORpmvSpHV6rc5NEwcG04SHavVzC5HmbOdTYRH2kWnLIeFrWuEBsuhnW7dtJ/Prsd3CjPdenq\r\n" +
                "AhptuXzzP8HIj6FLHJsymvGUeOWo4sGJYU8knqhN+KxgRMFaTEG9QZAf1auR6Iw0aqe6TwEKAqTm\r\n" +
                "tkL8gGvysJqZX2i6c4MNdvSRaVfoc0bSUAwulLTrrIyqZy4xUT8Sg9vA79rsuMSRKX3/yl3Exhvm\r\n" +
                "k1Jh0nkSypsYvh8o+vRZuQvUpgsoF78Guy6z7/LDSu1ypAaMBKhM1bN0MeDevnvjMwDP/8RX2y13\r\n" +
                "krCJuGDAdqz/O+iCWLy1gnepPbx8SSVc1z2rYj/DVA==");
            writer.WriteEndElement(); // X509Data
            writer.WriteEndElement(); // KeyInfo
            writer.WriteEndElement(); // Signature

            writer.WriteEndElement(); // CKyDTu
        }

        [HttpPost]
        public void ExportXML(string TenCty)
        {

        }
        string dbPath = "";
        public class Phuluc1
        {
            public int STT { get; set; }
            public string Tenhang { get; set; }
            public double TTrcthue { get; set; }
            public double TThue { get; set; }
        }
        List<Phuluc1> lstPhuluc1=new List<Phuluc1>();
        List<Phuluc1> lstPhuluc2 = new List<Phuluc1>();
        public ActionResult Index(string path, string ky)
        {
            if (!string.IsNullOrEmpty(ky))
            {
                if (ky.Contains("T"))
                {
                    ViewBag.ky = $"Tháng {ky.Replace("T", "")} năm {DateTime.Now.Year}";
                }
                if (ky.Contains("Q"))
                {
                    ViewBag.ky = $"Quý {ky.Replace("Q", "")} năm {DateTime.Now.Year}";
                }
            }
            else
            {
                ViewBag.ky = "";
            }
            
                try
                {
                    Noidungtax model = new Noidungtax();
                    if (!string.IsNullOrEmpty(path))
                    {
                        dbPath = path;
                        string query = "SELECT * FROM License";
                        DataTable tbLicence = ExecuteQuery(query, null);
                        if (tbLicence.Rows.Count > 0)
                        {
                            model.TenCty = Helpers.ConvertVniToUnicode(tbLicence.Rows[0].Field<string>("TenCty"));
                            model.Mst = tbLicence.Rows[0].Field<string>("MaSoThue");

                        }
                        query = "SELECT * FROM ToKhaiThue";
                        DataTable ToKhaiThue = ExecuteQuery(query, null);
                    //Kiểm tra xem có chốt chưa, nếu chưa thì lấy từ tờ khai thuế
                    query = @"SELECT * FROM tbThongTinToKhai   WHERE  Nam= ?";

                    var parameterss = new OleDbParameter[]
                    {
                 new OleDbParameter("?", DateTime.Now.Year.ToString())
                    };
                    var kq = ExecuteQuery(query, parameterss);
                    string contentXMl = "";
                    if (kq.Rows.Count > 0)
                    {
                        model.Nguoiky = kq.Rows[0].Field<string>("NguoiKy");
                        if (ky == "Q1")
                        {
                            string khoa1 = kq.Rows[0].Field<string>("k1");
                            if (khoa1 == "1")
                            {
                                ViewBag.khoa = 1;
                                contentXMl = kq.Rows[0].Field<string>("xml1");
                            }
                        }
                        if (ky == "Q2")
                        {
                            string khoa1 = kq.Rows[0].Field<string>("k2");
                            if (khoa1 == "1")
                            {
                                ViewBag.khoa = 1;
                                contentXMl = kq.Rows[0].Field<string>("xml2");
                            }
                        }
                        if (ky == "Q3")
                        {
                            string khoa1 = kq.Rows[0].Field<string>("k3");
                            if (khoa1 == "1")
                            {
                                ViewBag.khoa = 1;
                                contentXMl = kq.Rows[0].Field<string>("xml3");
                            }
                        }
                        if (ky == "Q4")
                        {
                            string khoa1 = kq.Rows[0].Field<string>("k4");
                            if (khoa1 == "1")
                            {
                                ViewBag.khoa = 1;
                                contentXMl = kq.Rows[0].Field<string>("xml4");
                            }
                        }
                        //Kiểm tra theo tháng
                        for (int i = 1; i <= 12; i++)
                        {
                            string monthKey = "T" + i; // T1, T2, ..., T12
                            string khoaKey = "k" + i;   // k1, k2, ..., k12
                            string xmlKey = "xml" + i;   // xml1, xml2, ..., xml12

                            if (ky == monthKey)
                            {
                                string khoaValue = kq.Rows[0].Field<string>(khoaKey);
                                if (khoaValue == "1")
                                {
                                    ViewBag.khoa = 1; // Lưu số tháng
                                    contentXMl = kq.Rows[0].Field<string>(xmlKey); // Lưu nội dung XML
                                    break; // Thoát khỏi vòng lặp nếu đã tìm thấy
                                }
                            }
                        }
                    }

                    if (contentXMl == "")
                    {
                        if (ToKhaiThue.Rows.Count > 0)
                        {
                            model.N22 = ToKhaiThue.Rows[0].Field<double>("N11");
                            model.N23 = ToKhaiThue.Rows[0].Field<double>("N12");
                            model.N24 = ToKhaiThue.Rows[0].Field<double>("N13");
                            model.N25 = ToKhaiThue.Rows[0].Field<double>("N13");
                            model.N29 = ToKhaiThue.Rows[0].Field<double>("N29");
                            model.N30 = ToKhaiThue.Rows[0].Field<double>("N30");
                            model.N31 = ToKhaiThue.Rows[0].Field<double>("N31");
                            model.N32 = ToKhaiThue.Rows[0].Field<double>("N32");
                            model.N33 = ToKhaiThue.Rows[0].Field<double>("N33");
                        }
                    }
                    else
                    {
                        //Đọc file xml
                        model.N22 = double.Parse(GetValue(contentXMl, "ct22").ToString());
                        model.N23 = double.Parse(GetValue(contentXMl, "ct23").ToString());
                        model.N24 = double.Parse(GetValue(contentXMl, "ct24").ToString());
                        model.N25 = double.Parse(GetValue(contentXMl, "ct25").ToString()); 
                        model.N29 = double.Parse(GetValue(contentXMl, "ct29").ToString());
                        model.N30 = double.Parse(GetValue(contentXMl, "ct30").ToString());
                        model.N31 = double.Parse(GetValue(contentXMl, "ct31").ToString());
                        model.N32 = double.Parse(GetValue(contentXMl, "ct32").ToString());
                        model.N33 = double.Parse(GetValue(contentXMl, "ct33").ToString());
                    }


                    //Load cho số dư trước
                    if (1 < 2 && (kq.Rows.Count > 0))
                    {
                        if (ky.Contains("Q"))
                        {
                            //Lấy từ năm trước
                            if (ky == "Q1")
                            {

                            }
                            //Lấy từ Q1
                            if (ky == "Q2")
                            {
                                model.N22 = kq.Rows[0].Field<double>("Quy1");
                            }
                            //Lấy từ Q2
                            if (ky == "Q3")
                            {
                                model.N22 = kq.Rows[0].Field<double>("Quy2");
                            }
                            //Lấy từ Q3
                            if (ky == "Q4")
                            {
                                model.N22 = kq.Rows[0].Field<double>("Quy3");
                            }
                        }
                        if (ky.Contains("T"))
                        {
                            if (ky == "T1")
                            {
                            }
                            if (ky == "T2")
                            {
                                model.N22 = kq.Rows[0].Field<double>("T1");
                            }
                            if (ky == "T3")
                            {
                                model.N22 = kq.Rows[0].Field<double>("T2");
                            }
                            if (ky == "T4")
                            {
                                model.N22 = kq.Rows[0].Field<double>("T3");
                            }
                            if (ky == "T5")
                            {
                                model.N22 = kq.Rows[0].Field<double>("T4");
                            }
                            if (ky == "T6")
                            {
                                model.N22 = kq.Rows[0].Field<double>("T5");
                            }
                            if (ky == "T7")
                            {
                                model.N22 = kq.Rows[0].Field<double>("T6");
                            }
                            if (ky == "T8")
                            {
                                model.N22 = kq.Rows[0].Field<double>("T7");
                            }
                            if (ky == "T9")
                            {
                                model.N22 = kq.Rows[0].Field<double>("T8");
                            }
                            if (ky == "T10")
                            {
                                model.N22 = kq.Rows[0].Field<double>("T9");
                            }
                            if (ky == "T11")
                            {
                                model.N22 = kq.Rows[0].Field<double>("T10");
                            }
                            if (ky == "T12")
                            {
                                model.N22 = kq.Rows[0].Field<double>("T11");
                            }
                        }
                    }

                    }
                    else
                    {
                        //return Redirect("Contact");
                    }

                //Lấy danh sách hoá đơn
                string sql = @"
SELECT DISTINCTROW 
    KyHieu,
    SoHD,
    ChungTu.NgayCT as NgayPH,
    MatHang,
    SoLuong,
    ThanhTien,
    KhachHang.Ten,
    KhachHang.MST,
    ChungTu.SoHieu,
    SoPS,
    KhachHang.DiaChi,
    TyLe,
    HTTT,
    MauSo,
    MaCT,
    HoaDon.MaSo,
    KCT 
FROM  
    (HoaDon 
    INNER JOIN ChungTu ON HoaDon.MaSo = ChungTu.MaSo) 
    LEFT JOIN KhachHang ON HoaDon.MaKhachHang = KhachHang.MaSo  
WHERE 
    Loai = -1 
    AND HD = 1 
    AND ThangCT >= 7 
    AND ThangCT <= 9
    AND (HDBL = 0 OR KCT = 0) 
    AND (HoaDon.DC = 0 OR HD = 1)
    AND TyLe = 8 
ORDER BY 
    NgayPH,
    MaCT";
                DataTable kqdv = ExecuteQuery(sql, null);
                foreach (DataRow item in kqdv.Rows)
                {
                    //Tìm dòng chung tu
                    Phuluc1 Phuluc1 = new Phuluc1();
                    Phuluc1.Tenhang = Helpers.ConvertVniToUnicode(item.Field<string>("MatHang"));
                    int step = 1;
                    Phuluc1.TTrcthue = item.Field<double>("ThanhTien");
                    Phuluc1.TThue = item.Field<double>("SoPs");

                    lstPhuluc1.Add(Phuluc1);
                }
                var result = lstPhuluc1
               .GroupBy(i => i.Tenhang)
               .Select(g => new Phuluc1
               {
                   Tenhang = g.Key,
                   TTrcthue = g.Sum(x => x.TTrcthue),
                   TThue = g.Sum(x => x.TThue)
               })
               .ToList();
                    ViewBag.Phuluc1 = result;
                ViewBag.Sum1 = result.Sum(m => m.TTrcthue);
                ViewBag.Sum2 = result.Sum(m => m.TThue);
                //Tính đầu ra
                sql = @"SELECT DISTINCTROW HoaDon.KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=3007,SoPS,-SoPS) AS Thue,ChungTu.MauSoHD as DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM  ((HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo) LEFT JOIN KhachHang ON HoaDon.MaKhachHang=KhachHang.MaSo  WHERE HoaDon.Loai=1 AND  (ThangCT>=7 AND ThangCT<=9)  AND (HoaDon.DC=0 OR HD=1) and TyLe=8  ORDER BY NgayPH";
                DataTable kqdr = ExecuteQuery(sql, null);
                foreach (DataRow item in kqdr.Rows)
                {

                    //Tìm dòng chung tu
                    Phuluc1 Phuluc1 = new Phuluc1();
                    Phuluc1.Tenhang = Helpers.ConvertVniToUnicode(item.Field<string>("MatHang"));
                    int step = 1;
                    Phuluc1.TTrcthue = item.Field<double>("ThanhTien"); 
                    lstPhuluc2.Add(Phuluc1);
                }
                result = lstPhuluc2
              .GroupBy(i => i.Tenhang)
              .Select(g => new Phuluc1
              {
                  Tenhang = g.Key,
                  TTrcthue = g.Sum(x => x.TTrcthue),
                  TThue = g.Sum(x => x.TTrcthue) * 0.02f
              })
              .ToList();
                    ViewBag.Phuluc2 = result;
                decimal s3 = (decimal)result.Sum(m => m.TTrcthue);
                ViewBag.Sum3=s3;
                ViewBag.Sum4 = Math.Round(s3 * 0.02m);
                return View(model);
                }


                catch (Exception ex)
                {
                    //return Redirect("Contact");
                    // throw ex;
                }
                
            return View();

        }
        static string GetValue(string xml, string tagName)
        {
            // Tạo chuỗi tìm kiếm cho thẻ mở và thẻ đóng
            string startTag = $"<{tagName}>";
            string endTag = $"</{tagName}>";

            // Tìm vị trí của thẻ mở
            int startIndex = xml.IndexOf(startTag);
            if (startIndex == -1) return null; // Thẻ không tồn tại

            // Tìm vị trí của thẻ đóng
            int endIndex = xml.IndexOf(endTag, startIndex);
            if (endIndex == -1) return null; // Thẻ không tồn tại

            // Lấy giá trị giữa thẻ mở và thẻ đóng
            startIndex += startTag.Length; // Chuyển đến vị trí sau thẻ mở
            return xml.Substring(startIndex, endIndex - startIndex);
        }
        public static string GetInvoiceQuery(int year, int quarter)
        {
            DateTime startDate;
            DateTime endDate;

            switch (quarter)
            {
                case 1:
                    startDate = new DateTime(year, 1, 1);
                    endDate = new DateTime(year, 3, 31);
                    break;
                case 2:
                    startDate = new DateTime(year, 4, 1);
                    endDate = new DateTime(year, 6, 30);
                    break;
                case 3:
                    startDate = new DateTime(year, 7, 1);
                    endDate = new DateTime(year, 9, 30);
                    break;
                case 4:
                    startDate = new DateTime(year, 10, 1);
                    endDate = new DateTime(year, 12, 31);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(quarter), "Quý phải trong khoảng từ 1 đến 4.");
            }

            return $"SELECT * FROM HoaDon WHERE ngayph >= #{startDate:yyyy-MM-dd}# AND ngayph <= #{endDate:yyyy-MM-dd}# AND TyLe = 8;";
        }
        public System.Data.DataTable ExecuteQuery(string query, params OleDbParameter[] parameters)
        {

            string password = "1@35^7*9)1";
            connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";

            System.Data.DataTable dataTable = new System.Data.DataTable();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        // Thêm các tham số vào command
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }   

                        using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command))
                        {
                            dataAdapter.Fill(dataTable);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }

            }

            return dataTable; // Trả về DataTable chứa dữ liệu
        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}