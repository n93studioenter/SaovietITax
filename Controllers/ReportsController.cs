using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Taxweb.Models;

namespace Taxweb.Controllers
{
    public class ReportsController : Controller
    {
        // GET: Reports
        public static string dbPath = "";
        string password, connectionString;
        //Type 1:Vào , 2: ra
        public ActionResult Bangkehoadon(string path,int ? type)
        {
            dbPath = path;
            ViewBag.path = dbPath;
            OleDbConnection conn = null;
            string password = "1@35^7*9)1";
            connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
            conn = new OleDbConnection(connectionString);
            conn.Open();

            if (!string.IsNullOrEmpty(path))
            {
                string sql = @"SELECT DISTINCTROW KyHieu, SoHD, ChungTu.NgayCT as NgayPH, MatHang, SoLuong, ThanhTien,
                      KhachHang.Ten, KhachHang.MST, ChungTu.SoHieu, SoPS, KhachHang.DiaChi,
                      TyLe, HTTT, MauSo, MaCT, HoaDon.MaSo, KCT
               FROM (HoaDon 
                     INNER JOIN ChungTu ON HoaDon.MaSo = ChungTu.MaSo)
                     LEFT JOIN KhachHang ON HoaDon.MaKhachHang = KhachHang.MaSo
               WHERE Loai = -1 
                     AND HD = 1 
                     AND (ThangCT >= ? AND ThangCT <= ?)
                     AND (HDBL = 0 OR KCT = 0)
                     AND (HoaDon.DC = 0 OR HD = 1)
               ORDER BY NgayPH, MaCT";

                string sqlra = @"SELECT DISTINCTROW 
                    HoaDon.KyHieu,
                    SoHD,
                    ChungTu.NgayCT AS NgayPH,
                    MatHang,
                    SoLuong,
                    ThanhTien,
                    KhachHang.Ten,
                    KhachHang.MST,
                    ChungTu.SoHieu,
                    IIF(HethongTK.TK_ID = 3007, SoPS, -SoPS) AS Thue,
                    ChungTu.MauSoHD AS DiaChi,
                    TyLe,
                    HTTT,
                    MauSo,
                    MaCT,
                    KCT
               FROM ((HoaDon 
                      INNER JOIN ChungTu 
                          ON HoaDon.MaSo = ChungTu.MaSo)
                     LEFT JOIN HethongTK 
                          ON ChungTu.MaTKCo = HethongTK.MaSo)
                     LEFT JOIN KhachHang 
                          ON HoaDon.MaKhachHang = KhachHang.MaSo
               WHERE HoaDon.Loai = 1
                     AND (ThangCT >= ? AND ThangCT <= ?)
                     AND (HoaDon.DC = 0 OR HD = 1)
               ORDER BY NgayPH";
                var parameters = new OleDbParameter[]
             {
                        new OleDbParameter("?",1),
                        new OleDbParameter("?",12), 
             };
                DataTable data = ExecuteQuery(sql, parameters);

              var model= data.AsEnumerable().Select(r => new Bangkehoadon
                {
                  KyHieu = r.Field<string>("KyHieu"),
                    SoHD = r.Field<string>("SoHD"), 
                    NgayPH = DateTime.Parse(r["NgayPH"].ToString()),
                    MatHang = Helpers.ConvertVniToUnicode(r.Field<string>("MatHang")),
                    SoLuong = int.Parse(r["SoLuong"].ToString() ),
                    ThanhTien = double.Parse(r["ThanhTien"].ToString()),
                    Ten = Helpers.ConvertVniToUnicode(r.Field<string>("Ten")),
                    MST = r.Field<string>("MST"),
                    SoHieu = r.Field<string>("SoHieu"),
                    SoPS = double.Parse(r["SoPS"].ToString()),
                    DiaChi = Helpers.ConvertVniToUnicode(r.Field<string>("DiaChi")),
                    TyLe = int.Parse(r["TyLe"].ToString()),
                    HTTT = r.Field<string>("HTTT"),
                    MauSo = int.Parse(r["MauSo"].ToString()),
                    MaCT = int.Parse(r["MaCT"].ToString()),
                    KCT = int.Parse(r["KCT"].ToString()), 

              }).ToList();
                ViewBag.Loai1 = model.Where(m => m.TyLe == 0 && m.HTTT != "5").ToList();
                ViewBag.Loai2 = model.Where(m => m.TyLe ==5 ).ToList();
                ViewBag.Loai3 = model.Where(m => m.TyLe == 8 ).ToList();
                ViewBag.Loai4 = model.Where(m => m.TyLe == 10 ).ToList();
                ViewBag.Loai5 = model.Where(m => m.TyLe == 0 && m.HTTT == "5").ToList();
                return View(model);
            }
            return View();
        }

        public int ExecuteQueryResult(string query, params OleDbParameter[] parameters)
        {
            string password = "1@35^7*9)1";
            connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                Console.WriteLine("Kết nối đến cơ sở dữ liệu thành công!");

                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    // Thêm tham số
                    if (parameters != null)
                        command.Parameters.AddRange(parameters);

                    // Thực thi INSERT, UPDATE, DELETE
                    command.ExecuteNonQuery();
                }

                // Lấy ID vừa thêm bằng @@IDENTITY
                using (OleDbCommand idCommand = new OleDbCommand("SELECT @@IDENTITY", connection))
                {
                    object result = idCommand.ExecuteScalar();
                    return Convert.ToInt32(result);
                }
            }
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
    }
}