using Infragistics.Windows.DataPresenter;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Linq;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Windows;

namespace Incd1
{
    public class StartUp : StartupBase
    {
        public static DataSet DataSourceReport = new DataSet();
        public static int m_Tinh_dc = 1;
        private static int KindStyleReport = -1;
        private static int KindPrint = -1;
        private static string Condition = string.Empty;
        private static string StartDate = string.Empty;
        private static string EndDate = string.Empty;
        public static string TableList = "v_INCD1";
        private static SqlCommand cmd = new SqlCommand();
        public static DataTable tbDetail = (DataTable)null;
        public static int M_ROUND_GIA = 0;
        public static int M_ROUND_GIA_NT = 0;
        public static DataTable DtGroupInfo = (DataTable)null;
        public static DataTable GroupSelected = (DataTable)null;
        private static string SumFields = "ton_dau;du_dau;du_dau_nt;sl_nhap;tien_nhap;tien_nt_n;sl_xuat;tien_xuat;tien_nt_x;ton_cuoi;du_cuoi;du_cuoi_nt";
        private static FrmLoc _frmSearch;
        private static SasFormBrowes.FormBrowse oBrowse;
        public static DataRow CommandInfo;
        public static DataRow CommandInfo0;
        public static DateTime M_ngay_ct0;
        public static string M_ma_nt0;
        public static string[] parameters;

        public override void Run()
        {
            StartupBase.Namespace = "Incd1";
            try
            {
                StartUp.CommandInfo = StartUp.CommandInfo0 = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                StartUp.M_ngay_ct0 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_KY1");
                StartUp.M_ROUND_GIA = (int)StartupBase.SasObj.GetSysvar("M_ROUND_GIA");
                StartUp.M_ROUND_GIA_NT = (int)StartupBase.SasObj.GetSysvar("M_ROUND_GIA_NT");
                StartUp.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString(); 
                if (StartUp.CommandInfo == null)
                    return;
                StartUp._frmSearch = new FrmLoc();
                StartUp._frmSearch.Title = StartupBase.M_LAN.Equals("V") ? SysFunc.Cat_Dau(StartUp.CommandInfo["bar"].ToString()) : SysFunc.Cat_Dau(StartUp.CommandInfo["bar2"].ToString());
                string menu_id = StartUp.CommandInfo0["menu_copy"].ToString().Trim();
                if (menu_id != "")
                {
                    StartUp.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, menu_id);
                    if (StartUp.CommandInfo == null)
                    {
                        StartUp.CommandInfo = StartUp.CommandInfo0;
                    }
                    else
                    {
                        StartUp.CommandInfo["rep_file"] = StartUp.CommandInfo0["rep_file"];
                        StartUp.CommandInfo["bar"] = StartUp.CommandInfo0["bar"];
                        StartUp.CommandInfo["bar2"] = StartUp.CommandInfo0["bar2"];
                        StartUp.CommandInfo["title"] = StartUp.CommandInfo0["title"];
                        StartUp.CommandInfo["title2"] = StartUp.CommandInfo0["title2"];
                    }
                }
                StartUp.parameters = StartUp.CommandInfo["parameter"].ToString().Trim().Split(';');
                StartUp._frmSearch.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public static void QueryData(
          bool isFirstLoad,
          object SDate,
          object EDate,
          string filter,
          string filterSD,
          string vttonkho,
          bool KReport,
          int KPrint)
        {
            try
            {
                if (isFirstLoad)
                {
                    StartUp.KindStyleReport = KReport ? 1 : 2;
                    StartUp.KindPrint = KPrint;
                    StartUp.Condition = filter;
                    StartUp.StartDate = string.IsNullOrEmpty(SDate.ToString()) ? "" : string.Format("{0:yyyyMMdd}", (object)(DateTime)SDate);
                    StartUp.EndDate = string.IsNullOrEmpty(EDate.ToString()) ? "" : string.Format("{0:yyyyMMdd}", (object)(DateTime)EDate);
                    StartUp.cmd.CommandText = "Exec " + StartUp.CommandInfo["store_proc"].ToString().Split('|')[0] + " @StartDate , @EndDate , @Tinh_dc, @Condition,@vttonkho,@ConditionSD";
                    StartUp.cmd.Parameters.Add("@StartDate", SqlDbType.VarChar).Value = (object)StartUp.StartDate;
                    StartUp.cmd.Parameters.Add("@EndDate", SqlDbType.VarChar).Value = (object)StartUp.EndDate;
                    StartUp.cmd.Parameters.Add("@Tinh_dc", SqlDbType.Int).Value = (object)StartUp.m_Tinh_dc;
                    StartUp.cmd.Parameters.Add("@Condition", SqlDbType.NVarChar).Value = (object)StartUp.Condition;
                    StartUp.cmd.Parameters.Add("@vttonkho", SqlDbType.Char, 1).Value = (object)vttonkho;
                    StartUp.cmd.Parameters.Add("@ConditionSD", SqlDbType.NVarChar).Value = (object)filterSD;
                    StartUp.tbDetail = StartupBase.SasObj.ExcuteReader(StartUp.cmd).Tables[0].Copy();
                    StartUp.tbDetail.TableName = "tbDetail";
                    StartUp.DataSourceReport.Tables.Add(StartUp.tbDetail);
                    StartUp.oBrowse = new SasFormBrowes.FormBrowse(StartupBase.SasObj, StartUp.tbDetail.DefaultView, StartUp.GetTableShow(StartUp.KindStyleReport, StartUp.KindPrint));
                    StartUp.oBrowse.F7 += new SasFormBrowes.FormBrowse.GridKeyUp_F7(StartUp.oBrowse_F7);
                    StartUp.oBrowse.F5 += new SasFormBrowes.FormBrowse.GridKeyUp_F5(StartUp.oBrowse_F5);
                    StartUp.oBrowse.CTRL_R += new SasFormBrowes.FormBrowse.GridKeyUp_CTRL_R(StartUp.oBrowse_CTRL_R);
                    StartUp.oBrowse.F11 += new SasFormBrowes.FormBrowse.GridKeyUp_F11(StartUp.oBrowse_F11);
                    StartUp.oBrowse.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);
                    StartUp.oBrowse.frmBrw.Title = StartupBase.M_LAN.Equals("V") ? SysFunc.Cat_Dau(StartUp.CommandInfo["bar"].ToString()) : SysFunc.Cat_Dau(StartUp.CommandInfo["bar2"].ToString());
                    
                }
                else
                {
                    StartUp.DataSourceReport.Tables.Remove("tbDetail");
                    StartUp.tbDetail = StartupBase.SasObj.ExcuteReader(StartUp.cmd).Tables[0].Copy();
                    StartUp.tbDetail.TableName = "tbDetail";
                    StartUp.DataSourceReport.Tables.Add(StartUp.tbDetail);
                    StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)StartUp.tbDetail.DefaultView;
                    StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
                    StartUp.oBrowse.UpdateSumaryFields();
                }
                if (!isFirstLoad)
                    return;
                //FieldSortDescription sfg1 = new FieldSortDescription();
                //sfg1.FieldName = "dvt";
                //sfg1.IsGroupBy = true;

                //StartUp.oBrowse.DataGrid.FieldLayouts[0].SortedFields.Add(sfg1);

                StartUp.oBrowse.frmBrw.LanguageID = "Incd1_1";
                StartUp.oBrowse.ShowDialog();
                StartUp._frmSearch.Close();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private static void oBrowse_CTRL_R(object sender, EventArgs e)
        {
            bool KReport = true;
            if (StartUp._frmSearch.cbmau_bc.Value.ToString().Equals("1"))
                KReport = false;
            string filterSD;
            string filter = StartUp._frmSearch.GetFilter(out filterSD);
            StartUp.QueryData(false, StartUp._frmSearch.TxtStartDateTime.Value, StartUp._frmSearch.TxtEndDateTime.Value, filter, filterSD, StartUp._frmSearch.txtVtTonkho.Text, KReport, 1);
        }

        private static void oBrowse_F5(object sender, EventArgs e)
        {
            try
            {
                if (StartUp.oBrowse.ActiveRecord == null || StartUp.oBrowse.ActiveRecord.RecordType != RecordType.DataRecord)
                    return;
                string str = StartUp.oBrowse.ActiveRecord.Cells["ma_vt"].Value.ToString();
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = "Exec " + StartUp.CommandInfo["store_proc"].ToString().Split('|')[1] + " @StartDate , @EndDate , @Tinh_dc, @Condition";
                sqlcmd.Parameters.Add("@StartDate", SqlDbType.VarChar).Value = (object)StartUp.StartDate;
                sqlcmd.Parameters.Add("@EndDate", SqlDbType.VarChar).Value = (object)StartUp.EndDate;
                sqlcmd.Parameters.Add("@Tinh_dc", SqlDbType.Int).Value = (object)StartUp.m_Tinh_dc;
                sqlcmd.Parameters.Add("@Condition", SqlDbType.NVarChar).Value = (object)(StartUp.Condition + " and ma_vt='" + str.Replace("'", "''") + "'");
                if (StartUp.CommandInfo["store_proc"].ToString().Split('|')[1].ToUpper().Equals("INCD1B_DETAIL"))
                {
                    sqlcmd.CommandText += ", @ConditionQD";
                    sqlcmd.Parameters.Add("@ConditionQD", SqlDbType.NVarChar).Value = (object)("dvt1=N'" + StartUp.oBrowse.ActiveRecord.Cells["dvt"].Value.ToString().Replace("'", "''") + "'");
                }
                DataTable dataTable = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Copy();
                SasFormBrowes.FormBrowse formBrowse = new SasFormBrowes.FormBrowse(StartupBase.SasObj, dataTable.DefaultView, StartUp.GetTable4DetailShow(StartUp.KindStyleReport, StartUp.KindPrint));
                formBrowse.Esc += new SasFormBrowes.FormBrowse.GridKeyUp_Esc(StartUp.FormBrowse_Esc);
                formBrowse.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);
                formBrowse.frmBrw.Title = SysFunc.Cat_Dau((StartupBase.M_LAN.Equals("V") ? "Chi tiết phát sinh của vật tư: " : "Details arising of the item: ") + StartUp.oBrowse.ActiveRecord.Cells["ma_vt"].Value.ToString().Trim() + " - " + (StartupBase.M_LAN.Equals("V") ? StartUp.oBrowse.ActiveRecord.Cells["ten_vt"].Value.ToString().Trim() : StartUp.oBrowse.ActiveRecord.Cells["ten_vt2"].Value.ToString().Trim()));
                formBrowse.AddValueSummary(new string[1]
                {
                    StartupBase.M_LAN.Equals("V") ? "ten_kh" : "ten_kh2"
                }, new string[1]
                {
                    StartupBase.M_LAN.Equals("V") ? "Tồn đầu kỳ:" : "Opening balance:"
                });
                formBrowse.AddValueSummary(new string[1]
                {
                    StartupBase.M_LAN.Equals("V") ? "ten_kh" : "ten_kh2"
                }, new string[1]
                {
                    StartupBase.M_LAN.Equals("V") ? "Nhập trong kỳ:" : "Receipt amount:"
                });
                formBrowse.AddValueSummary(new string[1]
                {
                    StartupBase.M_LAN.Equals("V") ? "ten_kh" : "ten_kh2"
                }, new string[1]
                {
                    StartupBase.M_LAN.Equals("V") ? "Xuất trong kỳ:" : "Issue amount:"
                });
                formBrowse.AddValueSummary(new string[1]
                {
                    StartupBase.M_LAN.Equals("V") ? "ten_kh" : "ten_kh2"
                }, new string[1]
                {
                    StartupBase.M_LAN.Equals("V") ? "Tồn cuối kỳ:" : "Closing balance:"
                });
                Decimal result1;
                Decimal.TryParse(StartUp.oBrowse.ActiveRecord.Cells["ton_dau"].Value.ToString(), out result1);
                Decimal result2;
                Decimal.TryParse(StartUp.oBrowse.ActiveRecord.Cells["sl_nhap"].Value.ToString(), out result2);
                Decimal result3;
                Decimal.TryParse(StartUp.oBrowse.ActiveRecord.Cells["sl_xuat"].Value.ToString(), out result3);
                Decimal result4;
                Decimal.TryParse(StartUp.oBrowse.ActiveRecord.Cells["ton_cuoi"].Value.ToString(), out result4);
                Decimal result5;
                Decimal.TryParse(StartUp.oBrowse.ActiveRecord.Cells["du_dau"].Value.ToString(), out result5);
                Decimal result6;
                Decimal.TryParse(StartUp.oBrowse.ActiveRecord.Cells["tien_nhap"].Value.ToString(), out result6);
                Decimal result7;
                Decimal.TryParse(StartUp.oBrowse.ActiveRecord.Cells["tien_xuat"].Value.ToString(), out result7);
                Decimal result8;
                Decimal.TryParse(StartUp.oBrowse.ActiveRecord.Cells["du_cuoi"].Value.ToString(), out result8);
                if (result1 != new Decimal(0))
                    formBrowse.AddValueSummary(new string[1] { "gia" }, new string[1]
                    {
            string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA").ToString() + "}", (object) SysFunc.Round(result5 / result1, StartUp.M_ROUND_GIA))
                    });
                else
                    formBrowse.AddValueSummary(new string[1] { "gia" }, new string[1]
                    {
            string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA").ToString() + "}", (object) 0)
                    });
                if (result2 != new Decimal(0))
                    formBrowse.AddValueSummary(new string[1] { "gia" }, new string[1]
                    {
            string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA").ToString() + "}", (object) SysFunc.Round(result6 / result2, StartUp.M_ROUND_GIA))
                    });
                else
                    formBrowse.AddValueSummary(new string[1] { "gia" }, new string[1]
                    {
            string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA").ToString() + "}", (object) 0)
                    });
                if (result3 != new Decimal(0))
                    formBrowse.AddValueSummary(new string[1] { "gia" }, new string[1]
                    {
            string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA").ToString() + "}", (object) SysFunc.Round(result7 / result3, StartUp.M_ROUND_GIA))
                    });
                else
                    formBrowse.AddValueSummary(new string[1] { "gia" }, new string[1]
                    {
            string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA").ToString() + "}", (object) 0)
                    });
                if (result4 != new Decimal(0))
                    formBrowse.AddValueSummary(new string[1] { "gia" }, new string[1]
                    {
            string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA").ToString() + "}", (object) SysFunc.Round(result8 / result4, StartUp.M_ROUND_GIA))
                    });
                else
                    formBrowse.AddValueSummary(new string[1] { "gia" }, new string[1]
                    {
            string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA").ToString() + "}", (object) 0)
                    });
                formBrowse.AddValueSummary(new string[1]
                {
          "sl_nhap"
                }, new string[1]
                {
          string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_SL").ToString() + "}", StartUp.oBrowse.ActiveRecord.Cells["ton_dau"].Value)
                });
                formBrowse.AddValueSummary(new string[1]
                {
          "tien_nhap"
                }, new string[1]
                {
          string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_TIEN").ToString() + "}", StartUp.oBrowse.ActiveRecord.Cells["du_dau"].Value)
                });
                formBrowse.AddValueSummary(new string[1]
                {
          "sl_xuat"
                }, new string[1] { "  " });
                formBrowse.AddValueSummary(new string[1]
                {
          "tien_xuat"
                }, new string[1] { "  " });
                formBrowse.AddValueSummary(new string[1]
                {
          "sl_nhap"
                }, new string[1]
                {
          string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_SL").ToString() + "}", StartUp.oBrowse.ActiveRecord.Cells["sl_nhap"].Value)
                });
                formBrowse.AddValueSummary(new string[1]
                {
          "tien_nhap"
                }, new string[1]
                {
          string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_TIEN").ToString() + "}", StartUp.oBrowse.ActiveRecord.Cells["tien_nhap"].Value)
                });
                formBrowse.AddValueSummary(new string[1]
                {
          "sl_xuat"
                }, new string[1] { "  " });
                formBrowse.AddValueSummary(new string[1]
                {
          "tien_xuat"
                }, new string[1] { "  " });
                formBrowse.AddValueSummary(new string[1]
                {
          "sl_nhap"
                }, new string[1] { "  " });
                formBrowse.AddValueSummary(new string[1]
                {
          "tien_nhap"
                }, new string[1] { "  " });
                formBrowse.AddValueSummary(new string[1]
                {
          "sl_xuat"
                }, new string[1]
                {
          string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_SL").ToString() + "}", StartUp.oBrowse.ActiveRecord.Cells["sl_xuat"].Value)
                });
                formBrowse.AddValueSummary(new string[1]
                {
          "tien_xuat"
                }, new string[1]
                {
          string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_TIEN").ToString() + "}", StartUp.oBrowse.ActiveRecord.Cells["tien_xuat"].Value)
                });
                formBrowse.AddValueSummary(new string[1]
                {
          "sl_nhap"
                }, new string[1]
                {
          string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_SL").ToString() + "}", StartUp.oBrowse.ActiveRecord.Cells["ton_cuoi"].Value)
                });
                formBrowse.AddValueSummary(new string[1]
                {
          "tien_nhap"
                }, new string[1]
                {
          string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_TIEN").ToString() + "}", StartUp.oBrowse.ActiveRecord.Cells["du_cuoi"].Value)
                });
                formBrowse.AddValueSummary(new string[1]
                {
          "sl_xuat"
                }, new string[1] { "  " });
                formBrowse.AddValueSummary(new string[1]
                {
          "tien_xuat"
                }, new string[1] { "  " });
                if (StartUp.oBrowse.DataGrid.DefaultFieldLayout.Fields.Where<Field>((Func<Field, bool>)(p => p.Name == "tien_nt_n")).Count<Field>() > 0)
                {
                    Decimal result9;
                    Decimal.TryParse(StartUp.oBrowse.ActiveRecord.Cells["du_dau_nt"].Value.ToString(), out result9);
                    Decimal result10;
                    Decimal.TryParse(StartUp.oBrowse.ActiveRecord.Cells["tien_nt_n"].Value.ToString(), out result10);
                    Decimal result11;
                    Decimal.TryParse(StartUp.oBrowse.ActiveRecord.Cells["tien_nt_x"].Value.ToString(), out result11);
                    Decimal result12;
                    Decimal.TryParse(StartUp.oBrowse.ActiveRecord.Cells["du_cuoi_nt"].Value.ToString(), out result12);
                    if (result1 != new Decimal(0))
                        formBrowse.AddValueSummary(new string[1]
                        {
              "gia_nt"
                        }, new string[1]
                        {
              string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA_NT").ToString() + "}", (object) SysFunc.Round(result9 / result1, StartUp.M_ROUND_GIA_NT))
                        });
                    else
                        formBrowse.AddValueSummary(new string[1]
                        {
              "gia_nt"
                        }, new string[1]
                        {
              string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA_NT").ToString() + "}", (object) 0)
                        });
                    if (result2 != new Decimal(0))
                        formBrowse.AddValueSummary(new string[1]
                        {
              "gia_nt"
                        }, new string[1]
                        {
              string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA_NT").ToString() + "}", (object) SysFunc.Round(result10 / result2, StartUp.M_ROUND_GIA_NT))
                        });
                    else
                        formBrowse.AddValueSummary(new string[1]
                        {
              "gia_nt"
                        }, new string[1]
                        {
              string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA_NT").ToString() + "}", (object) 0)
                        });
                    if (result3 != new Decimal(0))
                        formBrowse.AddValueSummary(new string[1]
                        {
              "gia_nt"
                        }, new string[1]
                        {
              string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA_NT").ToString() + "}", (object) SysFunc.Round(result11 / result3, StartUp.M_ROUND_GIA_NT))
                        });
                    else
                        formBrowse.AddValueSummary(new string[1]
                        {
              "gia_nt"
                        }, new string[1]
                        {
              string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA_NT").ToString() + "}", (object) 0)
                        });
                    if (result4 != new Decimal(0))
                        formBrowse.AddValueSummary(new string[1]
                        {
              "gia_nt"
                        }, new string[1]
                        {
              string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA_NT").ToString() + "}", (object) SysFunc.Round(result12 / result4, StartUp.M_ROUND_GIA_NT))
                        });
                    else
                        formBrowse.AddValueSummary(new string[1]
                        {
              "gia_nt"
                        }, new string[1]
                        {
              string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_GIA_NT").ToString() + "}", (object) 0)
                        });
                    formBrowse.AddValueSummary(new string[1]
                    {
            "tien_nt_n"
                    }, new string[1]
                    {
            string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString() + "}", StartUp.oBrowse.ActiveRecord.Cells["du_dau_nt"].Value)
                    });
                    formBrowse.AddValueSummary(new string[1]
                    {
            "tien_nt_x"
                    }, new string[1] { "  " });
                    formBrowse.AddValueSummary(new string[1]
                    {
            "tien_nt_n"
                    }, new string[1]
                    {
            string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString() + "}", StartUp.oBrowse.ActiveRecord.Cells["tien_nt_n"].Value)
                    });
                    formBrowse.AddValueSummary(new string[1]
                    {
            "tien_nt_x"
                    }, new string[1] { "  " });
                    formBrowse.AddValueSummary(new string[1]
                    {
            "tien_nt_x"
                    }, new string[1]
                    {
            string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString() + "}", StartUp.oBrowse.ActiveRecord.Cells["tien_nt_x"].Value)
                    });
                    formBrowse.AddValueSummary(new string[1]
                    {
            "tien_nt_n"
                    }, new string[1] { "  " });
                    formBrowse.AddValueSummary(new string[1]
                    {
            "tien_nt_n"
                    }, new string[1]
                    {
            string.Format("{0:" + StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString() + "}", StartUp.oBrowse.ActiveRecord.Cells["du_cuoi_nt"].Value)
                    });
                    formBrowse.AddValueSummary(new string[1]
                    {
            "tien_nt_x"
                    }, new string[1] { "  " });
                }
                formBrowse.frmBrw.LanguageID = "Incd1_2";
                formBrowse.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private static void _formDetail_Esc(object sender, EventArgs e)
        {
            (sender as SasFormBrowes.FormBrowse).frmBrw.Close();
        }

        private static void FormBrowse_Esc(object sender, EventArgs e)
        {
        }

        private static void oBrowse_F7(object sender, EventArgs e)
        {
            ReportManager reportManager = new ReportManager(StartupBase.SasObj, StartUp.CommandInfo["rep_file"].ToString(), StartUp.KindStyleReport);
            SysFunc.DSCopyWithFilter(StartUp.oBrowse.frmBrw.oBrowse, ref StartUp.DataSourceReport, "tbDetail");
            reportManager.Preview(StartUp.DataSourceReport);
        }

        private static void oBrowse_F11(object sender, EventArgs e)
        {
            Incd1F10 incd1F10 = new Incd1F10(StartUp.GroupSelected);
            incd1F10.DisplayLanguage = StartupBase.M_LAN;
            SysFunc.LoadIcon((Window)incd1F10);
            incd1F10.Title = SysFunc.Cat_Dau(incd1F10.Title);
            try
            {
                if (!incd1F10.ShowDialog())
                    return;
                StartUp.GroupSelected = incd1F10.DataOption.Copy();
                if (StartUp.DtGroupInfo == null)
                {
                    SqlCommand sqlcmd = new SqlCommand("Select loai_nh, ma_nh, ten_nh, ten_nh2 from dmnhvt");
                    StartUp.DtGroupInfo = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
                }
                string GroupFields = "";
                string SortFields = "";
                int result1 = 0;
                int result2;
                int.TryParse(incd1F10.DataOption.Rows[0]["group1"].ToString(), out result2);
                int result3;
                int.TryParse(incd1F10.DataOption.Rows[0]["group2"].ToString(), out result3);
                int result4;
                int.TryParse(incd1F10.DataOption.Rows[0]["group3"].ToString(), out result4);
                int.TryParse(incd1F10.DataOption.Rows[0]["sortby"].ToString(), out result1);
                if (result2 != 0)
                    GroupFields = GroupFields + (GroupFields == "" ? "" : ";") + incd1F10.DataOption.Rows[0]["group1"].ToString();
                if (result3 != 0)
                    GroupFields = GroupFields + (GroupFields == "" ? "" : ";") + incd1F10.DataOption.Rows[0]["group2"].ToString();
                if (result4 != 0)
                    GroupFields = GroupFields + (GroupFields == "" ? "" : ";") + incd1F10.DataOption.Rows[0]["group3"].ToString();
                switch (result1)
                {
                    case 0:
                        SortFields = StartupBase.M_LAN == "V" ? "ten_vt" : "ten_vt2";
                        break;
                    case 1:
                        SortFields = "ma_vt";
                        break;
                }
                SysFunc.InsertListGroup(StartUp.tbDetail, StartUp.DtGroupInfo, "nh_vt", GroupFields, "ten_vt;ten_vt2", SortFields, StartUp.SumFields);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public static string GetTableShow(int KindReport, int KindPrint)
        {
            string empty = string.Empty;
            switch (StartupBase.M_LAN.ToUpper())
            {
                case "V":
                    switch (KindReport)
                    {
                        case 1:
                            string[] strArray1 = StartUp.CommandInfo["Vbrowse1"].ToString().Split('|');
                            switch (KindPrint)
                            {
                                case 1:
                                    empty = strArray1[0];
                                    break;
                                case 2:
                                    empty = strArray1[1];
                                    break;
                                case 3:
                                    empty = strArray1[2];
                                    break;
                            }
                            break;
                        case 2:
                            string[] strArray2 = StartUp.CommandInfo["Vbrowse2"].ToString().Split('|');
                            switch (KindPrint)
                            {
                                case 1:
                                    empty = strArray2[0];
                                    break;
                                case 2:
                                    empty = strArray2[1];
                                    break;
                                case 3:
                                    empty = strArray2[2];
                                    break;
                            }
                            break;
                    }
                    break;
                default:
                    switch (KindReport)
                    {
                        case 1:
                            string[] strArray3 = StartUp.CommandInfo["Ebrowse1"].ToString().Split('|');
                            switch (KindPrint)
                            {
                                case 1:
                                    empty = strArray3[0];
                                    break;
                                case 2:
                                    empty = strArray3[1];
                                    break;
                                case 3:
                                    empty = strArray3[2];
                                    break;
                            }
                            break;
                        case 2:
                            string[] strArray4 = StartUp.CommandInfo["Ebrowse2"].ToString().Split('|');
                            switch (KindPrint)
                            {
                                case 1:
                                    empty = strArray4[0];
                                    break;
                                case 2:
                                    empty = strArray4[1];
                                    break;
                                case 3:
                                    empty = strArray4[2];
                                    break;
                            }
                            break;
                    }
                    break;
            }
            return empty;
        }

        public static string GetTable4DetailShow(int KindReport, int KindPrint)
        {
            string empty = string.Empty;
            switch (StartupBase.M_LAN.ToUpper())
            {
                case "V":
                    switch (KindReport)
                    {
                        case 1:
                            string[] strArray1 = StartUp.CommandInfo["Vbrowse1"].ToString().Split('|');
                            if (KindPrint == 1)
                            {
                                empty = strArray1[1];
                                break;
                            }
                            break;
                        case 2:
                            string[] strArray2 = StartUp.CommandInfo["Vbrowse2"].ToString().Split('|');
                            if (KindPrint == 1)
                            {
                                empty = strArray2[1];
                                break;
                            }
                            break;
                    }
                    break;
                default:
                    switch (KindReport)
                    {
                        case 1:
                            string[] strArray3 = StartUp.CommandInfo["Ebrowse1"].ToString().Split('|');
                            if (KindPrint == 1)
                            {
                                empty = strArray3[1];
                                break;
                            }
                            break;
                        case 2:
                            string[] strArray4 = StartUp.CommandInfo["Ebrowse2"].ToString().Split('|');
                            if (KindPrint == 1)
                            {
                                empty = strArray4[1];
                                break;
                            }
                            break;
                    }
                    break;
            }
            return empty;
        }
    }
}
