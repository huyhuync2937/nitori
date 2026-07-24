using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Windows;

namespace SasPermission
{
    public class StartUp : StartupBase
    {
        public static string sqlTableName = "userinfo";
        public static string SqlTableKey = nameof(user_id);
        public static string SqlTableObjectName = "user_name";
        public static string TableName = "userinfo";
        public static ActionTask currActionTask = ActionTask.None;
        public static string currSqlTableKey = string.Empty;
        public static string titleWindow = string.Empty;
        public static string strCmdUserInfo = "";
        public static DataRow CommandInfo;
        public static DataTable dtRegInfo;
        public static DataTable dt;
        private Permission _Form;
        private FrmLogin _frmIn;
        public static int user_id;
        public static int is_admin;
        public static int admin_root;

        public override void Run()
        {
            StartupBase.Namespace = "SasPermission";
            this._Form = new Permission();
            this._frmIn = new FrmLogin();
            StartUp.user_id = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["user_id"]);
            StartUp.admin_root = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["admin_root"]);
            StartUp.is_admin = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["is_admin"]);
            StartUp.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
            if (StartUp.CommandInfo == null)
            {
                if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    return;
                Application.Current.Shutdown();
            }
            else
            {
                StartUp.dtRegInfo = StartupBase.SasObj.GetRegInfo();
                this._frmIn.ShowDialog();
                if (!this._frmIn.IsLogined)
                {
                    this._frmIn.Close();
                    this._Form.Close();
                }
                else
                {
                    StartUp.strCmdUserInfo = "select * from " + StartUp.sqlTableName;
                    if (StartUp.is_admin != 1 || StartUp.admin_root != 1)
                    {
                        StartUp.strCmdUserInfo += " where (1 = 1 OR admin_root != 1)";
                        if (StartUp.is_admin == 0)
                            StartUp.strCmdUserInfo = StartUp.strCmdUserInfo + " and is_admin = 0 and user_id = " + StartUp.user_id.ToString();
                    }
                    StartUp.strCmdUserInfo = StartUp.strCmdUserInfo + " order by " + StartUp.SqlTableObjectName + " asc";
                    SqlCommand sqlcmd = new SqlCommand(StartUp.strCmdUserInfo);
                    StartUp.dt = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
                    this._Form.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.CommandInfo["title"].ToString() : StartUp.CommandInfo["title2"].ToString());
                    this._Form.ShowDialog();
                }
            }
        }

        public static DataTable GetRow(string sqlTableName)
        {
            DataTable dataTable = (DataTable)null;
            SqlCommand sqlcmd = new SqlCommand();
            try
            {
                string str = "select * from " + sqlTableName + " where ";
                if (StartUp.currActionTask == ActionTask.Add)
                {
                    str += "1=0";
                }
                else
                {
                    DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(StartupBase.SasObj, sqlTableName);
                    string[] strArray1 = StartUp.currSqlTableKey.Split(';');
                    string[] strArray2 = StartUp.SqlTableKey.Split(';');
                    for (int index = 0; index < ((IEnumerable<string>)strArray2).Count<string>(); ++index)
                    {
                        if (index > 0)
                            str += " and";
                        DataRow[] dataRowArray = sqlTableFieldList.Select("name='" + strArray2[index] + "'");
                        str = str + " " + strArray2[index] + "=@" + strArray2[index];
                        sqlcmd.Parameters.Add("@" + strArray2[index], ListFunc.GetSqlDBType(dataRowArray[0]["datatype"].ToString())).Value = (object)strArray1[index];
                    }
                }
                sqlcmd.CommandText = str;
                dataTable = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
            }
            catch (SqlException ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return dataTable;
        }
    }
}
