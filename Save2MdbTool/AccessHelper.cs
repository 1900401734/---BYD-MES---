using System;
using System.Data;
using System.Data.OleDb;
using System.Collections;
/// <summary>
/// AcceHelper数据库操作类
/// </summary>
public static class AccessHelper
{
    
    /// <summary>
    /// 给定连接的数据库用假设参数执行一个sql命令（不返回数据集）
    /// </summary>
    /// <param name="connectionString">一个有效的连接字符串</param>
    /// <param name="commandText">存储过程名称或者sql命令语句</param>
    /// <param name="commandParameters">执行命令所用参数的集合</param>
    /// <returns>执行命令所影响的行数</returns>
    public static int ExecuteNonQuery(string connectionString, string cmdText, params OleDbParameter[] commandParameters)
    {
        OleDbCommand cmd = new OleDbCommand();
        using (OleDbConnection conn = new OleDbConnection(connectionString))
        {
            PrepareCommand(cmd, conn, null, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }
    }

    /// <summary>
    /// 用执行的数据库连接执行一个返回数据集的sql命令
    /// </summary>
    /// <remarks>
    /// 举例: 
    /// OleDbDataReader r = ExecuteReader(connString, "PublishOrders", new OleDbParameter("@prodid", 24));
    /// </remarks>
    /// <param name="connectionString">一个有效的连接字符串</param>
    /// <param name="commandText">存储过程名称或者sql命令语句</param>
    /// <param name="commandParameters">执行命令所用参数的集合</param>
    /// <returns>包含结果的读取器</returns>
    public static OleDbDataReader ExecuteReader(string connectionString, string cmdText, params OleDbParameter[] commandParameters)
    {
        //创建一个SqlCommand对象
        OleDbCommand cmd = new OleDbCommand();
        //创建一个SqlConnection对象
        OleDbConnection conn = new OleDbConnection(connectionString);
        //在这里我们用一个try/catch结构执行sql文本命令/存储过程，因为如果这个方法产生一个异常我们要关闭连接，因为没有读取器存在，
        //因此commandBehaviour.CloseConnection 就不会执行
        try
        {
            //调用 PrepareCommand 方法，对 SqlCommand 对象设置参数
            PrepareCommand(cmd, conn, null, cmdText, commandParameters);
            //调用 SqlCommand 的 ExecuteReader 方法
            OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            //清除参数
            cmd.Parameters.Clear();
            return reader;
        }
        catch
        {
            //关闭连接，抛出异常
            conn.Close();
            throw;
        }
    }
    /// <summary>
    /// 返回一个DataSet数据集
    /// </summary>
    /// <param name="connectionString">一个有效的连接字符串</param>
    /// <param name="cmdText">存储过程名称或者sql命令语句</param>
    /// <param name="commandParameters">执行命令所用参数的集合</param>
    /// <returns>包含结果的数据集</returns>
    public static DataSet ExecuteDataSet(string connectionString, string cmdText, params OleDbParameter[] commandParameters)
    {
        //创建一个SqlCommand对象，并对其进行初始化
        OleDbCommand cmd = new OleDbCommand();
        using (OleDbConnection conn = new OleDbConnection(connectionString))
        {
            PrepareCommand(cmd, conn, null, cmdText, commandParameters);
            //创建SqlDataAdapter对象以及DataSet
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            try
            {
                //填充ds
                da.Fill(ds);
                // 清除cmd的参数集合 
                cmd.Parameters.Clear();
                //返回ds
                return ds;
            }
            catch
            {
                //关闭连接，抛出异常
                conn.Close();
                throw;
            }
        }
    }
    /// <summary>
    /// 返回一个DataTable
    /// </summary>
    /// 列如：
    /// Object obj = ExecuteScalar(connString, "PublishOrders", new OleDbParameter("@prodid", 24));
    ///<param name="connectionString">一个有效的连接字符串</param>
    /// <param name="commandText">存储过程名称或者sql命令语句</param>
    /// <param name="commandParameters">执行命令所用参数的集合</param>
    /// <returns></returns>
    public static DataTable ExecuteDataTable(string connectionString,string cmdText, params OleDbParameter[] commandParameters)
    {
        using (OleDbConnection conn = new OleDbConnection(connectionString))
        {
            using (OleDbCommand cmd = new OleDbCommand(cmdText, conn))
            {
                if (commandParameters != null)
                {
                    cmd.Parameters.AddRange(commandParameters);
                }
                DataTable dt = new DataTable();
                OleDbDataAdapter sda = new OleDbDataAdapter(cmd);
                sda.Fill(dt);
                return (dt);
            }
        }
    }

    /// <summary>
    /// 用指定的数据库连接字符串执行一个命令并返回一个数据集的第一列
    /// </summary>
    /// <remarks>
    ///例如: 
    /// Object obj = ExecuteScalar(connString, "PublishOrders", new OleDbParameter("@prodid", 24));
    /// </remarks>
    ///<param name="connectionString">一个有效的连接字符串</param>
    /// <param name="commandText">存储过程名称或者sql命令语句</param>
    /// <param name="commandParameters">执行命令所用参数的集合</param>
    /// <returns>用 Convert.To{Type}把类型转换为想要的 </returns>
    public static object ExecuteScalar(string connectionString, string cmdText, params OleDbParameter[] commandParameters)
    {
        OleDbCommand cmd = new OleDbCommand();
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            PrepareCommand(cmd, connection, null, cmdText, commandParameters);
            object val = cmd.ExecuteScalar();
            cmd.Parameters.Clear();
            return val;
        }
    }


    /// <summary>
    /// 准备执行一个命令
    /// </summary>
    /// <param name="cmd">sql命令</param>
    /// <param name="conn">Sql连接</param>
    /// <param name="trans">Sql事务</param>
    /// <param name="cmdText">命令文本,例如：Select * from Products</param>
    /// <param name="cmdParms">执行命令的参数</param>
    private static void PrepareCommand(OleDbCommand cmd, OleDbConnection conn, OleDbTransaction trans, string cmdText, OleDbParameter[] cmdParms)
    {
        //判断连接的状态。如果是关闭状态，则打开
        if (conn.State != ConnectionState.Open)
            conn.Open();
        //cmd属性赋值
        cmd.Connection = conn;
        cmd.CommandText = cmdText;
        //是否需要用到事务处理
        if (trans != null)
            cmd.Transaction = trans;
        cmd.CommandType = CommandType.Text;
        //添加cmd需要的存储过程参数
        if (cmdParms != null)
        {
            foreach (OleDbParameter parm in cmdParms)
                cmd.Parameters.Add(parm);
        }
    }
}

