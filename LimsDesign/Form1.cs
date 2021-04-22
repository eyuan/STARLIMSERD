using LimsDesign.Utils;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Configuration;

namespace LimsDesign
{
    public partial class Form1 : Form
    {
        private DBType dbType = DBType.SqlServer;
        private string[] tableFilterList = new string[] { };
        public Form1()
        {
            InitializeComponent();
            string dbTypeConfig= ConfigurationManager.AppSettings["DBType"];
            if (dbTypeConfig == "oracle")
            {
                dbType = DBType.Oracle;
            }
            else
            {
                dbType = DBType.SqlServer;
            }
            string TableFilters = ConfigurationManager.AppSettings["TableFilters"];
            tableFilterList = TableFilters.Split(',');
        }

        private void GenerateDesignDoc()
        {
            #region 创建word文档
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application(); ;
            //Microsoft.Office.Interop.Word._Application app = new Microsoft.Office.Interop.Word.ApplicationClass();
            //app.Visible = true;
            app.Visible = false;
            object nullobj = Type.Missing;
            object Nothing = Type.Missing;
            object obpic = "pic";
            object obtable = "obtable";
            object file = Environment.CurrentDirectory + @"\" + Guid.NewGuid().ToString() + ".docx";
            FileInfo fileInfo = new FileInfo("template.docx");
            fileInfo.CopyTo((string)file);
            Microsoft.Office.Interop.Word._Document doc = app.Documents.Open(
            ref file, ref nullobj, ref nullobj,
            ref nullobj, ref nullobj, ref nullobj,
            ref nullobj, ref nullobj, ref nullobj,
            ref nullobj, ref nullobj, ref nullobj,
            ref nullobj, ref nullobj, ref nullobj, ref nullobj) as Microsoft.Office.Interop.Word._Document;
            int tableRow = 0;
            int tableColumn = 0;
            #endregion

            //and TABLENAME in ('RESULTS')
            string sql = "select tablename, tableid, description from limstables where issystem = 'N' order by  tablename";
            string sqlMSSql = sql;
            string sqlOracle = sql;
            DataTable dataTable = null;
            if (dbType == DBType.SqlServer)
            {
                 dataTable = SQLHelper.GetDataTable(sqlMSSql);
            }
            else { 
                 dataTable = OracleHelper.ExecToSqlGetTable(sqlOracle);
            }
            //循环各个表
            int iTable = 0;
            foreach (DataRow dataRow in dataTable.Rows)
            { 
                #region 表概述
                string tableid = dataRow["tableid"].ToString();
                string tablename = dataRow["tablename"].ToString();
                string description = dataRow["description"].ToString();
                //表概述
                //添加标题段落 
                doc.Paragraphs.Last.Range.set_Style("1");
                doc.Paragraphs.Last.Range.InsertAfter("表 " + tablename + "\n");
                doc.Paragraphs.Last.Range.set_Style("正文");
                //
                //app.Selection.Paragraphs.Last.Range.Text = "表名称：" + tablename+ "\n";
                //app.Selection.Paragraphs.Last.Range.Text = "表描述：" + description + "\n";
                doc.Paragraphs.Last.Range.InsertAfter("表名称：" + tablename + "\n");
                //doc.ActiveWindow.Selection.TypeText("表名称：" + tablename + "\n");
                doc.Paragraphs.Last.Range.InsertAfter("表描述：" + description + "\n");
                #endregion
                
                #region 表字段定义
                doc.Paragraphs.Last.Range.set_Style("2");
                doc.Paragraphs.Last.Range.InsertAfter("字段定义\n");
                doc.Paragraphs.Last.Range.set_Style("表标题");
                //doc.Paragraphs.Last.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                doc.Paragraphs.Last.Range.InsertAfter("表"+tablename+"字段定义\n");
                doc.Paragraphs.Last.Range.set_Style("正文");
                //doc.Paragraphs.Last.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                sql = @"select f.name,f.description,f.type,f.defaultvalue,f.isnullable, r.caption 
                    from LIMSTABLEFIELDS f
                    left join LIMSTABLERESOURCES r 
                        on f.TABLEID = r.TABLEID 
                        and f.FIELDID = r.FIELDID 
                        and r.LANGID = 'CHS'
                    ";
                sqlMSSql = sql + @" where f.TABLEID = @TABLEID order by f.fieldid";
                sqlOracle = sql + @" where f.TABLEID = :TABLEID order by f.fieldid";
                DataTable dtField = null;
                if (dbType == DBType.SqlServer)
                {
                    dtField = SQLHelper.GetDataTable(sqlMSSql, new SqlParameter[] { new SqlParameter("@TABLEID", tableid) });
                }
                else
                {
                    dtField = OracleHelper.ExecToSqlGetTable(sqlOracle, new OracleParameter[] { new OracleParameter(":TABLEID", tableid) });
                }
                if (dtField.Rows.Count == 0)
                {
                    bgw.ReportProgress(++iTable);
                    continue;
                }
                //Word.Table newTable = doc.Tables.Add(doc.Bookmarks.get_Item(ref obtable).Range, dtField.Rows.Count, 6, ref nullobj, ref nullobj);
                tableRow = dtField.Rows.Count+1;
                tableColumn = 6;
                //Word.Table table = doc.Tables.Add(app.Selection.Range,  tableRow, tableColumn, ref Nothing, ref Nothing);
                Word.Table table = doc.Tables.Add(doc.Paragraphs.Last.Range, tableRow, tableColumn, ref Nothing, ref Nothing);
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                //table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "字段名";
                table.Cell(1, 2).Range.Text = "字段描述";
                table.Cell(1, 3).Range.Text = "数据类型";
                table.Cell(1, 4).Range.Text = "默认值";
                table.Cell(1, 5).Range.Text = "允许空?";
                table.Cell(1, 6).Range.Text = "标题";
                int rowIndex = 1;
                //循环字段
                foreach (DataRow drField in dtField.Rows)
                {
                    rowIndex++;
                    string name = drField["name"].ToString();
                    string fdescription = drField["description"].ToString();
                    string type = drField["type"].ToString();
                    string defaultvalue = drField["defaultvalue"].ToString();
                    string isnullable = drField["isnullable"].ToString();
                    string caption = drField["caption"].ToString();
                    table.Cell(rowIndex, 1).Range.Text = name;
                    table.Cell(rowIndex, 2).Range.Text = fdescription;
                    table.Cell(rowIndex, 3).Range.Text = type;
                    table.Cell(rowIndex, 4).Range.Text = defaultvalue;
                    table.Cell(rowIndex, 5).Range.Text = isnullable;
                    table.Cell(rowIndex, 6).Range.Text = caption;
                }
                doc.Paragraphs.Last.Range.InsertAfter("\n");
                #endregion

                #region 索引定义
                sql = "select * from LIMSTABLEINDEXES i where i.tableid = {0}";
                DataTable dtIndex = null;
                if (dbType == DBType.SqlServer)
                {
                    dtIndex = SQLHelper.GetDataTable(string.Format(sql, "@TABLEID"), new SqlParameter[] { new SqlParameter("@TABLEID", tableid) });
                }
                else
                {
                    dtIndex = OracleHelper.ExecToSqlGetTable(string.Format(sql, ":TABLEID"), new OracleParameter[] { new OracleParameter(":TABLEID", tableid) });
                } 
                if (dtIndex.Rows.Count > 0)
                { 
                    doc.Paragraphs.Last.Range.set_Style("2");
                    doc.Paragraphs.Last.Range.InsertAfter("索引定义\n");
                    doc.Paragraphs.Last.Range.set_Style("正文");
                }
                //循环各个索引
                foreach (DataRow drIndex in dtIndex.Rows)
                {
                    string indexName = drIndex["INDEXNAME"].ToString();
                    string indexType = drIndex["INDEXTYPE"].ToString();
                    doc.Paragraphs.Last.Range.set_Style("3");
                    doc.Paragraphs.Last.Range.InsertAfter("索引 " + indexName + "\n");
                    doc.Paragraphs.Last.Range.set_Style("正文");
                    doc.Paragraphs.Last.Range.InsertAfter("索引名称：" + indexName + "\n");
                    doc.Paragraphs.Last.Range.InsertAfter("索引类型：" + indexType + "\n");
                    doc.Paragraphs.Last.Range.set_Style("表标题");
                    //doc.Paragraphs.Last.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    doc.Paragraphs.Last.Range.InsertAfter("索引"+indexName+"字段定义\n");
                    doc.Paragraphs.Last.Range.set_Style("正文");
                    //doc.Paragraphs.Last.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                    #region 索引字段定义
                    sql = @"SELECT F.NAME, IDF.SORT, IDF.INCLUDED from LIMSTABLEINDEXES i
                        LEFT JOIN LIMSTABLEINDEXFIELDS IDF ON I.TABLEID = IDF.TABLEID AND I.INDEXID = IDF.INDEXID
                        LEFT JOIN LIMSTABLEFIELDS F ON IDF.TABLEID = F.TABLEID AND IDF.FIELDID = F.FIELDID
                        where i.TABLEID = {0}     
                              and idf.indexid = {1}  ";
                    DataTable dtIndexField = null;
                    if (dbType == DBType.SqlServer)
                    {
                        dtIndexField = SQLHelper.GetDataTable(string.Format(sql,"@TABLEID", "@INDEXID")
                            , new SqlParameter[] { new SqlParameter("@TABLEID", tableid), new SqlParameter("@INDEXID", drIndex["INDEXID"].ToString()) });
                    }
                    else
                    {
                        dtIndexField = OracleHelper.ExecToSqlGetTable(string.Format(sql, new string[] { ":TABLEID", ":INDEXID" })
                       , new OracleParameter[] { new OracleParameter(":TABLEID", tableid), new OracleParameter(":INDEXID", drIndex["INDEXID"].ToString()) });
                    } 
                    tableRow = dtIndexField.Rows.Count + 1;
                    tableColumn = 3;
                    Word.Table tableIndexField = doc.Tables.Add(doc.Paragraphs.Last.Range, tableRow, tableColumn, ref Nothing, ref Nothing);
                    tableIndexField.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tableIndexField.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tableIndexField.Cell(1, 1).Range.Text = "字段名";
                    tableIndexField.Cell(1, 2).Range.Text = "排序";
                    tableIndexField.Cell(1, 3).Range.Text = "包含";
                    rowIndex = 1;
                    //循环索引字段 
                    foreach (DataRow drIndexField in dtIndexField.Rows)
                    {
                        rowIndex++;
                        tableIndexField.Cell(rowIndex, 1).Range.Text = drIndexField["NAME"].ToString();
                        tableIndexField.Cell(rowIndex, 2).Range.Text = drIndexField["SORT"].ToString();
                        tableIndexField.Cell(rowIndex, 3).Range.Text = drIndexField["INCLUDED"].ToString();
                    }
                    doc.Paragraphs.Last.Range.InsertAfter("\n");
                    #endregion
                }
                #endregion

                #region 外键约束定义
                sql = "select TR.*, T.TABLENAME" +
                    " from LIMSTABLERELATIONS TR " +
                    " LEFT JOIN LIMSTABLES T ON TR.REFTABLEID = T.TABLEID" +
                    " WHERE TR.TABLEID = {0} order by TR.FKID";
                DataTable dtRel = null;
                if (dbType == DBType.SqlServer)
                {
                    dtRel = SQLHelper.GetDataTable(string.Format(sql, "@TABLEID")
                    , new SqlParameter[] { new SqlParameter("@TABLEID", tableid) });
                }
                else
                {
                    dtRel = OracleHelper.ExecToSqlGetTable(string.Format(sql, ":TABLEID")
                    , new OracleParameter[] { new OracleParameter(":TABLEID", tableid) });
                }
                
                if (dtRel.Rows.Count > 0)
                { 
                    doc.Paragraphs.Last.Range.set_Style("2");
                    doc.Paragraphs.Last.Range.InsertAfter("外键约束定义\n");
                    doc.Paragraphs.Last.Range.set_Style("正文");
                }
                foreach(DataRow drRel in dtRel.Rows)
                {
                    string fkName = drRel["FKNAME"].ToString();
                    doc.Paragraphs.Last.Range.set_Style("3"); 
                    doc.Paragraphs.Last.Range.InsertAfter("约束 "+ fkName + "\n");
                    doc.Paragraphs.Last.Range.set_Style("正文");
                    doc.Paragraphs.Last.Range.InsertAfter("关联表：" + drRel["TABLENAME"].ToString() + "\n");
                    doc.Paragraphs.Last.Range.InsertAfter("级联更新：" + drRel["CASCADEUPDATE"].ToString() + "\n");
                    doc.Paragraphs.Last.Range.InsertAfter("级联删除：" + drRel["CASCADEDELETE"].ToString() + "\n");

                    #region 关联字段
                    doc.Paragraphs.Last.Range.set_Style("表标题");
                    //doc.Paragraphs.Last.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    doc.Paragraphs.Last.Range.InsertAfter("外键"+ fkName + "字段定义\n");
                    doc.Paragraphs.Last.Range.set_Style("正文");
                    //doc.Paragraphs.Last.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    sql = @"SELECT TF.NAME FieldName, TFR.NAME RelFieldName 
                            FROM LIMSTABLERELATIONFIELDS TRF
                            --LEFT JOIN LIMSTABLES T ON TRF.TABLEID = T.TABLEID
                            --LEFT JOIN LIMSTABLES TR ON TRF.REFTABLEID = TR.TABLEID
                            LEFT JOIN LIMSTABLEFIELDS TF ON TRF.TABLEID = TF.TABLEID AND TRF.FIELDID = TF.FIELDID
                            LEFT JOIN LIMSTABLEFIELDS TFR ON TRF.REFTABLEID = TFR.TABLEID AND TRF.REFFIELDID = TFR.FIELDID
                            WHERE TRF.TABLEID = {0}
                                  AND TRF.FKID = {1}";
                    DataTable dtRelField = null;
                    if (dbType == DBType.SqlServer)
                    {
                        dtRelField = SQLHelper.GetDataTable(string.Format(sql, new string[] { "@TABLEID", "@FKID" })
                   , new SqlParameter[] { new SqlParameter("@TABLEID", tableid)
                                , new SqlParameter("@FKID", drRel["FKID"].ToString()) });
                    }
                    else
                    {
                        dtRelField = OracleHelper.ExecToSqlGetTable(string.Format(sql, new string[] { ":TABLEID", ":FKID" })
                   , new OracleParameter[] { new OracleParameter(":TABLEID", tableid)
                                , new OracleParameter(":FKID", drRel["FKID"].ToString()) });
                    }
                   
                    //
                    tableRow = dtRelField.Rows.Count + 1;
                    tableColumn = 3;
                    Word.Table tableRefField = doc.Tables.Add(doc.Paragraphs.Last.Range, tableRow, tableColumn, ref Nothing, ref Nothing);
                    tableRefField.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tableRefField.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    //tableRefField.Borders.Enable = 1;
                    tableRefField.Cell(1, 1).Range.Text = "字段名";
                    tableRefField.Cell(1, 2).Range.Text = "关联表名称"; 
                    tableRefField.Cell(1, 3).Range.Text = "关联表字段名";
                    rowIndex = 1;
                    //循环字段
                    foreach (DataRow drRelField in dtRelField.Rows)
                    {
                        rowIndex++;  
                        tableRefField.Cell(rowIndex, 1).Range.Text = drRelField["FieldName"].ToString();
                        tableRefField.Cell(rowIndex, 2).Range.Text = drRel["TABLENAME"].ToString();
                        tableRefField.Cell(rowIndex, 3).Range.Text = drRelField["RelFieldName"].ToString();
                    }
                    doc.Paragraphs.Last.Range.InsertAfter("\n");
                    #endregion
                }
                doc.Paragraphs.Last.Range.InsertAfter("\n");
                #endregion

                #region 引用该表的子表定义
                sql = "select TR.*, T.TABLENAME" +
                    " from LIMSTABLERELATIONS TR " +
                    " LEFT JOIN LIMSTABLES T ON TR.TABLEID = T.TABLEID" +
                    " WHERE TR.REFTABLEID = {0} " +
                    "order by T.TABLENAME";
                DataTable dtRel2 = null;
                if (dbType == DBType.SqlServer)
                {
                    dtRel2 = SQLHelper.GetDataTable(string.Format(sql, "@TABLEID")
                   , new SqlParameter[] { new SqlParameter("@TABLEID", tableid) });
                }
                else
                {
                    dtRel2 = OracleHelper.ExecToSqlGetTable(string.Format(sql, ":TABLEID")
                   , new OracleParameter[] { new OracleParameter(":TABLEID", tableid) });
                } 
                if (dtRel2.Rows.Count > 0)
                { 
                    doc.Paragraphs.Last.Range.set_Style("2");
                    doc.Paragraphs.Last.Range.InsertAfter("子表定义\n");
                    doc.Paragraphs.Last.Range.set_Style("正文");
                }
                foreach (DataRow drRel2 in dtRel2.Rows)
                {
                    string subTableName = drRel2["TABLENAME"].ToString();
                    doc.Paragraphs.Last.Range.set_Style("3");
                    doc.Paragraphs.Last.Range.InsertAfter("子表 " + subTableName + "\n");
                    doc.Paragraphs.Last.Range.set_Style("正文");
                    doc.Paragraphs.Last.Range.InsertAfter("子表约束：" + drRel2["FKNAME"].ToString() + "\n");
                    doc.Paragraphs.Last.Range.InsertAfter("级联更新：" + drRel2["CASCADEUPDATE"].ToString() + "\n");
                    doc.Paragraphs.Last.Range.InsertAfter("级联删除：" + drRel2["CASCADEDELETE"].ToString() + "\n");
                  
                    #region 关联字段
                    doc.Paragraphs.Last.Range.set_Style("表标题");
                    //doc.Paragraphs.Last.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    doc.Paragraphs.Last.Range.InsertAfter("子表"+ subTableName + "约束字段定义\n"); 
                    doc.Paragraphs.Last.Range.set_Style("正文");
                    //doc.Paragraphs.Last.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    sql = @"SELECT TF.NAME FieldName --子表字段
                               , TFR.NAME RelFieldName --父表字段 
                               --, TRF.* 
                        FROM LIMSTABLERELATIONFIELDS TRF --关联字段
                        --LEFT JOIN LIMSTABLES T ON TRF.TABLEID = T.TABLEID
                        --LEFT JOIN LIMSTABLES TR ON TRF.REFTABLEID = TR.TABLEID
                        LEFT JOIN LIMSTABLEFIELDS TF ON TRF.TABLEID = TF.TABLEID AND TRF.FIELDID = TF.FIELDID --子表字段
                        LEFT JOIN LIMSTABLEFIELDS TFR ON TRF.REFTABLEID = TFR.TABLEID AND TRF.REFFIELDID = TFR.FIELDID --父表字段
                        WHERE TRF.REFTABLEID = {0} --父表ID
                              AND TRF.TABLEID = {1} --子表ID";
                    DataTable dtRelField2 = null;
                    if (dbType == DBType.SqlServer)
                    { 
                        dtRelField2 = SQLHelper.GetDataTable(string.Format(sql, new string[] { "@REFTABLEID", "@TABLEID" })
                   , new SqlParameter[] { new SqlParameter("@REFTABLEID", tableid)
                                , new SqlParameter("@TABLEID", drRel2["TABLEID"].ToString()) });
                    }
                    else
                    {
                        dtRelField2 = OracleHelper.ExecToSqlGetTable(string.Format(sql, new string[] { ":REFTABLEID", ":TABLEID" })
                   , new OracleParameter[] { new OracleParameter(":REFTABLEID", tableid)
                                , new OracleParameter(":TABLEID", drRel2["TABLEID"].ToString()) });
                    }
                    //
                    tableRow = dtRelField2.Rows.Count + 1;
                    tableColumn = 2;
                    Word.Table tableRefField = doc.Tables.Add(doc.Paragraphs.Last.Range, tableRow, tableColumn, ref Nothing, ref Nothing);
                    tableRefField.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tableRefField.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    //tableRefField.Borders.Enable = 1;
                    tableRefField.Cell(1, 1).Range.Text = "字段名"; 
                    tableRefField.Cell(1, 2).Range.Text = "子表字段名";
                    rowIndex = 1;
                    //循环字段
                    foreach (DataRow drRelField2 in dtRelField2.Rows)
                    {
                        rowIndex++;
                        tableRefField.Cell(rowIndex, 1).Range.Text = drRelField2["RelFieldName"].ToString(); 
                        tableRefField.Cell(rowIndex, 2).Range.Text = drRelField2["FieldName"].ToString();
                    }
                    doc.Paragraphs.Last.Range.InsertAfter("\n");
                    #endregion
                }
                doc.Paragraphs.Last.Range.InsertAfter("\n");
                #endregion

                #region 报告进度
                bgw.ReportProgress(++iTable);
                #endregion
            }
            //
            //保存
            doc.Save();
            //
            doc.Close(ref nullobj, ref nullobj, ref nullobj);
            doc = null;
            //
            app.Quit(ref nullobj, ref nullobj, ref nullobj);
            app = null;
        }

        private void GenerateDesignMarkdown()
        {  
            string sql = "select tablename, tableid, description from limstables where issystem = 'N' ";
            foreach (string tableFilter in tableFilterList)
            {
                sql += " and TABLENAME " + tableFilter;
            }
            sql += " order by  tablename ";
            string sqlMSSql = sql;
            string sqlOracle = sql;
            DataTable dataTable = null;
            if (dbType == DBType.SqlServer)
            {
                dataTable = SQLHelper.GetDataTable(sqlMSSql);
            }
            else
            {
                dataTable = OracleHelper.ExecToSqlGetTable(sqlOracle);
            }
            //循环各个表
            int iTable = 0;
            foreach (DataRow dataRow in dataTable.Rows)
            {
                int head1Index = 0;
                int head2Index = 0;

                #region 表概述
                string tableid = dataRow["tableid"].ToString();
                string tablename = dataRow["tablename"].ToString();
                string description = dataRow["description"].ToString();

                //生成Markdown文件
                //File.Create(tablename + "-" + description.Replace("\r","").Replace("\n","").Replace("\b", "").Replace(":","").Replace("*","").Replace("?","").Replace("<","").Replace(">","").Replace("|","").Replace("\"","").Replace("/","").Replace("\\", "") + ".md");
                StringBuilder sb = new StringBuilder();
                //表概述 
                sb.AppendLine();
                head1Index++;
                sb.AppendLine("# "+ head1Index + " 表定义");
                sb.AppendLine();
                sb.AppendLine("* 表名称："+tablename);
                sb.AppendLine("* 表描述：" + description.Replace("*","").Replace("#","") ); 
                sb.AppendLine();
                #endregion

                #region 表字段定义  
                head1Index++;
                sb.AppendLine("# "+ head1Index + " 字段定义");
                sb.AppendLine();
                string tableFields = "";
                string tableFieldsUpdateSql = "";
                string tableFieldsDefineMySql = "";
                string tableFieldsDefineSqlServer = "";
                string tableFieldsDefineOracle = "";
                string tableFieldsCommentOracle = "";
                sql = @"select f.name,f.description,f.type,f.defaultvalue,f.isnullable, r.caption ,f.length,f.scale
                    from LIMSTABLEFIELDS f
                    left join LIMSTABLERESOURCES r 
                        on f.TABLEID = r.TABLEID 
                        and f.FIELDID = r.FIELDID 
                        and r.LANGID = 'CHS'
                    ";
                sqlMSSql = sql + @" where f.TABLEID = @TABLEID order by f.fieldid";
                sqlOracle = sql + @" where f.TABLEID = :TABLEID order by f.fieldid";
                DataTable dtField = null;
                if (dbType == DBType.SqlServer)
                {
                    dtField = SQLHelper.GetDataTable(sqlMSSql, new SqlParameter[] { new SqlParameter("@TABLEID", tableid) });
                }
                else
                {
                    dtField = OracleHelper.ExecToSqlGetTable(sqlOracle, new OracleParameter[] { new OracleParameter(":TABLEID", tableid) });
                }

                if (dtField.Rows.Count == 0)
                {
                    bgw.ReportProgress(++iTable);
                    continue;
                }
                sb.AppendLine("| 字段名 | 字段描述 | 数据类型 | 长度 | 精度 | 默认值 | 可空? | 标题 |");
                sb.AppendLine("| ------ | ------------ | ------ | ------ | ------ | ------ | ------ | ------------ |");
                int rowIndex = 1;
                //循环字段
                foreach (DataRow drField in dtField.Rows)
                {
                    rowIndex++;
                    string name = drField["name"].ToString(); 
                    tableFields += (name + ", ");
                    tableFieldsUpdateSql += (name + "=?"+name+"?, ");
                    string fdescription = drField["description"].ToString();

                    #region type length scale
                    string type = drField["type"].ToString();
                    string length = drField["length"].ToString();
                    string scale = drField["scale"].ToString();
                    string typeAllMySql = "";
                    string typeAllOracle = "";
                    string typeAllSqlServer = "";
                    switch (type)
                    {
                        case "INTEGER":
                            typeAllMySql += " INT";
                            typeAllOracle += " INTEGER";
                            typeAllSqlServer += " INT";
                            break;
                        case "CHAR":
                            typeAllMySql += " CHAR(" + length + ")";
                            typeAllOracle += " NCHAR(" + length + ")";
                            typeAllSqlServer += " NCHAR(" + length + ")";
                            break;
                        case "VARCHAR":
                            typeAllMySql += " VARCHAR(" + length + ")";
                            typeAllOracle += " NVARCHAR2(" + length + ")";
                            typeAllSqlServer += " NVARCHAR(" + length + ")";
                            break;
                        case "DATE":
                            typeAllMySql += " DATETIME";
                            typeAllOracle += " DATE";
                            typeAllSqlServer += " DATETIME";
                            break;
                        case "DECIMAL":
                            typeAllMySql += " DECIMAL(" + length + "," + scale + ")";
                            typeAllOracle += " NUMBER(" + length + "," + scale + ")";
                            typeAllSqlServer += " NUMERIC(" + length + "," + scale + ")";
                            break;
                        case "LONGVARBINARY":
                            typeAllMySql += " BLOB";
                            typeAllOracle += " BLOB";
                            typeAllSqlServer += " BINARY";
                            break;
                        case "LONGVARCHAR":
                            typeAllMySql += " TEXT";
                            typeAllOracle += " CLOB";
                            typeAllSqlServer += " TEXT";
                            break;
                        default:
                            typeAllMySql = type;
                            break;
                    }
                    #endregion
                    string defaultvalue = drField["defaultvalue"].ToString();
                    string isnullable = drField["isnullable"].ToString();
                    string caption = drField["caption"].ToString();
                    sb.AppendLine("| "+name+ " | "+fdescription+ " | "+type + " | " + length+ " | " + scale  + " | "+defaultvalue + " | "+isnullable + " | "+caption + " |");
                    tableFieldsDefineMySql += (name + typeAllMySql);
                    tableFieldsDefineSqlServer += (name + typeAllSqlServer);
                    tableFieldsDefineOracle += (name + typeAllOracle);

                    #region 允许空
                    switch (isnullable)
                    {
                        case "Y":
                            tableFieldsDefineMySql += " NULL";
                            tableFieldsDefineSqlServer += " NULL";
                            tableFieldsDefineOracle += " NULL";
                            break;
                        case "N":
                            tableFieldsDefineMySql += " NOT NULL";
                            tableFieldsDefineSqlServer += " NOT NULL";
                            tableFieldsDefineOracle += " NOT NULL";
                            break; 
                    }
                    #endregion

                    #region 备注
                    if (!string.IsNullOrEmpty(caption)) {
                        tableFieldsDefineMySql += " COMMENT '"+caption+"' ";
                        tableFieldsCommentOracle += ("COMMENT ON COLUMN " + tablename + "." + name + " IS '" + caption + "';" + Environment.NewLine);
                    }
                    #endregion

                    #region 默认值
                    if (!string.IsNullOrEmpty(defaultvalue))
                    { 
                        //tableFieldsDefineMySql += " DEFAULT '" + defaultvalue + "'";
                        //tableFieldsDefineSqlServer += " DEFAULT '" + defaultvalue + "'";
                        //tableFieldsDefineOracle += " DEFAULT '" + defaultvalue + "'";
                    }
                    #endregion

                    tableFieldsDefineMySql += (","+ Environment.NewLine);
                    tableFieldsDefineSqlServer += ("," + Environment.NewLine);
                    tableFieldsDefineOracle += ("," + Environment.NewLine);
                }
                tableFields = tableFields.Substring(0, tableFields.Length - 2);
                tableFieldsUpdateSql = tableFieldsUpdateSql.Substring(0, tableFieldsUpdateSql.Length - 2);
                tableFieldsDefineMySql = tableFieldsDefineMySql.Substring(0, tableFieldsDefineMySql.Length - 3);
                tableFieldsDefineSqlServer = tableFieldsDefineSqlServer.Substring(0, tableFieldsDefineSqlServer.Length - 3);
                tableFieldsDefineOracle = tableFieldsDefineOracle.Substring(0, tableFieldsDefineOracle.Length - 3);
                sb.AppendLine();
                #endregion

                #region 索引定义
                sql = "select * from LIMSTABLEINDEXES i where i.tableid = {0}";
                DataTable dtIndex = null;
                if (dbType == DBType.SqlServer)
                {
                    dtIndex = SQLHelper.GetDataTable(string.Format(sql, "@TABLEID"), new SqlParameter[] { new SqlParameter("@TABLEID", tableid) });
                }
                else
                {
                    dtIndex = OracleHelper.ExecToSqlGetTable(string.Format(sql, ":TABLEID"), new OracleParameter[] { new OracleParameter(":TABLEID", tableid) });
                }
                if (dtIndex.Rows.Count > 0)
                {
                    head1Index++;
                    sb.AppendLine("# "+head1Index+" 索引定义");
                    sb.AppendLine();
                    head2Index = 0;
                }
                //循环各个索引
                foreach (DataRow drIndex in dtIndex.Rows)
                {
                    head2Index++;
                    string indexName = drIndex["INDEXNAME"].ToString();
                    string indexType = drIndex["INDEXTYPE"].ToString();
                    sb.AppendLine("## "+head1Index+"."+head2Index+" 索引 " + indexName);
                    sb.AppendLine();
                    sb.AppendLine("* 索引名称：" + indexName);
                    sb.AppendLine("* 索引类型：" + indexType);
                    sb.AppendLine("* 索引字段定义：" );
                    sb.AppendLine(); 

                    #region 索引字段定义
                    sql = @"SELECT F.NAME, IDF.SORT, IDF.INCLUDED from LIMSTABLEINDEXES i
                        LEFT JOIN LIMSTABLEINDEXFIELDS IDF ON I.TABLEID = IDF.TABLEID AND I.INDEXID = IDF.INDEXID
                        LEFT JOIN LIMSTABLEFIELDS F ON IDF.TABLEID = F.TABLEID AND IDF.FIELDID = F.FIELDID
                        where i.TABLEID = {0}     
                              and idf.indexid = {1}  ";
                    DataTable dtIndexField = null;
                    if (dbType == DBType.SqlServer)
                    {
                        dtIndexField = SQLHelper.GetDataTable(string.Format(sql, "@TABLEID", "@INDEXID")
                            , new SqlParameter[] { new SqlParameter("@TABLEID", tableid), new SqlParameter("@INDEXID", drIndex["INDEXID"].ToString()) });
                    }
                    else
                    {
                        dtIndexField = OracleHelper.ExecToSqlGetTable(string.Format(sql, new string[] { ":TABLEID", ":INDEXID" })
                       , new OracleParameter[] { new OracleParameter(":TABLEID", tableid), new OracleParameter(":INDEXID", drIndex["INDEXID"].ToString()) });
                    }
                    sb.AppendLine("| 字段名 | 排序 | 包含 |");
                    sb.AppendLine("| ------ | ------ | ------ |");  
                    //循环索引字段 
                    foreach (DataRow drIndexField in dtIndexField.Rows)
                    { 
                        sb.AppendLine("| " + drIndexField["NAME"].ToString() + " | " + drIndexField["SORT"].ToString() + " | " + drIndexField["INCLUDED"].ToString() + " |");
                    }
                    sb.AppendLine();
                    #endregion
                }
                sb.AppendLine();
                #endregion

                #region 外键约束定义
                sql = "select TR.*, T.TABLENAME" +
                    " from LIMSTABLERELATIONS TR " +
                    " LEFT JOIN LIMSTABLES T ON TR.REFTABLEID = T.TABLEID" +
                    " WHERE TR.TABLEID = {0} order by TR.FKID";
                DataTable dtRel = null;
                if (dbType == DBType.SqlServer)
                {
                    dtRel = SQLHelper.GetDataTable(string.Format(sql, "@TABLEID")
                    , new SqlParameter[] { new SqlParameter("@TABLEID", tableid) });
                }
                else
                {
                    dtRel = OracleHelper.ExecToSqlGetTable(string.Format(sql, ":TABLEID")
                    , new OracleParameter[] { new OracleParameter(":TABLEID", tableid) });
                }

                if (dtRel.Rows.Count > 0)
                {
                    head1Index++;
                    sb.AppendLine("# "+ head1Index + " 外键约束定义");
                    sb.AppendLine();
                    head2Index = 0;
                }
                //循环各个外键
                foreach (DataRow drRel in dtRel.Rows)
                {
                    head2Index++;
                    string fkName = drRel["FKNAME"].ToString();
                    sb.AppendLine("## "+head1Index+"."+head2Index+" 约束 " + fkName);
                    sb.AppendLine();
                    sb.AppendLine("* 关联表：" + drRel["TABLENAME"].ToString());
                    sb.AppendLine("* 级联更新：" + drRel["CASCADEUPDATE"].ToString());
                    sb.AppendLine("* 级联删除：" + drRel["CASCADEDELETE"].ToString());
                    sb.AppendLine("* 外键字段定义：" );
                    sb.AppendLine();   
                    #region 关联字段 
                    sql = @"SELECT TF.NAME FieldName, TFR.NAME RelFieldName 
                            FROM LIMSTABLERELATIONFIELDS TRF
                            --LEFT JOIN LIMSTABLES T ON TRF.TABLEID = T.TABLEID
                            --LEFT JOIN LIMSTABLES TR ON TRF.REFTABLEID = TR.TABLEID
                            LEFT JOIN LIMSTABLEFIELDS TF ON TRF.TABLEID = TF.TABLEID AND TRF.FIELDID = TF.FIELDID
                            LEFT JOIN LIMSTABLEFIELDS TFR ON TRF.REFTABLEID = TFR.TABLEID AND TRF.REFFIELDID = TFR.FIELDID
                            WHERE TRF.TABLEID = {0}
                                  AND TRF.FKID = {1}";
                    DataTable dtRelField = null;
                    if (dbType == DBType.SqlServer)
                    {
                        dtRelField = SQLHelper.GetDataTable(string.Format(sql, new string[] { "@TABLEID", "@FKID" })
                   , new SqlParameter[] { new SqlParameter("@TABLEID", tableid)
                                , new SqlParameter("@FKID", drRel["FKID"].ToString()) });
                    }
                    else
                    {
                        dtRelField = OracleHelper.ExecToSqlGetTable(string.Format(sql, new string[] { ":TABLEID", ":FKID" })
                   , new OracleParameter[] { new OracleParameter(":TABLEID", tableid)
                                , new OracleParameter(":FKID", drRel["FKID"].ToString()) });
                    }
                    sb.AppendLine("| 字段名 | 关联表字段名 |");
                    sb.AppendLine("| ------ | ------ |");  
                    //循环字段
                    foreach (DataRow drRelField in dtRelField.Rows)
                    {  
                        sb.AppendLine("| " + drRelField["FieldName"].ToString() + " | "+ drRelField["RelFieldName"].ToString() + " |");
                    }
                    sb.AppendLine();
                    #endregion
                }
                sb.AppendLine();
                #endregion

                #region 引用该表的子表定义
                sql = "select TR.*, T.TABLENAME" +
                    " from LIMSTABLERELATIONS TR " +
                    " LEFT JOIN LIMSTABLES T ON TR.TABLEID = T.TABLEID" +
                    " WHERE TR.REFTABLEID = {0} " +
                    "order by T.TABLENAME";
                DataTable dtRel2 = null;
                if (dbType == DBType.SqlServer)
                {
                    dtRel2 = SQLHelper.GetDataTable(string.Format(sql, "@TABLEID")
                   , new SqlParameter[] { new SqlParameter("@TABLEID", tableid) });
                }
                else
                {
                    dtRel2 = OracleHelper.ExecToSqlGetTable(string.Format(sql, ":TABLEID")
                   , new OracleParameter[] { new OracleParameter(":TABLEID", tableid) });
                }
                if (dtRel2.Rows.Count > 0)
                {
                    head1Index++;
                    sb.AppendLine("# "+ head1Index + " 子表定义");
                    sb.AppendLine();
                    head2Index = 0;
                }
                //循环各个子表
                foreach (DataRow drRel2 in dtRel2.Rows)
                {
                    head2Index++;
                    string subTableName = drRel2["TABLENAME"].ToString();
                    sb.AppendLine("## "+head1Index+"."+ head2Index + " 子表 " + subTableName);
                    sb.AppendLine();
                    sb.AppendLine("* 子表约束：" + drRel2["FKNAME"].ToString());
                    sb.AppendLine("* 级联更新：" + drRel2["CASCADEUPDATE"].ToString());
                    sb.AppendLine("* 级联删除：" + drRel2["CASCADEDELETE"].ToString());
                    sb.AppendLine("* 约束字段定义：");
                    sb.AppendLine();   

                    #region 关联字段   
                    sql = @"SELECT TF.NAME FieldName --子表字段
                               , TFR.NAME RelFieldName --父表字段 
                               --, TRF.* 
                        FROM LIMSTABLERELATIONFIELDS TRF --关联字段
                        --LEFT JOIN LIMSTABLES T ON TRF.TABLEID = T.TABLEID
                        --LEFT JOIN LIMSTABLES TR ON TRF.REFTABLEID = TR.TABLEID
                        LEFT JOIN LIMSTABLEFIELDS TF ON TRF.TABLEID = TF.TABLEID AND TRF.FIELDID = TF.FIELDID --子表字段
                        LEFT JOIN LIMSTABLEFIELDS TFR ON TRF.REFTABLEID = TFR.TABLEID AND TRF.REFFIELDID = TFR.FIELDID --父表字段
                        WHERE TRF.REFTABLEID = {0} --父表ID
                              AND TRF.TABLEID = {1} --子表ID";
                    DataTable dtRelField2 = null;
                    if (dbType == DBType.SqlServer)
                    {
                        dtRelField2 = SQLHelper.GetDataTable(string.Format(sql, new string[] { "@REFTABLEID", "@TABLEID" })
                   , new SqlParameter[] { new SqlParameter("@REFTABLEID", tableid)
                                , new SqlParameter("@TABLEID", drRel2["TABLEID"].ToString()) });
                    }
                    else
                    {
                        dtRelField2 = OracleHelper.ExecToSqlGetTable(string.Format(sql, new string[] { ":REFTABLEID", ":TABLEID" })
                   , new OracleParameter[] { new OracleParameter(":REFTABLEID", tableid)
                                , new OracleParameter(":TABLEID", drRel2["TABLEID"].ToString()) });
                    }
                    //  
                    sb.AppendLine("| 字段名 | 子表字段名 |");
                    sb.AppendLine("| ------ | ------ |");  
                    //循环字段
                    foreach (DataRow drRelField2 in dtRelField2.Rows)
                    {  
                        sb.AppendLine("| " + drRelField2["RelFieldName"].ToString() + " | " + drRelField2["FieldName"].ToString() + " |");
                    }
                    sb.AppendLine();
                    #endregion
                }
                sb.AppendLine();
                #endregion

                #region DDL语句生成
                head1Index++;
                sb.AppendLine("# " + head1Index + " SQL");
                sb.AppendLine();
                head2Index = 0;

                #region create mysql
                head2Index++;
                sb.AppendLine("## " + head1Index + "." + head2Index + " mysql craete 创建表语句");
                sb.AppendLine();
                sb.AppendLine("```sql");
                string createTableSql = "create table " + tablename + " (" 
                    + Environment.NewLine+ tableFieldsDefineMySql
                    + Environment.NewLine + ") ";
                if (!string.IsNullOrEmpty(description))
                {
                    createTableSql += " COMMENT '" + description + "'";
                }
                createTableSql += ";";
                sb.AppendLine(createTableSql);
                sb.AppendLine("```");
                sb.AppendLine();
                #endregion

                #region create sqlserver
                head2Index++;
                sb.AppendLine("## " + head1Index + "." + head2Index + " sql server craete 创建表语句");
                sb.AppendLine();
                sb.AppendLine("```sql");
                createTableSql = "create table " + tablename + " ("
                    + Environment.NewLine + tableFieldsDefineSqlServer
                    + Environment.NewLine + ") "; 
                createTableSql += ";";
                sb.AppendLine(createTableSql);
                sb.AppendLine("```");
                sb.AppendLine();
                #endregion

                #region create oracle
                head2Index++;
                sb.AppendLine("## " + head1Index + "." + head2Index + " oracle craete 创建表语句");
                sb.AppendLine();
                sb.AppendLine("```sql");
                createTableSql = "create table " + tablename + " ("
                    + Environment.NewLine + tableFieldsDefineOracle
                    + Environment.NewLine + ") ;"; 
                sb.AppendLine(createTableSql);
                sb.AppendLine(tableFieldsCommentOracle);//oracle备注
                if (!string.IsNullOrEmpty(description))
                {
                    sb.AppendLine("COMMENT ON TABLE "+tablename+" IS '"+description+"';");
                }
                sb.AppendLine("```");
                sb.AppendLine();
                #endregion

                #endregion

                #region DML语句生成
                head1Index++;
                sb.AppendLine("# " + head1Index + " SQL");
                sb.AppendLine();
                head2Index = 0;

                #region select 
                head2Index++;
                sb.AppendLine("## " + head1Index+"."+head2Index + " select 查询语句");
                sb.AppendLine();
                sb.AppendLine("```sql");
                sb.AppendLine("select "+ tableFields + Environment.NewLine+ "from "+ tablename + " ");
                sb.AppendLine("```");
                sb.AppendLine();
                #endregion

                #region insert 
                head2Index++;
                sb.AppendLine("## " + head1Index + "." + head2Index + " insert 添加语句");
                sb.AppendLine();
                sb.AppendLine("```sql");
                sb.AppendLine("insert into " + tablename + Environment.NewLine
                        + "(" + tableFields + ")" + Environment.NewLine
                        + "values" + Environment.NewLine
                        + "(?" + tableFields.Replace(", ","?, ?") + "?)");
                sb.AppendLine("```");
                sb.AppendLine();
                #endregion

                #region delete 
                head2Index++;
                sb.AppendLine("## " + head1Index + "." + head2Index + " delete 删除语句");
                sb.AppendLine();
                sb.AppendLine("```sql");
                sb.AppendLine("delete from " + tablename + Environment.NewLine + "where ");
                sb.AppendLine("```");
                sb.AppendLine();
                #endregion

                #region update 
                head2Index++;
                sb.AppendLine("## " + head1Index + "." + head2Index + " update 更新语句");
                sb.AppendLine();
                sb.AppendLine("```sql");
                sb.AppendLine("update " + tablename + " set " + Environment.NewLine 
                    + tableFieldsUpdateSql + Environment.NewLine
                    + "where ");
                sb.AppendLine("```");
                sb.AppendLine();
                #endregion

                #endregion

                string fileName = tablename + "-" + description.Replace("\r", "").Replace("\n", "").Replace("\b", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("<", "").Replace(">", "").Replace("|", "").Replace("\"", "").Replace("/", "").Replace("\\", "") + ".md";
                File.WriteAllText(fileName, sb.ToString());
                #region 报告进度
                bgw.ReportProgress(++iTable);
                #endregion
            }
            //
            
        }

        private void bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            string type = (string) e.Argument;
            if (type == "Word")
            {
                GenerateDesignDoc();
            }
            else
            {
                GenerateDesignMarkdown();
            } 
        }

        private void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        { 
            tspb.Value++;
        }

        private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        { 
            tsbtnExportTableDefine2Markdown.Enabled = true;
            tsbtnExportTableDefine2Word.Enabled = true;
            tssl.Text = "导出完成";
        }
         
        private void tsbtnExportTableDefine2Markdown_Click(object sender, EventArgs e)
        {
            ExportTableDefine("MarkDown");
        }

        private void tsbtnExportTableDefine2Word_Click(object sender, EventArgs e)
        {
            ExportTableDefine("Word");
        }

        private void ExportTableDefine(string type)
        { 
            tsbtnExportTableDefine2Markdown.Enabled = false;
            tsbtnExportTableDefine2Word.Enabled = false;
            //
            string sql = "select count(1) from limstables where issystem = 'N' " ;
            foreach(string tableFilter in tableFilterList)
            {
                sql += " and TABLENAME " +tableFilter;
            }
            int count = 1;
            if (dbType == DBType.SqlServer)
            {
                count = (int)SQLHelper.ExecuteScalar(sql);
            }
            else
            {
                count = int.Parse(OracleHelper.ExecToSqlGetTable(sql).Rows[0][0].ToString());
            } 
            tspb.Maximum = count;
            //
            bgw.RunWorkerAsync(type);
        }
    }
}
