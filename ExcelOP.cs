using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Collections;
using System.Text;

namespace ExcelCtr
{
    /// <summary>
    /// Excel操作类,用于控制excel的读取和写入
    /// </summary>
    public class ExcelOP
    {
        static ExcelOP()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
        #region 读取excel

        /// <summary>
        /// 将excel中的每一个表第一行为列名组合读取成一个dataset
        /// </summary>
        /// <param name="filePath">excel文件路径</param>
        /// <returns></returns>
        public static DataSet Read(string filePath)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(filePath);
            return ds;
        }

        /// <summary>
        /// 将excel中的每一个表第一行为列名组合读取成一个dataset
        /// </summary>
        /// <param name="stream">excel文件流</param>
        /// <returns></returns>
        public static DataSet Read(Stream stream)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(stream);
            return ds;
        }

        /// <summary>
        /// 读取excel中指定表名和指定相应列头所在索引行的表
        /// <para>如:Read("c:\test.xls",new List&lt;string&gt;() { "Sheet1", "Sheet2" },new List&lt;int&gt;() { 1, 1 })</para>
        /// </summary>
        /// <param name="filePath">excel文件路径</param>
        /// <param name="sheetNames">excel中的sheet名称列表</param>
        /// <param name="indexOfColNames">每个sheet对应的列头所在的行索引集合</param>
        /// <returns></returns>
        public static DataSet Read(string filePath, List<string> sheetNames, List<int> indexOfColNames)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(filePath, sheetNames, indexOfColNames);
            return ds;
        }

        /// <summary>
        /// 读取excel中指定Sheet名的Sheet,并且指定每个Sheet的列头所在的行索引
        /// <para>如:Read(stream,new List&lt;string&gt;() { "Sheet1", "Sheet2" },new List&lt;int&gt;() { 1, 1 })</para>
        /// </summary>
        /// <param name="stream">excel文件流</param>
        /// <param name="sheetNames">excel中的sheet名称列表</param>
        /// <param name="indexOfColNames">每个sheet对应的列头所在的行索引集合</param>
        /// <returns></returns>
        public static DataSet Read(Stream stream, List<string> sheetNames, List<int> indexOfColNames)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(stream, sheetNames, indexOfColNames);
            return ds;
        }

        /// <summary>
        /// 读取excel中指定Sheet索引的Sheet,并且指定每个Sheet的列头所在的行索引
        /// <para>如:Read("c:\test.xls",new List&lt;int&gt;() { 0, 1 },new List&lt;int&gt;() { 1, 1 })</para>
        /// </summary>
        /// <param name="filePath">excel文件路径</param>
        /// <param name="sheetIndexs">excel中的sheet索引列表</param>
        /// <param name="indexOfColNames">每个sheet对应的列头所在的行索引集合</param>
        /// <returns></returns>
        public static DataSet Read(string filePath, List<int> sheetIndexs, List<int> indexOfColNames)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(filePath, sheetIndexs, indexOfColNames);
            return ds;
        }

        /// <summary>
        /// 读取excel中指定索引的Sheet,并且指明每个Sheet列头的行索引
        /// <para>如:Read(stream,new List&lt;int&gt;() { 0, 1 },new List&lt;int&gt;() { 1, 1 })</para>
        /// </summary>
        /// <param name="stream">excel文件流</param>
        /// <param name="sheetIndexs">excel中的sheet索引列表</param>
        /// <param name="indexOfColNames">每个sheet对应的列头所在的行索引集合</param>
        /// <returns></returns>
        public static DataSet Read(Stream stream, List<int> sheetIndexs, List<int> indexOfColNames)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(stream, sheetIndexs, indexOfColNames);
            return ds;
        }

        /// <summary>
        /// 读取excel中指定索引的Sheet,并且指明每个Sheet是否有列头以及列头和数据的行索引
        /// <para>如:Read("c:\test.xls",new List&lt;int&gt;() { 0, 1 },new List&lt;bool&gt;() { true, true }, new List&lt;int[]&gt;() { new int[] { 0, 0 }, new int[] { 0,0} })</para>
        /// </summary>
        /// <param name="filePath">excel文件路径</param>
        /// <param name="sheetIndexs">excel中的sheet索引列表</param>
        /// <param name="hasColNames">每个sheet是否列头行的布尔说明</param>
        /// <param name="dataStartIndex">每个sheet的列头行号和数据内容起始的起始行索引</param>
        /// <returns></returns>
        public static DataSet Read(string filePath, List<int> sheetIndexs, List<bool> hasColNames, List<int[]> dataStartIndex)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(filePath, sheetIndexs, hasColNames, dataStartIndex);
            return ds;
        }

        /// <summary>
        /// 读取excel中指定索引的Sheet,并且指明每个Sheet是否有列头以及列头和数据的行索引
        /// <para>如:Read(stream,new List&lt;int&gt;() { 0, 1 },new List&lt;bool&gt;() { true, true }, new List&lt;int[]&gt;() { new int[] { 0, 0 }, new int[] { 0,0} })</para>
        /// </summary>
        /// <param name="stream">excel文件流</param>
        /// <param name="sheetIndexs">excel中的sheet索引列表</param>
        /// <param name="hasColNames">每个sheet是否列头行的布尔说明</param>
        /// <param name="dataStartIndex">每个sheet的列头行号和数据内容起始的起始行索引</param>
        /// <returns></returns>
        public static DataSet Read(Stream stream, List<int> sheetIndexs, List<bool> hasColNames, List<int[]> dataStartIndex)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(stream, sheetIndexs, hasColNames, dataStartIndex);
            return ds;
        }

        /// <summary>
        /// 读取excel中指定Sheet名的Sheet,并且指明每个Sheet是否有列头以及列头和数据的行索引
        /// <para>如:Read("c:\test.xls",new List&lt;string&gt;() { "Sheet1", "Sheet2" },new List&lt;bool&gt;() { true, true }, new List&lt;int[]&gt;() { new int[] { 0, 0 }, new int[] { 0,0} })</para>
        /// </summary>
        /// <param name="filePath">excel文件路径</param>
        /// <param name="sheetNames">excel中的sheet名称列表</param>
        /// <param name="hasColNames">每个sheet是否列头行的布尔说明</param>
        /// <param name="dataStartIndex">每个sheet的列头行号和数据内容起始的起始行索引</param>
        /// <returns></returns>
        public static DataSet Read(string filePath, List<string> sheetNames, List<bool> hasColNames, List<int[]> dataStartIndex)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(filePath, sheetNames, hasColNames, dataStartIndex);
            return ds;
        }

        /// <summary>
        /// 读取excel中指定Sheet名的Sheet,并且指明每个Sheet是否有列头以及列头和数据的行索引
        /// <para>如:Read(stream,new List&lt;string&gt;() { "Sheet1", "Sheet2" },new List&lt;bool&gt;() { true, true }, new List&lt;int[]&gt;() { new int[] { 0, 0 }, new int[] { 0,0} })</para>
        /// </summary>
        /// <param name="stream">excel文件流</param>
        /// <param name="sheetNames">excel中的sheet名称列表</param>
        /// <param name="hasColNames">每个sheet是否列头行的布尔说明</param>
        /// <param name="dataStartIndex">每个sheet的列头行号和数据内容起始的起始行索引</param>
        /// <returns></returns>
        public static DataSet Read(Stream stream, List<string> sheetNames, List<bool> hasColNames, List<int[]> dataStartIndex)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(stream, sheetNames, hasColNames, dataStartIndex);
            return ds;
        }
        #endregion

        #region 写入excel

        /// <summary>
        /// 将ds数据写入excel文件中
        /// </summary>
        /// <param name="filePath">生成excel文件的路径</param>
        /// <param name="ds">生成使用的数据集</param>
        public static void Write(string filePath, DataSet ds)
        {
            Write(filePath, ds, null);
        }

        /// <summary>
        /// 将ds数据写入excel文件中
        /// </summary>
        /// <param name="filePath">生成excel文件的路径</param>
        /// <param name="ds">生成使用的数据集</param>
        /// <param name="SheetHeaders">每个sheet的表头集合(顺序和ds的table对应)</param>
        public static void Write(string filePath, DataSet ds, List<string> SheetHeaders)
        {
            FileStream fs = new FileStream(filePath, FileMode.Create);
            Write(fs, ds, SheetHeaders, filePath.EndsWith(".xlsx") ? EnumExcelType.openxml : EnumExcelType.office2003);
        }

        /// <summary>
        /// 将ds数据写入文件流中
        /// </summary>
        /// <param name="fs">目的文件流</param>
        /// <param name="ds">生成使用的数据集</param>
        /// <param name="enumExcelType">excel文件格式</param>
        public static void Write(FileStream fs, DataSet ds, EnumExcelType enumExcelType)
        {
            Write(fs, ds, null, enumExcelType);
        }

        /// <summary>
        /// 将ds数据写入文件流中
        /// </summary>
        /// <param name="fs">目的文件流</param>
        /// <param name="ds">生成使用的数据集</param>
        /// <param name="SheetHeaders">每个sheet的表头集合(顺序和ds的table对应)</param>
        /// <param name="enumExcelType">excel文件格式</param>
        public static void Write(FileStream fs, DataSet ds, List<string> SheetHeaders, EnumExcelType enumExcelType)
        {
            Write(fs, ds, SheetHeaders, new List<string>(), enumExcelType);
        }

        /// <summary>
        /// 将ds数据写入文件流中并指定合并行信息
        /// </summary>
        /// <param name="fs">目的文件流</param>
        /// <param name="ds">生成使用的数据集</param>
        /// <param name="SheetHeaders">每个sheet的表头集合(顺序和ds的table对应)</param>
        /// <param name="combineColIndexs">要进行纵向合并的列索引集合</param>
        /// <param name="enumExcelType">excel文件格式</param>
        public static void Write(FileStream fs, DataSet ds, List<string> SheetHeaders, List<string> combineColIndexs, EnumExcelType enumExcelType)
        {
            MemoryStream stream = ExcelHelper.ExportDS(ds, SheetHeaders, combineColIndexs, enumExcelType);
            byte[] bs = stream.ToArray();
            fs.Write(bs, 0, bs.Length);
            fs.Flush();
            fs.Close();
        }

        /// <summary>
        /// 将ds数据写入excel文件中并指定合并行信息
        /// </summary>
        /// <param name="filePath">生成excel文件的路径</param>
        /// <param name="ds">生成使用的数据集</param>
        /// <param name="SheetHeaders">每个sheet的表头集合(顺序和ds的table对应)</param>
        /// <param name="combineColIndexs">要进行纵向合并的列索引集合</param>
        public static void Write(string filePath, DataSet ds, List<string> SheetHeaders, List<string> combineColIndexs)
        {
            FileStream fs = new FileStream(filePath, FileMode.Create);
            Write(fs, ds, SheetHeaders, combineColIndexs, filePath.EndsWith(".xlsx") ? EnumExcelType.openxml : EnumExcelType.office2003);
        }

        /// <summary>
        /// 根据模板导出excel
        /// </summary>
        /// <param name="ht">传进去的参数</param>
        /// <param name="templateConfPath">模板配置文件的绝对路径,后缀名为.xml,注意仅支持97-2003格式Excel</param>
        /// <param name="destfilepath">生成的Excel路径</param>
        public static void WriteWithTemplate(Hashtable ht, string templateConfPath, string destfilepath)
        {
            ExcelTemplateOP op = new ExcelTemplateOP(templateConfPath, ht);
            op.Write(destfilepath);
        }

        #endregion
    }
}