using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Web;
using NPOI.HPSF;
using NPOI.XSSF.UserModel;

namespace Common
{
    /// <summary>
    /// Excel操作类
    /// </summary>
    public class ExcelHelper
    {
        #region Excel导入

        /// <summary>
        /// 从Excel取数据并记录到List集合里
        /// </summary>
        /// <param name="cellHeader">单元头的值和名称：{ { "UserName", "姓名" }, { "Age", "年龄" } };</param>
        /// <param name="filePath">保存文件绝对路径</param>
        /// <param name="errorMsg">错误信息</param>
        /// <param name="startIndex">数据行开始序列，默认为1（即第二列，从0开始）</param>
        /// <returns>转换后的List对象集合</returns>
        public static List<T> ExcelToEntityList<T>(Dictionary<string, string> cellHeader, string filePath, out StringBuilder errorMsg, int startIndex = 1) where T : new()
        {
            List<T> enlist = new List<T>();
            errorMsg = new StringBuilder();
            try
            {
                if (Regex.IsMatch(filePath, ".xls$")) // 2003
                {
                    enlist = Excel2003ToEntityList<T>(cellHeader, filePath, out errorMsg, startIndex);
                }
                else if (Regex.IsMatch(filePath, ".xlsx$")) // 2007
                {
                    enlist = Excel2007ToEntityList<T>(cellHeader, filePath, out errorMsg, startIndex);
                }
                return enlist;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 从Excel2003取数据并记录到List集合里
        /// </summary>
        /// <param name="cellHeader">单元头的Key和Value：{ { "UserName", "姓名" }, { "Age", "年龄" } };</param>
        /// <param name="filePath">保存文件绝对路径</param>
        /// <param name="errorMsg">错误信息</param>
        /// <returns>转换好的List对象集合</returns>
        private static List<T> Excel2003ToEntityList<T>(Dictionary<string, string> cellHeader, string filePath, out StringBuilder errorMsg, int startIndex = 1) where T : new()
        {
            errorMsg = new StringBuilder(); // 错误信息,Excel转换到实体对象时，会有格式的错误信息
            List<T> enlist = new List<T>(); // 转换后的集合
            try
            {
                using (FileStream fs = File.OpenRead(filePath))
                {
                    HSSFWorkbook workbook = new HSSFWorkbook(fs);
                    HSSFSheet sheet = (HSSFSheet)workbook.GetSheetAt(0); // 获取此文件第一个Sheet页
                    for (int rowIndex = startIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        // 1.判断当前行是否空行，若空行就不在进行读取下一行操作，结束Excel读取操作
                        IRow row = sheet.GetRow(rowIndex);
                        if (row == null)
                        {
                            break;
                        }
                        // 2.每一个Excel row转换为一个实体对象
                        T en = new T();
                        ExcelRowToEntity<T>(cellHeader, row, rowIndex, en, ref errorMsg);

                        enlist.Add(en);
                    }
                }
                return enlist;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 从Excel2007取数据并记录到List集合里
        /// </summary>
        /// <param name="cellHeader">单元头的Key和Value：{ { "UserName", "姓名" }, { "Age", "年龄" } };</param>
        /// <param name="filePath">保存文件绝对路径</param>
        /// <param name="errorMsg">错误信息</param>
        /// <param name="startIndex">数据行开始序列，默认为1（即第二列，从0开始）</param>
        /// <returns>转换好的List对象集合</returns>
        private static List<T> Excel2007ToEntityList<T>(Dictionary<string, string> cellHeader, string filePath, out StringBuilder errorMsg, int startIndex = 1) where T : new()
        {
            errorMsg = new StringBuilder(); // 错误信息,Excel转换到实体对象时，会有格式的错误信息
            List<T> enlist = new List<T>(); // 转换后的集合
            try
            {
                using (FileStream fs = File.OpenRead(filePath))
                {
                    XSSFWorkbook workbook = new XSSFWorkbook(fs);
                    XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(0); // 获取此文件第一个Sheet页
                    for (int rowIndex = startIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        // 1.判断当前行是否空行，若空行就不在进行读取下一行操作，结束Excel读取操作
                        IRow row = sheet.GetRow(rowIndex);
                        if (row == null)
                        {
                            break;
                        }
                        // 2.每一个Excel row转换为一个实体对象
                        T en = new T();
                        ExcelRowToEntity<T>(cellHeader, row, rowIndex, en, ref errorMsg);
                        enlist.Add(en);
                    }
                }
                return enlist;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion Excel导入

        #region Excel导出

        /// <summary>
        /// 实体类集合导出到EXCLE2003
        /// </summary>
        /// <param name="cellHeader">单元头的Key和Value：{ { "UserName", "姓名" }, { "Age", "年龄" } };</param>
        /// <param name="enList">数据源</param>
        /// <param name="sheetName">工作表名称</param>
        /// <returns>文件的下载地址</returns>
        public static string EntityListToExcel2003(Dictionary<string, string> cellHeader, IList enList, string sheetName)
        {
            try
            {
                string fileName = sheetName + "-" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xls"; // 文件名称
                string urlPath = "UpFiles/ExcelFiles/" + fileName; // 文件下载的URL地址，供给前台下载
                string filePath = HttpContext.Current.Server.MapPath("\\" + urlPath); // 文件路径

                // 1.检测是否存在文件夹，若不存在就建立个文件夹
                string directoryName = Path.GetDirectoryName(filePath);
                if (!Directory.Exists(directoryName))
                {
                    Directory.CreateDirectory(directoryName);
                }

                // 2.解析单元格头部，设置单元头的中文名称
                HSSFWorkbook workbook = new HSSFWorkbook(); // 工作簿
                ISheet sheet = workbook.CreateSheet(sheetName); // 工作表
                IRow row = sheet.CreateRow(0);
                List<string> keys = cellHeader.Keys.ToList();
                for (int i = 0; i < keys.Count; i++)
                {
                    row.CreateCell(i).SetCellValue(cellHeader[keys[i]]); // 列名为Key的值
                }

                // 3.List对象的值赋值到Excel的单元格里
                int rowIndex = 1; // 从第二行开始赋值(第一行已设置为单元头)
                foreach (var en in enList)
                {
                    IRow rowTmp = sheet.CreateRow(rowIndex);
                    for (int i = 0; i < keys.Count; i++) // 根据指定的属性名称，获取对象指定属性的值
                    {
                        string cellValue = ""; // 单元格的值
                        object properotyValue = null; // 属性的值
                        System.Reflection.PropertyInfo properotyInfo = null; // 属性的信息

                        // 3.1 若属性头的名称包含'.',就表示是子类里的属性，那么就要遍历子类，eg：UserEn.UserName
                        if (keys[i].IndexOf(".") >= 0)
                        {
                            // 3.1.1 解析子类属性(这里只解析1层子类，多层子类未处理)
                            string[] properotyArray = keys[i].Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                            string subClassName = properotyArray[0]; // '.'前面的为子类的名称
                            string subClassProperotyName = properotyArray[1]; // '.'后面的为子类的属性名称
                            System.Reflection.PropertyInfo subClassInfo = en.GetType().GetProperty(subClassName); // 获取子类的类型
                            if (subClassInfo != null)
                            {
                                // 3.1.2 获取子类的实例
                                var subClassEn = en.GetType().GetProperty(subClassName).GetValue(en, null);
                                // 3.1.3 根据属性名称获取子类里的属性类型
                                properotyInfo = subClassInfo.PropertyType.GetProperty(subClassProperotyName);
                                if (properotyInfo != null)
                                {
                                    properotyValue = properotyInfo.GetValue(subClassEn, null); // 获取子类属性的值
                                }
                            }
                        }
                        else
                        {
                            // 3.2 若不是子类的属性，直接根据属性名称获取对象对应的属性
                            properotyInfo = en.GetType().GetProperty(keys[i]);
                            if (properotyInfo != null)
                            {
                                properotyValue = properotyInfo.GetValue(en, null);
                            }
                        }

                        // 3.3 属性值经过转换赋值给单元格值
                        if (properotyValue != null)
                        {
                            cellValue = properotyValue.ToString();
                            // 3.3.1 对时间初始值赋值为空
                            if (cellValue.Trim() == "0001/1/1 0:00:00" || cellValue.Trim() == "0001/1/1 23:59:59")
                            {
                                cellValue = "";
                            }
                        }

                        // 3.4 填充到Excel的单元格里
                        rowTmp.CreateCell(i).SetCellValue(cellValue);
                    }
                    rowIndex++;
                }

                // 4.生成文件
                FileStream file = new FileStream(filePath, FileMode.Create);
                workbook.Write(file);
                file.Close();

                // 5.返回下载路径
                return urlPath;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion Excel导出

        #region Common

        /// <summary>
        /// Excel row转换为实体对象
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="cellHeader">单元头的Key和Value：{ { "UserName", "姓名" }, { "Age", "年龄" } };</param>
        /// <param name="row">Excel row</param>
        /// <param name="rowIndex">row index</param>
        /// <param name="en">实体</param>
        /// <param name="errorMsg">错误信息</param>
        private static void ExcelRowToEntity<T>(Dictionary<string, string> cellHeader, IRow row, int rowIndex, T en, ref StringBuilder errorMsg)
        {
            List<string> keys = cellHeader.Keys.ToList(); // 要赋值的实体对象属性名称
            string errStr = ""; // 当前行转换时，是否有错误信息，格式为：第1行数据转换异常：XXX列；
            for (int i = 0; i < keys.Count; i++)
            {
                // 1.若属性头的名称包含'.',就表示是子类里的属性，那么就要遍历子类，eg：UserEn.TrueName
                if (keys[i].IndexOf(".") >= 0)
                {
                    // 1)解析子类属性
                    string[] properotyArray = keys[i].Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                    string subClassName = properotyArray[0]; // '.'前面的为子类的名称
                    string subClassProperotyName = properotyArray[1]; // '.'后面的为子类的属性名称
                    System.Reflection.PropertyInfo subClassInfo = en.GetType().GetProperty(subClassName); // 获取子类的类型
                    if (subClassInfo != null)
                    {
                        // 2)获取子类的实例
                        var subClassEn = en.GetType().GetProperty(subClassName).GetValue(en, null);
                        // 3)根据属性名称获取子类里的属性信息
                        System.Reflection.PropertyInfo properotyInfo = subClassInfo.PropertyType.GetProperty(subClassProperotyName);
                        if (properotyInfo != null)
                        {
                            try
                            {
                                // Excel单元格的值转换为对象属性的值，若类型不对，记录出错信息
                                properotyInfo.SetValue(subClassEn, GetExcelCellToProperty(properotyInfo.PropertyType, row.GetCell(i)), null);
                            }
                            catch (Exception e)
                            {
                                if (errStr.Length == 0)
                                {
                                    errStr = "第" + rowIndex + "行数据转换异常：";
                                }
                                errStr += cellHeader[keys[i]] + "列；";
                            }

                        }
                    }
                }
                else
                {
                    // 2.给指定的属性赋值
                    System.Reflection.PropertyInfo properotyInfo = en.GetType().GetProperty(keys[i]);
                    if (properotyInfo != null)
                    {
                        try
                        {
                            // Excel单元格的值转换为对象属性的值，若类型不对，记录出错信息
                            properotyInfo.SetValue(en, GetExcelCellToProperty(properotyInfo.PropertyType, row.GetCell(i)), null);
                        }
                        catch (Exception e)
                        {
                            if (errStr.Length == 0)
                            {
                                errStr = "第" + rowIndex + "行数据转换异常：";
                            }
                            errStr += cellHeader[keys[i]] + "列；";
                        }
                    }
                }
            }
            // 若有错误信息，就添加到错误信息里
            if (errStr.Length > 0)
            {
                errorMsg.AppendLine(errStr);
            }
        }

        /// <summary>
        /// Excel Cell转换为实体的属性值
        /// </summary>
        /// <param name="distanceType">目标对象类型</param>
        /// <param name="sourceCell">对象属性的值</param>
        private static Object GetExcelCellToProperty(Type distanceType, ICell sourceCell)
        {
            object rs = distanceType.IsValueType ? Activator.CreateInstance(distanceType) : null;

            // 1.判断传递的单元格是否为空
            if (sourceCell == null || string.IsNullOrEmpty(sourceCell.ToString()))
            {
                return rs;
            }

            // 2.Excel文本和数字单元格转换，在Excel里文本和数字是不能进行转换，所以这里预先存值
            object sourceValue = null;
            switch (sourceCell.CellType)
            {
                case CellType.Blank:
                    break;

                case CellType.Boolean:
                    break;

                case CellType.Error:
                    break;

                case CellType.Formula:
                    break;

                case CellType.Numeric:
                    sourceValue = sourceCell.NumericCellValue;
                    break;

                case CellType.String:
                    sourceValue = sourceCell.StringCellValue;
                    break;

                case CellType.Unknown:
                    break;

                default:
                    break;
            }

            string valueDataType = distanceType.Name;

            // 在这里进行特定类型的处理
            switch (valueDataType.ToLower()) // 以防出错，全部小写
            {
                case "string":
                    rs = sourceValue.ToString();
                    break;
                case "int":
                case "int16":
                case "int32":
                    rs = (int)Convert.ChangeType(sourceCell.NumericCellValue.ToString(), distanceType);
                    break;
                case "float":
                case "single":
                    rs = (float)Convert.ChangeType(sourceCell.NumericCellValue.ToString(), distanceType);
                    break;
                case "datetime":
                    rs = sourceCell.DateCellValue;
                    break;
                case "guid":
                    rs = (Guid)Convert.ChangeType(sourceCell.NumericCellValue.ToString(), distanceType);
                    return rs;
            }
            return rs;
        }

        #endregion

        /// <summary>
        /// 保存Excel文件
        /// <para>Excel的导入导出都会在服务器生成一个文件</para>
        /// <para>路径：UpFiles/ExcelFiles</para>
        /// </summary>
        /// <param name="file">传入的文件对象</param>
        /// <returns>如果保存成功则返回文件的位置;如果保存失败则返回空</returns>
        public static string SaveExcelFile(HttpPostedFile file)
        {
            try
            {
                var fileName = file.FileName.Insert(file.FileName.LastIndexOf('.'), "-" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                var filePath = Path.Combine(HttpContext.Current.Server.MapPath("~/UpFiles/ExcelFiles"), fileName);
                string directoryName = Path.GetDirectoryName(filePath);
                if (!Directory.Exists(directoryName))
                {
                    Directory.CreateDirectory(directoryName);
                }
                file.SaveAs(filePath);
                return filePath;
            }
            catch
            {
                return string.Empty;
            }
        }


    }
}