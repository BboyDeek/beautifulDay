package com.example.demo.util;

import com.sun.xml.internal.messaging.saaj.packaging.mime.internet.MimeUtility;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.util.StringUtils;
import javax.servlet.http.HttpServletRequest;
import java.net.URLEncoder;
import java.util.List;

public class ExcelUtil {

	/**
	 * 样式设置
	 */
	public static HSSFCellStyle createCellStyle(HSSFWorkbook workbook) {
		// *生成一个样式
		HSSFCellStyle style = workbook.createCellStyle();
		// 设置这些样式
		// 前景色
		style.setFillForegroundColor(HSSFColor.WHITE.index);
		// 背景色
		style.setFillBackgroundColor(HSSFColor.WHITE.index);
		// 填充样式
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		// 设置底边框
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		// 设置底边框颜色
		style.setBottomBorderColor(HSSFColor.WHITE.index);
		// 设置左边框
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		// 设置左边框颜色
		style.setLeftBorderColor(HSSFColor.BLACK.index);
		// 设置右边框
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);
		// 设置右边框颜色
		style.setRightBorderColor(HSSFColor.BLACK.index);
		// 设置顶边框
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);
		// 设置顶边框颜色
		style.setTopBorderColor(HSSFColor.BLACK.index);
		// 设置自动换行
		style.setWrapText(false);
		// 设置水平对齐的样式为居中对齐
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		// 设置垂直对齐的样式为居中对齐
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		// HSSFFont font = createCellFont(workbook);
		// // 把字体应用到当前的样式
		// style.setFont(font);
		return style;
	}

	/**
	 * 设置下载文件中文件的名称
	 *
	 * @param filename
	 * @param request
	 * @return
	 */
	public static String encodeFilename(String filename, HttpServletRequest request) {
		/**
		 * 获取客户端浏览器和操作系统信息 在IE浏览器中得到的是：User-Agent=Mozilla/4.0 (compatible; MSIE
		 * 6.0; Windows NT 5.1; SV1; Maxthon; Alexa Toolbar)
		 * 在Firefox中得到的是：User-Agent=Mozilla/5.0 (Windows; U; Windows NT 5.1;
		 * zh-CN; rv:1.7.10) Gecko/20050717 Firefox/1.0.6
		 */
		String agent = request.getHeader("USER-AGENT");
		try {

			if ((agent != null) && (-1 != agent.indexOf("MSIE"))) {
				String newFileName = URLEncoder.encode(filename, "UTF-8");
				newFileName = StringUtils.replace(newFileName, "+", "%20");
				if (newFileName.length() > 150) {
					newFileName = new String(filename.getBytes("GB2312"), "ISO8859-1");
					newFileName = StringUtils.replace(newFileName, " ", "%20");
				}
				return newFileName;
			}
			if ((agent != null) && (-1 != agent.indexOf("Mozilla")))
				return MimeUtility.encodeText(filename, "UTF-8", "B");

			return filename;
		} catch (Exception ex) {
			return filename;
		}
	}
	/**
	 * 添加excel的sheet
	 * @param workbook 工作簿
	 * @param head excle的头部
	 * @param list 2维数组,excle的内容
	 * @param sheetName sheet的名称
	 */
	public static XSSFWorkbook createExcelSheet(XSSFWorkbook workbook, String[] head, List<List<Object>> list, String sheetName) {
		if (workbook == null) {
			workbook = new XSSFWorkbook();
		}
		// 生成一个表格
		XSSFSheet sheet = workbook.createSheet(sheetName);
		int rowIndex = 0;

		if (ArrayUtils.isNotEmpty(head)) {
			XSSFRow row = sheet.createRow(rowIndex);
			rowIndex++;

			// 设置表格头部
			for (int i = 0; i < head.length; i++) {
				// 创建单元格
				XSSFCell cell = row.createCell(i);
				// 获取内容
				XSSFRichTextString text = new XSSFRichTextString(head[i]);
				// 设置单元格内容
				cell.setCellValue(text);
				// 指定单元格格式：数值、公式或字符串
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			}
		}

		// 遍历集合数据,产生表格数据行
		for (int i=0; i<list.size(); i++) {
			// 创建新行(row)
			XSSFRow row = sheet.createRow(i + rowIndex);
			// 获取内容
			List<Object> valueList = list.get(i);
			for (int j=0; j<valueList.size(); j++) {
				// 创建单元格
				XSSFCell cell = row.createCell(j);
				Object obj = valueList.get(j);
				setCellValue(cell, obj);
			}
		}

		return workbook;
	}

	/**
	 * 设置单元格的格式和值
	 * @param cell
	 * @param obj
	 */
	private static void setCellValue(XSSFCell cell, Object obj) {
		if (obj instanceof String) {
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			cell.setCellValue((String) obj);
		} else if (obj instanceof Integer) {
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			cell.setCellValue((Integer) obj);
		} else if (obj instanceof Double) {
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			cell.setCellValue((Double) obj);
		} else if (obj instanceof Long) {
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			cell.setCellValue((Long) obj);
		} else {
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			cell.setCellValue(String.valueOf(obj));
		}
	}

}
