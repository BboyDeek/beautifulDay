package com.example.demo.controller;

import com.example.demo.util.ExcelUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;


@Controller
@RequestMapping("/mail")
public class downloadExampleController {
	/**
	 *下载示例
	 *@param:
	 *@return:
	 *@Author:zym
	 *@Date:2019/6/3
	 */
	@RequestMapping("/downloadExample.do")
	public void downloadExample( String fileExt, HttpServletRequest request,
								HttpServletResponse response) throws ServletException, IOException {
		response.reset();
		String fileName = "联系人地址示例";
	    if (fileExt != null && (fileExt.equalsIgnoreCase("xlsx"))) {
			fileName = fileName + ".xlsx";
		} else {
			fileName = fileName + ".txt";
		}

		// 表头
		StringBuilder data = new StringBuilder();
	    //equalsIgnoreCase方法比较两个字符串是否相等并且忽略大小写
		if (StringUtils.equalsIgnoreCase(fileExt, "xlsx")) {
			// Excel文件表头
			String[] excelHead = {"邮箱", "昵称", "手机"};
			XSSFWorkbook workbook = new XSSFWorkbook();
			//行
			List<List<Object>> rowList = new ArrayList<>();
			//列
			List<Object> columnList1 = new ArrayList<>();
				//Excel第一行，第一列
				columnList1.add("example@example.com");
				List<Object> columnList2=new ArrayList<>();
				//Excel第二行，第一列
				columnList2.add("example@example.com");
				//Excel第二行，第二列
				columnList2.add("姓名");
				List<Object> columnList3=new ArrayList<>();
				//Excel第三行，第一列
				columnList3.add("example@example.com");
				//Excel第三行，第二列
				columnList3.add("姓名1");
				//Excel第三行，第三列
				columnList3.add("13112345678");

				rowList.add(columnList1);
				rowList.add(columnList2);
				rowList.add(columnList3);


			ExcelUtil.createExcelSheet(workbook, excelHead, rowList, fileName.substring(0, fileName.lastIndexOf(".")));
			ByteArrayOutputStream os = new ByteArrayOutputStream();
			workbook.write(os);
			fileName = ExcelUtil.encodeFilename(fileName, request);
			this.download(fileName, os.toByteArray(), request, response);
		} else {
				data.append("邮箱，昵称，手机").append("\r\n");
				data.append("example@example.org").append("\r\n");
				data.append("example1@example.org,姓名1").append("\r\n");
				data.append("example2@example.org,姓名2,13112345678").append("\r\n");
			fileName = ExcelUtil.encodeFilename(fileName, request);
			this.download(fileName, data, request, response);
		}
	}


	//下载Excel文件
	public void download(String fileName, byte[] data, HttpServletRequest request, HttpServletResponse response) throws IOException {
		try {
			OutputStream out = response.getOutputStream();
			response.setContentType("application/octet-stream");
			response.setHeader("Content-disposition", "attachment;filename=" + fileName);
			out.write(data);
			out.flush();
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	//下载文本文件
	public void download(String fileName, StringBuilder data, HttpServletRequest request, HttpServletResponse response) throws IOException {
		OutputStream ouputStream = response.getOutputStream();
		try {
			response.setContentType("application/csv;charset=GBK");
			request.setCharacterEncoding("GBK");
			response.setHeader("Content-disposition", "attachment;filename=" + fileName);
			ouputStream.write(data.toString().getBytes("GBK"));
			ouputStream.flush();
			ouputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
