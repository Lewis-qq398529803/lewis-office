package com.taozi.utils;

import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import fr.opensagres.xdocreport.template.formatter.FieldsMetadata;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.List;
import java.util.Map;

/**
 * 导出word utils
 *
 * @author taozi - 2021年7月27日, 027 - 10:32:41
 */
public class ExportWordUtils {

	/**
	 * 导出到浏览器
	 * @param templatePath
	 * @param data
	 * @param fileName
	 */
	public static void exportToBrowser(String templatePath, Map data, String fileName, HttpServletResponse response) {
		//获取Word模板，模板存放路径在项目的resources目录下
		//模板生成方式：选中文本按ctrl+f9,右键编辑域，选择 邮件合并.代码格式：${code}
		InputStream ins = ExportWordUtils.class.getResourceAsStream(templatePath);
		//注册xdocreport实例并加载FreeMarker模板引擎
		IXDocReport report = null;
		try {
			report = XDocReportRegistry.getRegistry().loadReport(ins, TemplateEngineKind.Freemarker);
			//创建xdocreport上下文对象
			IContext context = report.createContext();
			FieldsMetadata fm = report.createFieldsMetadata();
			//创建要替换的文本变量
			data.forEach((k, v) -> {
				try {
					context.put(k.toString(), v);
					if (v instanceof List) {
						//Word模板中的表格数据对应的集合类型
						fm.load(k.toString(), ((List<?>) v).get(0).getClass(), true);
					}
				} catch (XDocReportException e) {
					e.printStackTrace();
				}
			});

			//输出到本地目录
//			FileOutputStream out = new FileOutputStream(outPutPath);
//			report.process(context, out);
			//浏览器端下载
			response.setCharacterEncoding("utf-8");
			response.setContentType("application/msword");
			response.setHeader("Content-Disposition", "attachment;filename=".concat(String.valueOf(URLEncoder.encode(fileName, "UTF-8"))));
			report.process(context, response.getOutputStream());
		} catch (IOException | XDocReportException e) {
			e.printStackTrace();
		}
	}


	/**
	 * 导出到本地
	 * @param templatePath
	 * @param data
	 * @param outPutPath
	 */
	public static void exportToLocal(String templatePath, Map data, String outPutPath) {
		//获取Word模板，模板存放路径在项目的resources目录下
		//模板生成方式：选中文本按ctrl+f9,右键编辑域，选择 邮件合并.代码格式：${code}
		InputStream ins = ExportWordUtils.class.getResourceAsStream(templatePath);
		//注册xdocreport实例并加载FreeMarker模板引擎
		IXDocReport report = null;
		try {
			report = XDocReportRegistry.getRegistry().loadReport(ins, TemplateEngineKind.Freemarker);
			//创建xdocreport上下文对象
			IContext context = report.createContext();
			FieldsMetadata fm = report.createFieldsMetadata();
			//创建要替换的文本变量
			data.forEach((k, v) -> {
				try {
					context.put(k.toString(), v);
					if (v instanceof List) {
						//Word模板中的表格数据对应的集合类型
						fm.load(k.toString(), ((List<?>) v).get(0).getClass(), true);
					}
				} catch (XDocReportException e) {
					e.printStackTrace();
				}
			});
			//输出到本地目录
			File outPutPathIsExist = new File(outPutPath);
			//文件路径不存在则创建
			if (!outPutPathIsExist.exists()) {
				outPutPathIsExist.mkdirs();
			}
			FileOutputStream out = new FileOutputStream(outPutPath);
			report.process(context, out);
		} catch (IOException | XDocReportException e) {
			e.printStackTrace();
		}
	}
}
