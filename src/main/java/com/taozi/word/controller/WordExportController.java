package com.taozi.word.controller;

import com.taozi.utils.ExportWordUtils;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.HashMap;
import java.util.Map;

/**
 * word导出controller
 *
 * @author taozi - 2021年8月17日, 017 - 17:13:12
 */
@Api(tags = "word导出")
@RestController
@RequestMapping("/word")
public class WordExportController {

	@ApiOperation("word导出到浏览器（浏览器下载）")
	@GetMapping("/exportToBrowser")
	public void exportToBrowser(HttpServletResponse response) {
		// 通过模板编号队对应模板路径
		String templatePathNum = "1";
		String templatePath = "/wordTemplates/" + templatePathNum + ".docx";
		Map data = new HashMap();
		data.put("title", "title");
		data.put("content", "content");
		data.put("date", "date");
		ExportWordUtils.exportToBrowser(templatePath, data, "title.docx", response);
	}

	@ApiOperation("word导出到浏览器（浏览器下载）")
	@GetMapping("/exportToLocal")
	public void exportToLocal() {
		// 通过模板编号队对应模板路径
		String templatePathNum = "1";
		String templatePath = "/wordTemplates/" + templatePathNum + ".docx";
		Map data = new HashMap();
		data.put("title", "title");
		data.put("content", "content");
		data.put("date", "date");
		ExportWordUtils.exportToLocal(templatePath, data, "D:/");
	}
}
