package com.converter.controller;

import java.io.File;
import java.nio.file.Files;
import java.util.Base64;
import java.util.Map;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang.RandomStringUtils;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.aspose.cells.Workbook;
import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;
import com.aspose.words.Document;
import com.converter.util.FileUploadUtil;

@RestController
@RequestMapping("/converter")
public class ConverterController {


	@PostMapping("/upload_file")
	public Map<String, Object> uploadFile(@RequestParam("file") MultipartFile multipartFile) throws Exception {
		String random = RandomStringUtils.randomAlphabetic(5);
		String fileNameWithotExtention = FilenameUtils.removeExtension(multipartFile.getOriginalFilename());
		String fileName = fileNameWithotExtention + "-" + random;
		String extension = FilenameUtils.getExtension(multipartFile.getOriginalFilename());
		String fileWithExtention = fileName + "." + extension;
		String dataDir = FileUploadUtil.getDirectory();
		String fileInput = dataDir + "/" + fileName + "." + extension;
		String fileOutput = dataDir + "/" + fileName + "." + "pdf";

		FileUploadUtil.saveFile(fileWithExtention, multipartFile);

		if (extension.equalsIgnoreCase("pptx") || extension.equalsIgnoreCase("ppt")) {
			Presentation presentation = new Presentation(fileInput);
			presentation.save(fileOutput, SaveFormat.Pdf);
		} else if (extension.equalsIgnoreCase("xlsx") || extension.equalsIgnoreCase("xls")) {
			Workbook workbook = new Workbook(fileInput);
			workbook.save(fileOutput, SaveFormat.Pdf);
		} else if (extension.equalsIgnoreCase("docx") || extension.equalsIgnoreCase("doc")) {
			Document doc = new Document(fileInput);
			doc.save(fileOutput);
		}else {
			throw new Exception("Format Not Supporterd ,Supported Format Is .pptx , .ppt , .xlsx , .xls , .docx and .doc");
		}

		File file = new File(fileOutput);
		byte[] bytes = Files.readAllBytes(file.toPath());

		String b64 = Base64.getEncoder().encodeToString(bytes);
		Map<String, Object> result = new HashedMap<>();
		result.put("base64", b64);
		
		return result;
	}

}
