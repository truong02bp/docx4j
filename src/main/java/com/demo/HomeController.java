package com.demo;

import com.demo.entities.Person;
import com.demo.repository.FileRepository;
import com.demo.service.impl.DocxService;
//import com.spire.doc.Document;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import java.io.ByteArrayOutputStream;
import java.util.HashMap;
import java.util.Map;

@Controller
public class HomeController {
    @Autowired
    private DocxService docxService;

    @Autowired
    private FileRepository fileRepository;

    private static Map<Integer, String> map = new HashMap<>();
    private static int id = 0;

    @GetMapping("/docx")
    public ResponseEntity<byte[]> ckeditor(@RequestParam("id") int id) throws Exception {
        String content = map.get(id);
        byte[] bytes = docxService.addHtmlToDocx(content).toByteArray();
        String name = "result";
        String type = "docx";
        return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename=" + name + "." + type)
                .body(bytes);
    }

//    @GetMapping("/ckeditor")
//    public ResponseEntity<byte[]> ckeditor() {
//        HttpHeaders headers = new HttpHeaders();
//        headers.setContentType(MediaType.parseMediaType("application/pdf"));
//        String filename = "pdf1.pdf";
//        headers.add("content-disposition", "inline;filename=" + filename);
//        headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
//        ResponseEntity<byte[]> result = new ResponseEntity<byte[]>(bytes, headers, HttpStatus.OK);
//        return result;
//    }

    @GetMapping("/export-docx")
    public ResponseEntity<?> exportDocx() {
        byte[] bytes = docxService.fillMailMerge("docx");
        return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename=result.docx")
                .body(bytes);
    }

    @GetMapping("/export-pdf")
    public ResponseEntity<?> exportPdf() {
        byte[] bytes = docxService.fillMailMerge("pdf");
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("application/pdf"));
        String filename = "result.pdf";
        headers.add("content-disposition", "inline;filename=" + filename);
        headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
        ResponseEntity<byte[]> result = new ResponseEntity<byte[]>(bytes, headers, HttpStatus.OK);
        return result;
    }
}
