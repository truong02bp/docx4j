package com.demo;

import com.demo.service.impl.DocxService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;

//import com.spire.doc.Document;

@Controller
public class HomeController {
    @Autowired
    private DocxService docxService;


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
    @PostMapping("/upload-files")
    public ResponseEntity<List<String>> getField(@RequestBody MultipartFile[] files) {
        List<String> list = docxService.getAllField(files);
        return ResponseEntity.ok(list);
    }
    @PostMapping("/doc-to-docx")
    public ResponseEntity<?> toDocx(@RequestBody MultipartFile file) {
        byte[] bytes = new byte[0];
        try {
            bytes = docxService.docToDocx(file.getBytes());
        } catch (IOException e) {
            e.printStackTrace();
        }
        return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename=result.docx")
                .body(bytes);
    }
    @PostMapping("/export-zip")
    public ResponseEntity<?> exportZip(@RequestBody MultipartFile[] files) {
        Random random = new Random();
        String[] content = {"TỪ NAY DUYÊN KIẾP\na", "BỎ LẠI PHÍA SAU\na", "NGÀY VÀ BÓNG TỐI\na", "CHẲNG CÒN KHÁC NHAU\na"
                , "CHẲNG CÓ NƠI NÀO YÊN BÌNH\na", "ĐƯỢC NHƯ EM BÊN ANH\na"};
        Map<String,String> map = new HashMap<>();
        docxService.getAllField(files).forEach(s -> {
            map.put(s,content[random.nextInt(4)]);
        });
        String fileName = "result1.zip";
        byte[] bytes = docxService.filesToZip(files,map);
        return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType("application/zip"))
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename="+fileName)
                .body(bytes);
    }

//    @GetMapping("/export-docx")
//    public ResponseEntity<?> exportDocx() {
//        byte[] bytes = docxService.fillMailMerge();
//        return ResponseEntity.ok()
//                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
//                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename=result.docx")
//                .body(bytes);
//    }

    @GetMapping("/export-pdf")
    public ResponseEntity<?> exportPdf(@RequestBody MultipartFile file) {
        byte[] bytes = null;
        try {
            bytes = docxService.docxToPdf(file.getBytes());
        } catch (IOException e) {
            e.printStackTrace();
        }
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("application/pdf"));
        String fileName = "result.pdf";
        headers.add("content-disposition", "inline;filename=" + fileName);
        headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
        return new ResponseEntity<>(bytes, headers, HttpStatus.OK);
    }

    @GetMapping("/insert-image")
    public ResponseEntity<?> insertImage() {
        try {
            docxService.insertImageToDocx();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return ResponseEntity.ok("Success");
    }
}
