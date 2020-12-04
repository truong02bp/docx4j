package com.demo;

import com.demo.entities.File;
import com.demo.entities.Person;
import com.demo.repository.FileRepository;
import com.demo.service.impl.DocxService;
import com.spire.doc.Document;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.ModelAndView;

import java.io.ByteArrayOutputStream;
import java.util.HashMap;
import java.util.Map;

@Controller
public class HomeController
{
    @Autowired
    private DocxService docxService;

    @Autowired
    private FileRepository fileRepository;

    private static Map<Integer, String> map = new HashMap<>();
    private static int id=0;

    @GetMapping("/trang-chu")
    public ResponseEntity<byte[]> homePage() throws Exception {
        Document document = new Document();
        document.loadFromFile("/home/truong02_bp/Desktop/template.docx");
        String[] fields = document.getMailMerge().getMergeFieldNames();
        Map<String,String> values = new HashMap<>();
        Person person = new Person("Trường","Hà Đông","Siten","0964279710");
        String[] items = person.toString().split(";");
        for (String item : items) {
            String[] data = item.split("=");
            values.put(data[0],data[1]);
        }
        String[] value = new String[fields.length+1];
        for (int i=0;i<fields.length;i++)
            value[i] = values.get(fields[i]);
        document.getMailMerge().execute(fields,value);
        ByteArrayOutputStream os = new ByteArrayOutputStream();
//        document.saveToStream(os,FileFormat.Docx);
        byte[] bytes = os.toByteArray();
        String name = "result";
        String type = "docx";
        return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                .header(HttpHeaders.CONTENT_DISPOSITION,"attachment;filename="+name+"."+type)
                .body(bytes);
    }

    @GetMapping("/")
    public ResponseEntity<?> readDocx() throws Exception {
        docxService.addHtmlToDocx();
//        docxService.readMailMerge();
//        docxService.docxToHtml();
//        docxService.docxToHtmlWithSpire();
//        docxService.htmlToDocxWithSpire();
//        docxService.resolveMailMerge();
//        docxService.docxToHtmlWithSpire();
//        docxService.htmlToDocx();
        return ResponseEntity.ok("ok");
    }

    @GetMapping("/docx")
    public ResponseEntity<byte[]> ckeditor(@RequestParam("id") int id){
        String content = map.get(id);
        byte[] bytes = content.getBytes();
        String name = "result";
        String type = "docx";
        return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                .header(HttpHeaders.CONTENT_DISPOSITION,"attachment;filename="+name+"."+type)
                .body(bytes);
    }
    @GetMapping("/ckeditor")
    public ModelAndView ckeditor(){
        ModelAndView mav = new ModelAndView("ckeditor");
        return mav;
    }
    @PostMapping("/add")
    public ResponseEntity<Integer> toDocx(@RequestBody Content content){
        id++;
        map.put(id,content.getContent());
        return ResponseEntity.ok(id);
    }
}
