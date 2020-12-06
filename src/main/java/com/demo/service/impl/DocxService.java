package com.demo.service.impl;


import com.spire.doc.CssStyleSheetType;
import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.collections.SectionCollection;
import com.spire.doc.documents.HorizontalAlignment;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.*;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.jaxb.Context;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.model.fields.merge.MailMerger;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.JaxbXmlPartAltChunkHost;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.AltChunkType;
import org.docx4j.openpackaging.parts.WordprocessingML.AlternativeFormatInputPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
import org.springframework.stereotype.Service;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.*;
import java.util.*;

@Service
public class DocxService {

//    public ByteArrayOutputStream exportDocx(String content){
////        String outPath = "/home/truong02_bp/Desktop/value.docx";
//        ByteArrayInputStream is = new ByteArrayInputStream(content.getBytes());
//        Document document = new Document();
//        document.loadFromStream(is,FileFormat.Html);
//        ByteArrayOutputStream os = new ByteArrayOutputStream();
//        document.saveToStream(os,FileFormat.Docx);
//        return os;
//    }
//
//    public void docxToHtmlWithSpire() {
//        String outPath = "/home/truong02_bp/Desktop/result";
//        String inputPath = "/home/truong02_bp/Downloads/1.ADD CK GH,DKX 1NG KCC.docx";
//        Document document = new Document();
//        document.loadFromFile(inputPath);
//        document.getHtmlExportOptions().setImageEmbedded(true);
//        document.getSections().get(0).getPageSetup().getMargins().setLeft(72f);
//        document.getHtmlExportOptions().setCssStyleSheetType(CssStyleSheetType.Internal);
//        document.saveToFile(outPath + ".html", FileFormat.Html);
//    }
//
//    public void htmlToDocxWithSpire() {
//        String outPath = "/home/truong02_bp/Desktop/value.docx";
//        String inputPath = "/home/truong02_bp/Desktop/result.html";
//        Document document = new Document();
//        document.loadFromFile(inputPath);
//        document.saveToFile(outPath, FileFormat.Docx);
//        System.out.println(document.getSections().get(0).getPageSetup().getMargins().getLeft() + " " + document.getSections().get(0).getPageSetup().getMargins().getRight() + " " + document.getSections().get(0).getPageSetup().getMargins().getTop() + " " + document.getSections().get(0).getPageSetup().getMargins().getBottom());
//    }
//
//    public void resolveMailMerge() throws Exception {
//        String outPath = "/home/truong02_bp/Desktop/solve";
//        String inputPath = "/home/truong02_bp/Downloads/test.docx";
//        Document document = new Document();
//        document.loadFromFile(inputPath);
//        solveFormField(document , createTemplate(document));
//        document.saveToFile(outPath + ".docx", FileFormat.Docx);
//    }
//
//    public Map<String, String> createTemplate(Document document) {
//        String[] values = {"0964279710", "Truong"};
//        String[] template = {"TỪ NAY DUYÊN KIẾP", "BỎ LẠI PHÍA SAU", "NGÀY VÀ BÓNG TỐI", "CHẲNG CÒN KHÁC NHAU"
//                , "CHẲNG CÓ NƠI NÀO YÊN BÌNH", "ĐƯỢC NHƯ EM BÊN ANH"};
//        Random random = new Random();
//        Map<String, String> map = new HashMap<>();
//        for (Object o : document.getFields()) {
//            if (o instanceof CheckBoxFormField) {
//                CheckBoxFormField checkbox = (CheckBoxFormField) o;
//                map.put(checkbox.getName(), String.valueOf(random.nextBoolean()));
//            }
//            else
//            if (o instanceof TextFormField)
//            {
//                TextFormField field = (TextFormField) o;
//                map.put(field.getName(),values[random.nextInt(2)]);
//            }
//            else
//            if (o instanceof MergeField)
//            {
//                MergeField mergeField = (MergeField) o;
//                map.put(mergeField.getFieldName(),template[random.nextInt(6)]);
//            }
//        }
//        return map;
//    }
//
//    public void solveFormField(Document document, Map<String, String> map) {
//        for (Object o : document.getFields()) {
//            if (o instanceof CheckBoxFormField) {
//                CheckBoxFormField checkbox = (CheckBoxFormField) o;
//                checkbox.setChecked(Boolean.parseBoolean(map.get(checkbox.getName()).toLowerCase()));
//            } else if (o instanceof TextFormField) {
//                TextFormField field = (TextFormField) o;
//                field.setText(map.get(field.getName()));
//            } else if (o instanceof MergeField) {
//                MergeField mergeField = (MergeField) o;
//                mergeField.setText(map.get(mergeField.getFieldName()));
//            }
//        }
//    }

    public void addHtmlToDocx(String content) throws Exception {
        FileInputStream is = null;
        try {
            is = new FileInputStream("/home/truong02_bp/Desktop/1.ADD-CK-GHDKX-1NG-KCC.docx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        WordprocessingMLPackage document = null;
        try {
            document = WordprocessingMLPackage.load(is);
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
        if (document != null) {
            VariablePrepare.prepare(document);
            MainDocumentPart documentPart = document.getMainDocumentPart();
            String html = "<html><head><title>Import me</title></head><body><p style='color:#ff0000;'>Hello World!</p></body></html>";
            AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(new PartName("/hw.html"));
            afiPart.setBinaryData(html.toString().getBytes());
            afiPart.setContentType(new ContentType("text/html"));
            Relationship altChunkRel = documentPart.addTargetPart(afiPart);
            CTAltChunk ac = Context.getWmlObjectFactory().createCTAltChunk();
            ac.setId(altChunkRel.getId());
            HashMap<String, String> mappings = new HashMap<String, String>();
            mappings.put("abc", ac.toString());
            documentPart.variableReplace(mappings);
            document.save(new File("/home/truong02_bp/Desktop/result.docx"));
//            XHTMLImporterImpl importer = new XHTMLImporterImpl(document);
//            pkg.getMainDocumentPart().getContent().addAll(importer.convert(xhtml, null));
        }
    }

    public void readMailMerge() {
        FileInputStream is = null;
        try {
            is = new FileInputStream("/home/truong02_bp/Downloads/1.ADD CK GH,DKX 1NG KCC.docx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        WordprocessingMLPackage document = null;
        try {
            document = WordprocessingMLPackage.load(is);
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
        if (document != null) {
            List<String> mailMerges = getAllMailMerge(document.getMainDocumentPart().getContent());
            List<Object> list = null;
            try {
                list = document.getMainDocumentPart().getJAXBNodesViaXPath("//w:checkBox",
                        true);
            } catch (JAXBException e) {
                e.printStackTrace();
            } catch (XPathBinderAssociationIsPartialException e) {
                e.printStackTrace();
            }
            for (Object o : list) {
                o = XmlUtils.unwrap(o);
                CTFFCheckBox checkBox = (CTFFCheckBox) o;
                BooleanDefaultTrue booleanDefaultTrue = new BooleanDefaultTrue();
                booleanDefaultTrue.setVal(true);
                checkBox.setChecked(booleanDefaultTrue);
                CTFFData data = (CTFFData) checkBox.getParent();
                CTFFName name = (CTFFName) data.getNameOrEnabledOrCalcOnExit().get(0).getValue();
                System.out.println(name.getVal());
            }
            String[] content = {"Trường", "Chường", "Hello", "Goodbye"};
            Map<DataFieldName, String> values = new HashMap<>();
            Random random = new Random();
            for (String mailMerge : mailMerges) {
                String value = content[random.nextInt(4)];
                values.put(new DataFieldName(mailMerge), value);
            }
            MailMerger.setMERGEFIELDInOutput(MailMerger.OutputField.REMOVED);
            try {
                MailMerger.performMerge(document, values, false);
            } catch (Docx4JException e) {
                e.printStackTrace();
            }

            try {
                document.save(new File("/home/truong02_bp/Desktop/result.docx"));
            } catch (Docx4JException e) {
                e.printStackTrace();
            }
        }
    }

    public List<String> getAllMailMerge(List<Object> objects) {
        List<String> fields = new ArrayList<>();
        for (Object o : objects) {
            if (o instanceof JAXBElement) {
                fields.addAll(getMailMergeFromTable(o));
            } else
                fields.addAll(getMailMerge(o.toString()));
        }
        return fields;
    }

    public List<String> getMailMergeFromTable(Object o) {
        List<String> fields = new ArrayList<>();
        o = ((JAXBElement<?>) o).getValue();
        List<Object> texts = null;
        if (o.getClass().equals(Tbl.class)) {
            Tbl table = (Tbl) o;
            texts = getAllElementFromObject(table, Text.class);
        } else if (o.getClass().equals(CTBookmark.class)) {
            CTBookmark ctBookmark = (CTBookmark) o;
            texts = getAllElementFromObject(ctBookmark.getParent(), Text.class);
        }
        if (texts != null) {
            StringBuilder stringBuilder = new StringBuilder("");
            for (Object t :
                    texts) {
                Text tx = (Text) t;
                stringBuilder.append(tx.getValue());
            }
            fields.addAll(getMailMerge(stringBuilder.toString()));
        }
        return fields;
    }

    public List<String> getMailMerge(String content) {
        List<String> fields = new ArrayList<>();
        if ((content.contains("MERGEFIELD") && content.contains("«") && content.contains("»"))) {
            StringTokenizer stringTokenizer = new StringTokenizer(content, " ");
            while (stringTokenizer.hasMoreTokens()) {
                String value = stringTokenizer.nextToken();
                if (value.equals("MERGEFIELD")) {
                    String nextToken = stringTokenizer.nextToken();
                    if (nextToken.contains("\""))
                        nextToken = nextToken.substring(1, nextToken.length() - 1);
                    fields.add(nextToken);
                }
            }
        }
        return fields;
    }

    private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch))
            result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }


//    public void docxToHtml() throws Exception {
//        File file = new File("/home/truong02_bp/Downloads/1.ADD CK GH,DKX 1NG KCC.docx");
//        FileInputStream fis = new FileInputStream(file);
//        WordprocessingMLPackage docx = WordprocessingMLPackage.load(fis);
//        String path = "/home/truong02_bp/Desktop/result";
//        HTMLSettings htmlSettings = Docx4J.createHTMLSettings();
//        htmlSettings.setWmlPackage(docx);
//        htmlSettings.setImageIncludeUUID(true);
//        htmlSettings.setImageDirPath(path+"_images");
//        htmlSettings.setImageTargetUri(path.substring(path.lastIndexOf("/")+1)
//                + "_images");
//        Docx4jProperties.setProperty("docx4j.Convert.Out.HTML.OutputMethodXML",true);
//        Docx4jProperties.setProperty("docx", true);
//        FileOutputStream os = new FileOutputStream(new File(path+".html"));
//        Docx4J.toHTML(htmlSettings,os,Docx4J.FLAG_EXPORT_PREFER_XSL);
//    }

//    public void htmlToDocx() throws IOException, Docx4JException, JAXBException {
//        String stringFromFile = FileUtils.readFileToString(new File("/home/truong02_bp/Desktop/result.html"),"UTF-8");
//        WordprocessingMLPackage docx = WordprocessingMLPackage.createPackage();
//        NumberingDefinitionsPart parts = new NumberingDefinitionsPart();
//        docx.getMainDocumentPart().addTargetPart(parts);
//        parts.unmarshalDefaultNumbering();
//        RFonts arialRFonts = Context.getWmlObjectFactory().createRFonts();
//        arialRFonts.setAscii("Arial");
//        arialRFonts.setHAnsi("Arial");
//        XHTMLImporterImpl.addFontMapping("Arial",arialRFonts);
//        XHTMLImporterImpl importer = new XHTMLImporterImpl(docx);
//        importer.setHyperlinkStyle("Hyperlink");
//        docx.getMainDocumentPart().getContent().addAll(importer.convert(stringFromFile,null));
//        docx.save(new File("/home/truong02_bp/Desktop/result.docx"));
//    }
}
