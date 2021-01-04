package com.demo.service.impl;


import org.docx4j.TextUtils;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.dml.CTTextAutonumberBullet;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.model.fields.merge.MailMerger;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.springframework.stereotype.Service;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.*;
import java.math.BigInteger;
import java.util.*;

@Service
public class DocxService {
    public ByteArrayOutputStream addHtmlToDocx(String content) throws Exception {
        ByteArrayOutputStream result = new ByteArrayOutputStream();
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
            String html = "<html><body>" + content + "</body></html>";
            XHTMLImporterImpl importer = new XHTMLImporterImpl(document);
            // look for all P elements in the specified object
            final List<P> paragraphs = new ArrayList<>();
            MainDocumentPart mainPart = document.getMainDocumentPart();
            new TraversalUtil(mainPart, new TraversalUtil.CallbackImpl() {
                @Override
                public List<Object> apply(Object o) {
                    if (o instanceof P) {
                        paragraphs.add((P) o);
                    }
                    return null;
                }
            });
            for (final P paragraph : paragraphs) {
                final StringWriter paragraphText = new StringWriter();
                try {
                    TextUtils.extractText(paragraph, paragraphText);
                } catch (Exception ex) {
                    System.out.println("Toang");
                }
                final String identifier = paragraphText.toString();
                if (identifier.equals("abc")) {
                    List<Object> listToModify;
                    if (paragraph.getParent() instanceof Tc) {
                        // paragraph located in table-cell
                        final Tc parent = (Tc) paragraph.getParent();
                        listToModify = parent.getContent();
                    } else {
                        // paragraph located in main document part
                        listToModify = mainPart.getContent();
                    }
                    if (listToModify != null) {
                        int index = listToModify.indexOf(paragraph);
                        // remove the paragraph
                        listToModify.remove(index);
                        // add html
                        listToModify.addAll(index, importer.convert(html, null));
                    }
                }
            }
            document.save(result);
        }
        return result;
    }

    public WordprocessingMLPackage replaceTextByBullets(WordprocessingMLPackage document, String text, String key) {
        final List<P> paragraphs = new ArrayList<>();
        MainDocumentPart mainPart = document.getMainDocumentPart();
        new TraversalUtil(mainPart, new TraversalUtil.CallbackImpl() {
            @Override
            public List<Object> apply(Object o) {
                if (o instanceof P) {
                    paragraphs.add((P) o);
                }
                return null;
            }
        });
        for (final P paragraph : paragraphs) {
            final StringWriter paragraphText = new StringWriter();
            try {
                TextUtils.extractText(paragraph, paragraphText);
            } catch (Exception ex) {
                System.out.println("Toang");
            }
            final String identifier = paragraphText.toString();
            if (identifier.contains(key)) {
                String prefix = "";
                if (identifier.lastIndexOf("MERGEFIELD " + key) != -1)
                    prefix = identifier.substring(0,identifier.lastIndexOf("MERGEFIELD " + key));
                List<Object> listToModify;
                if (paragraph.getParent() instanceof Tc) {
                    // paragraph located in table-cell
                    final Tc parent = (Tc) paragraph.getParent();
                    listToModify = parent.getContent();
                } else {
                    // paragraph located in main document part
                    listToModify = mainPart.getContent();
                }
                if (listToModify != null) {
                    int index = listToModify.indexOf(paragraph);
                    // remove the paragraph
                    listToModify.remove(index);
                    // add html
                    if (paragraph.getPPr().getNumPr() == null)
                    {
                        ObjectFactory factory = new ObjectFactory();
                        PPrBase.NumPr numPr = factory.createPPrBaseNumPr();
                        paragraph.getPPr().setNumPr(numPr);
                        PPrBase.NumPr.NumId numIdElement = factory.createPPrBaseNumPrNumId();
                        numPr.setNumId(numIdElement);
                        numIdElement.setVal(BigInteger.ZERO);
                        PPrBase.NumPr.Ilvl ilvlElement = factory.createPPrBaseNumPrIlvl();
                        numPr.setIlvl(ilvlElement);
                        ilvlElement.setVal(BigInteger.ZERO);
                    }
                    listToModify.addAll(index, createBullets(prefix, text, paragraph));
                }
            }
        }
        return document;
    }

    public List<P> createBullets(String prefix, String text, P template) {
        List<P> result = new ArrayList<>();
        String[] values = text.trim().split("\n");
        ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
        if (!prefix.isEmpty()){
            P p = factory.createP();
            org.docx4j.wml.Text t = factory.createText();
            t.setValue(prefix);
            org.docx4j.wml.R run = factory.createR();
            run.getContent().add(t);
            p.getContent().add(run);
            p.setPPr(template.getPPr());
            result.add(p);
        }
        for (int i = 0; i < values.length; i++) {
            P p = factory.createP();
            org.docx4j.wml.Text t = factory.createText();
            t.setValue(values[i]);
            org.docx4j.wml.R run = factory.createR();
            run.getContent().add(t);
            p.getContent().add(run);
            org.docx4j.wml.PPr ppr = factory.createPPr();
            p.setPPr(ppr);
            ppr.setPStyle(template.getPPr().getPStyle());
            ppr.setSpacing(template.getPPr().getSpacing());
            ppr.setInd(template.getPPr().getInd());
            PPrBase.NumPr numPr = factory.createPPrBaseNumPr();
            ppr.setNumPr(numPr);
            PPrBase.NumPr.NumId numIdElement = factory.createPPrBaseNumPrNumId();
            numPr.setNumId(numIdElement);
            numIdElement.setVal(BigInteger.valueOf(2));
            PPrBase.NumPr.Ilvl ilvlElement = factory.createPPrBaseNumPrIlvl();
            numPr.setIlvl(ilvlElement);
            if (prefix.isEmpty())
                ilvlElement.setVal(BigInteger.valueOf(0));
            else{
                PPrBase.NumPr level = template.getPPr().getNumPr();
                if (level != null)
                    ilvlElement.setVal(BigInteger.ONE.add(level.getIlvl().getVal()));
            }

            result.add(p);
        }
        return result;
    }

    public void fillMailMerge() {
        FileInputStream is = null;
        try {
            is = new FileInputStream("C:\\Users\\truon\\Desktop\\1.ADD-CK-GHDKX-1NG-KCC.docx");
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
            } catch (JAXBException | XPathBinderAssociationIsPartialException e) {
                e.printStackTrace();
            }
            if (list != null)
                for (Object o : list) {
                    o = XmlUtils.unwrap(o);
                    CTFFCheckBox checkBox = (CTFFCheckBox) o;
                    BooleanDefaultTrue booleanDefaultTrue = new BooleanDefaultTrue();
                    booleanDefaultTrue.setVal(true);
                    checkBox.setChecked(booleanDefaultTrue);
                    // get name of checkbox
//                    CTFFData data = (CTFFData) checkBox.getParent();
//                    CTFFName name = (CTFFName) data.getNameOrEnabledOrCalcOnExit().get(0).getValue();
                }
            Map<DataFieldName, String> values = new HashMap<>();
            Random random = new Random();
            String[] content = {"TỪ NAY DUYÊN KIẾP", "BỎ LẠI PHÍA SAU", "NGÀY VÀ BÓNG TỐI", "CHẲNG CÒN KHÁC NHAU"
                    , "CHẲNG CÓ NƠI NÀO YÊN BÌNH", "ĐƯỢC NHƯ EM BÊN ANH"};
            for (String mailMerge : mailMerges) {
                String value = content[random.nextInt(4)];
                if (mailMerge.equals("ĐT_HDTC") || mailMerge.equals("Lãi_suất_ghi_trên_KUNN") || mailMerge.equals("ĐT_HDTC")) {
                    document = replaceTextByBullets(document, "Điện thoại 1\nĐiện thoại 2\nĐiện thoại 3", mailMerge);
                } else
                    values.put(new DataFieldName(mailMerge), value);
            }

            MailMerger.setMERGEFIELDInOutput(MailMerger.OutputField.KEEP_MERGEFIELD);
            try {
                MailMerger.performMerge(document, values, false);

                document.save(new File("C:\\Users\\truon\\Desktop\\result.docx"));
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
                    int index = nextToken.lastIndexOf("«");
                    if (index != -1)
                        nextToken = nextToken.substring(0, index);
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