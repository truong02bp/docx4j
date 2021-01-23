package com.demo.service.impl;


import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import org.apache.commons.io.FileUtils;
import org.docx4j.*;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.model.fields.merge.MailMerger;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

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

    public List<String> getAllField(MultipartFile[] files) {
        List<String> fields = new ArrayList<>();
        for (MultipartFile file : files) {
            InputStream is = null;
            WordprocessingMLPackage document = null;
            try {
                is = new ByteArrayInputStream(file.getBytes());
                document = WordprocessingMLPackage.load(is);
            } catch (IOException | Docx4JException e) {
                e.printStackTrace();
            }
            if (document != null) {
                List<Object> checkboxes = null;
                List<Object> textFields = null;
                try {
                    checkboxes = document.getMainDocumentPart().getJAXBNodesViaXPath("//w:checkBox",
                            false);
                    textFields = document.getMainDocumentPart().getJAXBNodesViaXPath("//w:instrText ", true);
                } catch (JAXBException | XPathBinderAssociationIsPartialException e) {
                    e.printStackTrace();
                }
                if (checkboxes != null)
                    for (Object o : checkboxes) {
                        o = XmlUtils.unwrap(o);
                        CTFFCheckBox checkBox = (CTFFCheckBox) o;
                        CTFFData data = (CTFFData) checkBox.getParent();
                        CTFFName ctffName = (CTFFName) data.getNameOrEnabledOrCalcOnExit().get(0).getValue();
                        String name = ctffName.getVal();
                        System.out.println(ctffName.getVal());
                        System.out.println(ctffName.getParent().toString());
                        fields.add(name);
                    }
                if (textFields != null)
                    fields.addAll(getAllMergeFields(textFields));
            }
        }
        return fields;
    }

    public byte[] fillMailMerge(String type) {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        InputStream is = null;
        WordprocessingMLPackage document = null;
        try {
            is = new FileInputStream("C:\\Users\\truon\\Desktop\\1.ADD-CK-GHDKX-1NG-KCC.docx");
//            is = new ByteArrayInputStream(file.getBytes());
            document = WordprocessingMLPackage.load(is);
        } catch (IOException | Docx4JException e) {
            e.printStackTrace();
        }
        if (document != null) {
            List<String> mailMerges = new ArrayList<>();
            List<Object> checkboxes = null;
            List<Object> textFields = null;
            try {
                checkboxes = document.getMainDocumentPart().getJAXBNodesViaXPath("//w:checkBox",
                        false);
                textFields = document.getMainDocumentPart().getJAXBNodesViaXPath("//w:instrText ", true);
            } catch (JAXBException | XPathBinderAssociationIsPartialException e) {
                e.printStackTrace();
            }
            Map<DataFieldName, String> values = new HashMap<>();
            Random random = new Random();
            String[] content = {"TỪ NAY DUYÊN KIẾP", "BỎ LẠI PHÍA SAU", "NGÀY VÀ BÓNG TỐI", "CHẲNG CÒN KHÁC NHAU"
                    , "CHẲNG CÓ NƠI NÀO YÊN BÌNH", "ĐƯỢC NHƯ EM BÊN ANH"};
            // solve check box
            if (checkboxes != null)
                for (Object o : checkboxes) {
                    o = XmlUtils.unwrap(o);
                    CTFFCheckBox checkBox = (CTFFCheckBox) o;
                    CTFFData data = (CTFFData) checkBox.getParent();
                    CTFFName ctffName = (CTFFName) data.getNameOrEnabledOrCalcOnExit().get(0).getValue();
                    String name = ctffName.getVal();
                    BooleanDefaultTrue booleanDefaultTrue = new BooleanDefaultTrue();
                    booleanDefaultTrue.setVal(true);
                    checkBox.setChecked(booleanDefaultTrue);
                }
            // solve merge field
            if (textFields != null)
                mailMerges = getAllMergeFields(textFields);
            for (String mailMerge : mailMerges) {
                String value = content[random.nextInt(4)];
                if (mailMerge.equals("Lãi_suất_ghi_trên_KUNN") || mailMerge.equals("ĐT_HDTC")) {
                    document = replaceTextByBullets(document, "0964279710\nĐiện thoại 2 : 12345\nĐiện thoại 3 : 5678", mailMerge);
                } else
                    values.put(new DataFieldName(mailMerge), value);
            }
            MailMerger.setMERGEFIELDInOutput(MailMerger.OutputField.KEEP_MERGEFIELD);
            try {
                MailMerger.performMerge(document, values, false);
                Docx4J.save(document, byteArrayOutputStream);
            } catch (Docx4JException e) {
                e.printStackTrace();
            }
            if (type.equals("pdf"))
                return docxToPdf(byteArrayOutputStream.toByteArray());
        }
        return byteArrayOutputStream.toByteArray();
    }

    public byte[] docxToPdf(byte[] bytes) {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        try {
            InputStream docxInputStream = new ByteArrayInputStream(bytes);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(byteArrayOutputStream).as(DocumentType.PDF).execute();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return byteArrayOutputStream.toByteArray();
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
                    prefix = identifier.substring(0, identifier.lastIndexOf("MERGEFIELD " + key));
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
                    if (paragraph.getPPr().getNumPr() == null) {
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
        int index = 1;
        if (!prefix.isEmpty()) {
            P p = factory.createP();
            org.docx4j.wml.Text t = factory.createText();
            t.setValue(prefix + " " + values[0]);
            org.docx4j.wml.R run = factory.createR();
            run.getContent().add(t);
            p.getContent().add(run);
            p.setPPr(template.getPPr());
            result.add(p);
        } else
            index = 0;
        for (int i = index; i < values.length; i++) {
            P p = factory.createP();
            org.docx4j.wml.Text t = factory.createText();
            t.setValue(values[i]);
            org.docx4j.wml.R run = factory.createR();
            run.getContent().add(t);
            p.getContent().add(run);
            p.setPPr(template.getPPr());
            result.add(p);
        }
        return result;
    }

    public List<String> getAllMergeFields(List<Object> textFields) {
        List<String> mailMerges = new ArrayList<>();
        for (Object o : textFields) {
            o = XmlUtils.unwrap(o);
            Text text = (Text) o;
            String value = text.getValue();
            String mergeField = "MERGEFIELD ";
            if (value.contains(mergeField)) {
                String name = value.substring(mergeField.length());
                if (name.contains("\""))
                    name = name.substring(1, name.length() - 1);
                if (name.contains("\\"))
                    name = name.substring(0, name.lastIndexOf("\\")).trim();
                mailMerges.add(name);
            }
        }
        return mailMerges;
    }

    public void insertImageToDocx() throws Exception {
        InputStream is = null;
        WordprocessingMLPackage wordPackage = null;
        try {
            is = new FileInputStream("C:\\Users\\truon\\Desktop\\truong.docx");
//            is = new ByteArrayInputStream(file.getBytes());
            wordPackage = WordprocessingMLPackage.load(is);
        } catch (IOException | Docx4JException e) {
            e.printStackTrace();
        }
        assert wordPackage != null;
        File file = new File("C:\\Users\\truon\\Desktop\\qr.jpg");
        byte[] bytes = FileUtils.readFileToByteArray(file);
        String filenameHint = null;
        String altText = null;
        int id1 = 0;
        int id2 = 1;
        P p = newImage(wordPackage, bytes, filenameHint, altText, id1, id2);
        wordPackage.getMainDocumentPart().getContent().add(p);
        wordPackage.save(new File("C:\\Users\\truon\\Desktop\\b.docx"));
    }

    public P newImage(WordprocessingMLPackage wordMLPackage, byte[] bytes,
                      String filenameHint, String altText, int id1, int id2) throws Exception {
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
        Inline inline = imagePart.createImageInline(filenameHint, altText, id1, id2, false);
        ObjectFactory factory = new ObjectFactory();
        P p = factory.createP();
        R run = factory.createR();
        p.getContent().add(run);
        Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        return p;
    }

    public void docxToHtml() throws Exception {
        File file = new File("/home/truong02_bp/Downloads/1.ADD CK GH,DKX 1NG KCC.docx");
        FileInputStream fis = new FileInputStream(file);
        WordprocessingMLPackage docx = WordprocessingMLPackage.load(fis);
        String path = "/home/truong02_bp/Desktop/result";
        HTMLSettings htmlSettings = Docx4J.createHTMLSettings();
        htmlSettings.setWmlPackage(docx);
        htmlSettings.setImageIncludeUUID(true);
        htmlSettings.setImageDirPath(path + "_images");
        htmlSettings.setImageTargetUri(path.substring(path.lastIndexOf("/") + 1)
                + "_images");
        Docx4jProperties.setProperty("docx4j.Convert.Out.HTML.OutputMethodXML", true);
        Docx4jProperties.setProperty("docx", true);
        FileOutputStream os = new FileOutputStream(new File(path + ".html"));
        Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);
    }

    public void htmlToDocx() throws IOException, Docx4JException, JAXBException {
        String stringFromFile = FileUtils.readFileToString(new File("/home/truong02_bp/Desktop/result.html"), "UTF-8");
        WordprocessingMLPackage docx = WordprocessingMLPackage.createPackage();
        NumberingDefinitionsPart parts = new NumberingDefinitionsPart();
        docx.getMainDocumentPart().addTargetPart(parts);
        parts.unmarshalDefaultNumbering();
        RFonts arialRFonts = Context.getWmlObjectFactory().createRFonts();
        arialRFonts.setAscii("Arial");
        arialRFonts.setHAnsi("Arial");
        XHTMLImporterImpl.addFontMapping("Arial", arialRFonts);
        XHTMLImporterImpl importer = new XHTMLImporterImpl(docx);
        importer.setHyperlinkStyle("Hyperlink");
        docx.getMainDocumentPart().getContent().addAll(importer.convert(stringFromFile, null));
        docx.save(new File("/home/truong02_bp/Desktop/result.docx"));
    }

}