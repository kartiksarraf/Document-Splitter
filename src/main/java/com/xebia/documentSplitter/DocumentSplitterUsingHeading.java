package com.xebia.documentSplitter;


import com.appiancorp.suiteapi.common.Name;
import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.content.DocumentOutputStream;
import com.appiancorp.suiteapi.knowledge.Document;
import com.appiancorp.suiteapi.knowledge.DocumentDataType;
import com.appiancorp.suiteapi.knowledge.FolderDataType;
import com.appiancorp.suiteapi.process.exceptions.SmartServiceException;
import com.appiancorp.suiteapi.process.framework.*;
import com.appiancorp.suiteapi.process.palette.PaletteInfo;
import com.appiancorp.suiteapi.type.TypeService;
import com.xebia.documentSplitter.utils.DocumentUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static com.appiancorp.suiteapi.content.ContentConstants.VERSION_CURRENT;


@PaletteInfo(paletteCategory = "Appian Smart Services", palette = "Document Generation")
@Order({ "OriginalDocument", "NewDocumentPrefix", "SaveInFolder", "Success", "ErrorMessage", "NewDocumentsCreatedInString" })
public class DocumentSplitterUsingHeading extends AppianSmartService {

    private static final Logger logger = LoggerFactory.getLogger(DocumentSplitterUsingHeading.class);
    private final SmartServiceContext smartServiceCtx;
    private ContentService cs;
    private String newDocumentPrefix;
    private Long saveInFolder;
    private Long originalDocument;
    private String newDocumentsCreatedInString;
    private final List<Long> newDocumentsCreated = new ArrayList<>();
    private String errorMessage;
    private Boolean success = false;


    @Override
    public void run() throws SmartServiceException {

        try {

            // Open the document
            Long doc = cs.getVersionId(originalDocument, VERSION_CURRENT);
            try (InputStream fis = cs.getDocumentInputStream(doc);
                 XWPFDocument originalDoc = new XWPFDocument(OPCPackage.open(fis))) {

                CTStyles styles = originalDoc.getStyle();

                // Split and process document
                List<DocumentSection> splitDocuments = extractSections(originalDoc);

                logger.info("Splitting documents into {} sections", splitDocuments.size());

                // Write individual documents
                for (int i = 0; i < splitDocuments.size(); i++) {
                    DocumentSection section = splitDocuments.get(i);
                    XWPFDocument newDoc = new XWPFDocument();

                    copyDocumentSettings(originalDoc, newDoc);

                    for (IBodyElement element: section.elements) {
                        cloneElementWithFormatting(element, newDoc);
                    }

                    String filename = section.title != null && !section.title.isEmpty()
                            ? sanitizeFileName(section.title)
                            : String.format("section_%03d", i + 1);
                    String tempNewDocName = newDocumentPrefix + "_" + filename;;
                    Long tempDocument = DocumentUtil.createDocument(cs, tempNewDocName, tempNewDocName, "docx", saveInFolder);

                    newDocumentsCreated.add(tempDocument);

                    List<Document> tempDoc = Arrays.asList(cs.download(tempDocument, VERSION_CURRENT, false));
                    try (DocumentOutputStream out = tempDoc.get(0).getOutputStream()) {
                        newDoc.write(out);
                        logger.info("Created section file: {}", tempNewDocName);
                    }

                }
            }
            success = true;
            newDocumentsCreatedInString = Arrays.toString(this.newDocumentsCreated.toArray());
        } catch (Exception e) {
            errorMessage = e.getMessage();
        }

    }

    public DocumentSplitterUsingHeading(SmartServiceContext smartServiceCtx,
                                             ContentService cs, TypeService ts) {
        super();
        this.smartServiceCtx = smartServiceCtx;
        this.cs = cs;
    }

    private static class DocumentSection {
        String title;
        List<IBodyElement> elements = new ArrayList<>();
    }


    private static void copyDocumentSettings(XWPFDocument source, XWPFDocument target) {
        try {
            if (source.getStyles() != null) {
                XWPFStyles newStyles = target.createStyles();
                CTStyles styles = source.getStyle();
                if (styles != null) {
                    newStyles.setStyles(styles);
                }
            }

            if (source.getNumbering() != null) {
                target.createNumbering();
                for (XWPFAbstractNum abstractNum : source.getNumbering().getAbstractNums()) {
                    BigInteger abstractNumId = target.getNumbering().addAbstractNum(abstractNum);
                    target.getNumbering().addNum(abstractNumId, abstractNumId);
                }
            }
        } catch (Exception e) {
            logger.warn("Could not copy some document settings", e);
        }
    }

    private static List<DocumentSection> extractSections(XWPFDocument doc) {
        List<DocumentSection> sections = new ArrayList<>();
        DocumentSection currentSection = new DocumentSection();
        sections.add(currentSection);

        for (IBodyElement element : doc.getBodyElements()) {
            if (element instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                if (isHeading(paragraph)) {
                    currentSection = new DocumentSection();
                    currentSection.title = paragraph.getText();
                    sections.add(currentSection);
                }
            }
            currentSection.elements.add(element);
        }

        return sections;
    }

    private static boolean isHeading(XWPFParagraph paragraph) {
        String styleId = paragraph.getStyleID();
        if (styleId == null) return false;

        styleId = styleId.toLowerCase();
        String style = paragraph.getStyle();

        return styleId.startsWith("heading") ||
                (style != null && style.startsWith("Heading")) ||
                (styleId.matches("^h[1-9]$"));
    }

    private static void cloneElementWithFormatting(IBodyElement element, XWPFDocument target) {
        if (element instanceof XWPFParagraph) {
            XWPFParagraph sourcePara = (XWPFParagraph) element;
            XWPFParagraph targetPara = target.createParagraph();

            CTPPr ppr = sourcePara.getCTP().getPPr();
            if (ppr != null) {
                targetPara.getCTP().setPPr(ppr);
            }

            if (sourcePara.getStyle() != null) {
                targetPara.setStyle(sourcePara.getStyle());
            }

            if (sourcePara.getNumID() != null) {
                targetPara.setNumID(sourcePara.getNumID());
                if (sourcePara.getCTP().getPPr() != null &&
                        sourcePara.getCTP().getPPr().getNumPr() != null &&
                        sourcePara.getCTP().getPPr().getNumPr().getIlvl() != null) {
                    targetPara.setNumILvl(sourcePara.getCTP().getPPr().getNumPr().getIlvl().getVal());
                }
            }

            for (XWPFRun sourceRun : sourcePara.getRuns()) {
                XWPFRun targetRun = targetPara.createRun();
                copyRun(sourceRun, targetRun);
            }

        } else if (element instanceof XWPFTable) {
            XWPFTable sourceTable = (XWPFTable) element;
            XWPFTable targetTable = target.createTable();
            copyTable(sourceTable, targetTable);
        } else if (element instanceof XWPFSDT) {
            XWPFSDT sourceSDT = (XWPFSDT) element;
            if (isTOCField(sourceSDT)) {
                createTOCField(target);
            } else {
                copySDTContent(sourceSDT, target);
            }
        }
    }

    private static boolean isTOCField(XWPFSDT sdt) {
        try {
            String sdtTag = sdt.getTag();
            return sdtTag != null && sdtTag.toLowerCase().contains("toc");
        } catch (Exception e) {
            logger.warn("Could not check if SDT is TOC field", e);
        }
        return false;
    }

    private static void createTOCField(XWPFDocument target) {
        try {
            XWPFParagraph tocHeader = target.createParagraph();
            XWPFRun tocRun = tocHeader.createRun();
            tocRun.setText("Table of Contents");
            tocRun.setBold(true);
            tocRun.setFontSize(14);

            XWPFParagraph tocPara = target.createParagraph();
            CTP ctp = tocPara.getCTP();

            CTR run1 = ctp.addNewR();
            run1.addNewFldChar().setFldCharType(STFldCharType.BEGIN);

            CTR run2 = ctp.addNewR();
            CTText text = run2.addNewInstrText();
            text.setStringValue(" TOC \\o \"1-3\" \\h \\z \\u ");

            CTR run3 = ctp.addNewR();
            run3.addNewFldChar().setFldCharType(STFldCharType.END);

            XWPFParagraph spacer = target.createParagraph();
            spacer.setSpacingAfter(400);
        } catch (Exception e) {
            logger.warn("Could not create TOC field", e);
        }
    }

    private static void copySDTContent(XWPFSDT source, XWPFDocument target) {
        try {
            ISDTContent content = source.getContent();
            if (content != null) {
                XWPFParagraph newPara = target.createParagraph();
                if (content instanceof XWPFParagraph) {
                    XWPFParagraph sourcePara = (XWPFParagraph) content;
                    for (XWPFRun run : sourcePara.getRuns()) {
                        XWPFRun newRun = newPara.createRun();
                        copyRun(run, newRun);
                    }
                } else {
                    // If it's not a paragraph, just copy the text content
                    XWPFRun newRun = newPara.createRun();
                    newRun.setText(content.getText());
                }
            }
        } catch (Exception e) {
            logger.warn("Could not copy SDT content", e);
        }
    }

    private static void copyRun(XWPFRun source, XWPFRun target) {
        target.setText(source.text());
        target.setBold(source.isBold());
        target.setItalic(source.isItalic());
        target.setUnderline(source.getUnderline());
        target.setColor(source.getColor());
        target.setFontFamily(source.getFontFamily());
        if (source.getFontSize() != -1) {
            target.setFontSize(source.getFontSize());
        }

        CTRPr rpr = source.getCTR().getRPr();
        if (rpr != null) {
            target.getCTR().setRPr(rpr);
        }

        try {
            for (XWPFPicture pic : source.getEmbeddedPictures()) {
                try (ByteArrayInputStream bis = new ByteArrayInputStream(pic.getPictureData().getData())) {
                    target.addPicture(
                            bis,
                            pic.getPictureData().getPictureType(),
                            pic.getDescription(),
                            (int)(pic.getWidth() * 9525),
                            (int)(pic.getDepth() * 9525)
                    );
                }
            }
        } catch (Exception e) {
            logger.warn("Could not copy picture", e);
        }
    }

    private static void copyTable(XWPFTable source, XWPFTable target) {
        try {
            if (source.getCTTbl() != null) {
                if (source.getCTTbl().getTblPr() != null) {
                    target.getCTTbl().setTblPr(source.getCTTbl().getTblPr());
                }
                if (source.getCTTbl().getTblGrid() != null) {
                    target.getCTTbl().setTblGrid(source.getCTTbl().getTblGrid());
                }
            }

            for (XWPFTableRow sourceRow : source.getRows()) {
                XWPFTableRow targetRow = target.createRow();
                targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());

                List<XWPFTableCell> sourceCells = sourceRow.getTableCells();
                List<XWPFTableCell> targetCells = targetRow.getTableCells();

                for (int i = 0; i < sourceCells.size(); i++) {
                    XWPFTableCell sourceCell = sourceCells.get(i);
                    XWPFTableCell targetCell;

                    if (i < targetCells.size()) {
                        targetCell = targetCells.get(i);
                    } else {
                        targetCell = targetRow.createCell();
                    }

                    targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());

                    for (XWPFParagraph sourcePara : sourceCell.getParagraphs()) {
                        XWPFParagraph targetPara = targetCell.addParagraph();
                        if (sourcePara.getCTP().getPPr() != null) {
                            targetPara.getCTP().setPPr(sourcePara.getCTP().getPPr());
                        }
                        if (sourcePara.getStyle() != null) {
                            targetPara.setStyle(sourcePara.getStyle());
                        }
                        for (XWPFRun sourceRun : sourcePara.getRuns()) {
                            XWPFRun targetRun = targetPara.createRun();
                            copyRun(sourceRun, targetRun);
                        }
                    }
                }
            }
        } catch (Exception e) {
            logger.warn("Could not copy table completely", e);
        }
    }

    private static String sanitizeFileName(String input) {
        return input.replaceAll("[\\\\/:*?\"<>|]", "_")
                .replaceAll("\\s+", "_")
                .trim();
    }

    @Input(required = Required.OPTIONAL)
    @Name("NewDocumentPrefix")
    public void setNewDocumentPrefix(String val) {
        this.newDocumentPrefix = val;
    }


    @Input(required = Required.OPTIONAL)
    @Name("SaveInFolder")
    @FolderDataType
    public void setSaveInFolder(Long val) {
        this.saveInFolder = val;
    }

    @Input(required = Required.ALWAYS)
    @Name("OriginalDocument")
    @DocumentDataType
    public void setOriginalDocument(Long val) {
        this.originalDocument = val;
    }

    @Name("NewDocumentsCreatedInString")
    public String getNewDocumentsCreatedInString() {
        return newDocumentsCreatedInString;
    }

    @Name("ErrorMessage")
    public String getErrorMessage() {
        return errorMessage;
    }

    @Name("Success")
    public Boolean getSuccess() {
        return success;
    }

}
