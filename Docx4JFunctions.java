/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.dikmen;

//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
//import java.io.IOException;
//import java.io.InputStream;
//import java.math.BigInteger;
//import java.util.ArrayList;
//import java.util.List;
//import javax.xml.bind.JAXBElement;
//import org.docx4j.dml.wordprocessingDrawing.Inline;
//import org.docx4j.jaxb.Context;
//import org.docx4j.model.structure.SectionWrapper;
//import org.docx4j.openpackaging.exceptions.Docx4JException;
//import org.docx4j.openpackaging.exceptions.InvalidFormatException;
//import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
//import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
//import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
//import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
//import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
//import org.docx4j.relationships.Relationship;
//import org.docx4j.w14.CTSdtCheckbox;
//import org.docx4j.w14.CTSdtCheckboxSymbol;
//import org.docx4j.wml.BooleanDefaultTrue;
//import org.docx4j.wml.Br;
//import org.docx4j.wml.CTBorder;
//import org.docx4j.wml.CTHeight;
//import org.docx4j.wml.CTPageNumber;
//import org.docx4j.wml.CTSimpleField;
//import org.docx4j.wml.CTVerticalJc;
//import org.docx4j.wml.ContentAccessor;
//import org.docx4j.wml.Drawing;
//import org.docx4j.wml.FooterReference;
//import org.docx4j.wml.Ftr;
//import org.docx4j.wml.HdrFtrRef;
//import org.docx4j.wml.HpsMeasure;
//import org.docx4j.wml.Id;
//import org.docx4j.wml.Jc;
//import org.docx4j.wml.JcEnumeration;
//import org.docx4j.wml.ObjectFactory;
//import org.docx4j.wml.P;
//import org.docx4j.wml.PPr;
//import org.docx4j.wml.PPrBase;
//import org.docx4j.wml.R;
//import org.docx4j.wml.RFonts;
//import org.docx4j.wml.STBorder;
//import org.docx4j.wml.STBrType;
//import org.docx4j.wml.STHeightRule;
//import org.docx4j.wml.STVerticalJc;
//import org.docx4j.wml.SdtBlock;
//import org.docx4j.wml.SdtContentBlock;
//import org.docx4j.wml.SdtPr;
//import org.docx4j.wml.SectPr;
//import org.docx4j.wml.Style;
//import org.docx4j.wml.Styles;
//import org.docx4j.wml.Tbl;
//import org.docx4j.wml.TblBorders;
//import org.docx4j.wml.TblPr;
//import org.docx4j.wml.TblWidth;
//import org.docx4j.wml.Tc;
//import org.docx4j.wml.TcPr;
//import org.docx4j.wml.TcPrInner;
//import org.docx4j.wml.Text;
//import org.docx4j.wml.Tr;
//import org.docx4j.wml.TrPr;
//import org.docx4j.wml.U;
//import org.docx4j.wml.UnderlineEnumeration;
//
///**
// *
// * @author mutlu.dikmen
// */
public class Docx4JFunctions {

//    private static org.docx4j.w14.ObjectFactory w14ObjectFactory = new org.docx4j.w14.ObjectFactory();
//    public static final int writableWidth = 9000; //wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips(); //9027
//    public static final int firstCellWidth = 2500;
//    public static final int secondCellWidth = 6500;
//    public static final int checkBoxWidth = 400;
//    public static final int checkBoxExplanationWidth = 8600;
//    public static final String fontPageHeader = "22";
//    public static final String fontHeader = "20";
//    public static final String font = "20";
//    public static final String fontSmall = "16";
//
//    public static enum tableBorder {
//        ALL, OUTER, NONE
//    };
//
//    public static void addFooterToDocument(WordprocessingMLPackage wordMLPackage, ObjectFactory factory, PPr pharagraphProperties) throws InvalidFormatException {
//        FooterPart footerPart = new FooterPart();
//        CTSimpleField ctSimple = factory.createCTSimpleField();
//        ctSimple.setInstr(" PAGE \\* MERGEFORMAT ");
//
//        org.docx4j.wml.RPr RPr = factory.createRPr();
//        RPr.setNoProof(new BooleanDefaultTrue());
//
//        R footerRun = factory.createR();
//        footerRun.setRPr(RPr);
//
//        ctSimple.getContent().add(footerRun);
//        JAXBElement<CTSimpleField> fldSimple = factory.createPFldSimple(ctSimple);
//
//        P para = factory.createP();
//        para.getContent().add(fldSimple);
//        para.setPPr(pharagraphProperties);
//
//        Ftr ftr = factory.createFtr();
//        ftr.getContent().add(para);
//
//        footerPart.setJaxbElement(ftr);
//
//        Relationship addTargetPart = wordMLPackage.getMainDocumentPart().addTargetPart(footerPart);
//
//        FooterReference ft = factory.createFooterReference();
//        ft.setType(HdrFtrRef.DEFAULT);
//        ft.setId(addTargetPart.getId());
//
//        SectPr sectPr = factory.createSectPr();
//        sectPr.getEGHdrFtrReferences().add(ft);
//
//        ft.setParent(sectPr);
//
//        wordMLPackage.getMainDocumentPart().addObject(sectPr);
//    }
//
//    public static void addTableWithOneRowAndOneCellToDocument(ObjectFactory factory, MainDocumentPart document, Tc cell, tableBorder border, JcEnumeration justification) {
//        Tbl table = factory.createTbl();
//        Tr row = factory.createTr();
//        row.getContent().add(cell);
//        table.getContent().add(row);
//        setTblAllJcAlign(table, justification);
//        addBordersToTable(table, border);
//        document.getContent().add(table);
//    }
//
//    public static void addRowToTableWithCells(ObjectFactory factory, Tbl table, List<Tc> tcList) {
//        Tr row = factory.createTr();
//        if (tcList != null) {
//            for (Tc cell : tcList) {
//                row.getContent().add(cell);
//            }
//        }
//        table.getContent().add(row);
//    }
//
//    public static Tr createRowWithCells(ObjectFactory factory, List<Tc> tcList) {
//        Tr row = factory.createTr();
//        if (tcList != null) {
//            for (Tc cell : tcList) {
//                row.getContent().add(cell);
//            }
//        }
//        return row;
//    }
//
//    public static Tr createRowWithOneCell(ObjectFactory factory, Tc cell) {
//        Tr row = factory.createTr();
//
//        row.getContent().add(cell);
//
//        return row;
//    }
//
//    public static void addTableToDocument(ObjectFactory factory, MainDocumentPart document, List<Tr> rowList, tableBorder border, JcEnumeration justification) {
//        Tbl table = factory.createTbl();
//        if (rowList != null) {
//            for (Tr row : rowList) {
//                table.getContent().add(row);
//            }
//        }
//        setTblAllJcAlign(table, justification);
//        addBordersToTable(table, border);
//        document.getContent().add(table);
//    }
//
//    public static void addPharagraphWithContent(WordprocessingMLPackage wordMLPackage, ObjectFactory factory, PPr paragraphProperties, org.docx4j.wml.RPr runProperties, String content) {
//        if (content == null) {
//            content = " ";
//        }
//        P pharagraph = factory.createP();
//        pharagraph.setPPr(paragraphProperties);
//        R run = factory.createR();
//        run.setRPr(runProperties);
//        addContentToRun(run, content);
//        pharagraph.getContent().add(run);
//        wordMLPackage.getMainDocumentPart().addObject(pharagraph);
//    }
//
//    public static void addTableCell(WordprocessingMLPackage wordMLPackage, ObjectFactory factory, Tr tableRow, String content) {
//if (content == null) {
//            content = " ";
//        }
//        Tc tableCell = factory.createTc();
//        tableCell.getContent().add(
//                wordMLPackage.getMainDocumentPart().createParagraphOfText(content));
//        tableRow.getContent().add(tableCell);
//    }
//
//    public static Tbl createTableWithContent(WordprocessingMLPackage wordMLPackage, ObjectFactory factory, String content) {
//if (content == null) {
//            content = " ";
//        }
//        Tbl table = factory.createTbl();
//        Tr tableRow = factory.createTr();
//        addTableCell(wordMLPackage, factory, tableRow, content);
//
//        table.getContent().add(tableRow);
//
//        return table;
//    }
//
//    public static void addBordersToTable(Tbl table, tableBorder tb) {
//        CTBorder border = new CTBorder();
//        TblBorders borders = new TblBorders();
//        if (table.getTblPr() == null) {
//            TblPr tblPr = new TblPr();
//            table.setTblPr(tblPr);
//        }
//        border.setColor("auto");
//        border.setSz(new BigInteger("4"));
//        border.setSpace(new BigInteger("0"));
//        border.setVal(STBorder.SINGLE);
//        switch (tb) {
//            case NONE:
//                break;
//            case OUTER:
//                borders.setBottom(border);
//                borders.setLeft(border);
//                borders.setRight(border);
//                borders.setTop(border);
//                table.getTblPr().setTblBorders(borders);
//                break;
//            case ALL:
//                borders.setBottom(border);
//                borders.setLeft(border);
//                borders.setRight(border);
//                borders.setTop(border);
//                borders.setInsideH(border);
//                borders.setInsideV(border);
//                table.getTblPr().setTblBorders(borders);
//                break;
//            default:
//                break;
//        }
//    }
//
//    public static void setFontSize(org.docx4j.wml.RPr runProperties, String fontSize) {
//        HpsMeasure size = new HpsMeasure();
//        size.setVal(new BigInteger(fontSize));
//        runProperties.setSz(size);
//        runProperties.setSzCs(size);
//    }
//
//    public static void addBoldStyle(org.docx4j.wml.RPr runProperties) {
//        BooleanDefaultTrue b = new BooleanDefaultTrue();
//        b.setVal(true);
//        runProperties.setB(b);
//    }
//
//    public static void setCellWidth(Tc tableCell, int width) {
//        TcPr tableCellProperties = new TcPr();
//        TblWidth tableWidth = new TblWidth();
//        tableWidth.setW(BigInteger.valueOf(width));
//        tableCellProperties.setTcW(tableWidth);
//        tableCell.setTcPr(tableCellProperties);
//    }
//
//    public static void setTableRowHeight(ObjectFactory factory, Tr tableRow, long height) {
//        TrPr trPr = new TrPr();
//        CTHeight ctHeight = new CTHeight();
//        ctHeight.setHRule(STHeightRule.EXACT);
//        ctHeight.setVal(BigInteger.valueOf(height));
//        JAXBElement<CTHeight> jaxbElement = factory.createCTTrPrBaseTrHeight(ctHeight);
//        trPr.getCnfStyleOrDivIdOrGridBefore().add(jaxbElement);
//        tableRow.setTrPr(trPr);
//    }
//
//    public static Tc createCheckBox(ObjectFactory factory, String explanation, int width) {
//        PPr paragraphProperties = factory.createPPr();
//        Jc justification = factory.createJc();
//        justification.setVal(JcEnumeration.CENTER);
//        paragraphProperties.setJc(justification);
//        Tc tc = factory.createTc();
//        P pharagraph = factory.createP();
//        pharagraph.setPPr(paragraphProperties);
//        R run = factory.createR();
//        R run2 = factory.createR();
//        SdtBlock stBlock = createCheckboxSdtBlock(factory, false);
//        run.getContent().add(stBlock);
//        Text t = factory.createText();
//        t.setSpace("preserve");
//        t.setValue(" " + explanation);
////        R.Tab rtab = factory.createRTab();
////        JAXBElement<org.docx4j.wml.R.Tab> rtabWrapped = factory.createRTab(rtab);
////        run2.getContent().add(rtabWrapped);
//        run2.getContent().add(t);
//        pharagraph.getContent().add(run);
//        pharagraph.getContent().add(run2);
//        CTVerticalJc vjc = new CTVerticalJc();
//        vjc.setVal(STVerticalJc.CENTER);
//        TcPr tcPr = factory.createTcPr();
//        tcPr.setVAlign(vjc);
//        TblWidth tableWidth = new TblWidth();
//        tableWidth.setW(BigInteger.valueOf(width));
//        tcPr.setTcW(tableWidth);
//        tc.setTcPr(tcPr);
//
//        tc.getContent().add(pharagraph);
//
//        return tc;
//    }
//
//    public static SdtBlock createCheckboxSdtBlock(ObjectFactory factory, boolean isChecked) {
//        SdtBlock sdtblock = factory.createSdtBlock();
//
//        SdtPr sdtpr = factory.createSdtPr();
//        sdtblock.setSdtPr(sdtpr);
//
//        CTSdtCheckbox sdtcheckbox = w14ObjectFactory.createCTSdtCheckbox();
//        JAXBElement<CTSdtCheckbox> sdtcheckboxWrapped = w14ObjectFactory.createCheckbox(sdtcheckbox);
//        sdtpr.getRPrOrAliasOrLock().add(sdtcheckboxWrapped);
//
//        org.docx4j.w14.CTOnOff onoff = w14ObjectFactory.createCTOnOff();
//        sdtcheckbox.setChecked(onoff);
//        onoff.setVal(isChecked ? "1" : "0");
//
//        CTSdtCheckboxSymbol sdtcheckboxsymbol = w14ObjectFactory.createCTSdtCheckboxSymbol();
//        sdtcheckbox.setUncheckedState(sdtcheckboxsymbol);
//        sdtcheckboxsymbol.setVal("2610");
//        sdtcheckboxsymbol.setFont("Arial");
//
//        CTSdtCheckboxSymbol sdtcheckboxsymbol2 = w14ObjectFactory.createCTSdtCheckboxSymbol();
//        sdtcheckbox.setCheckedState(sdtcheckboxsymbol2);
//        sdtcheckboxsymbol2.setVal("2612");
//        sdtcheckboxsymbol2.setFont("Arial");
//
//        Id id2 = factory.createId();
//        sdtpr.setId(id2);
//        id2.setVal(BigInteger.valueOf(1487646757));
//
//        SdtContentBlock contentBlock = factory.createSdtContentBlock();
//        R run = factory.createR();
//        Text t = factory.createText();
//
//        org.docx4j.wml.RPr runProperties = factory.createRPr();
//        RFonts runFont = new RFonts();
//        runFont.setAscii("Arial");
//        runFont.setHAnsi("Arial");
//        runProperties.setRFonts(runFont);
//        setFontSize(runProperties, "20");
//
//        run.setRPr(runProperties);
//
//        t.setValue(isChecked ? "☒" : "☐");
//        run.getContent().add(t);
//
//        contentBlock.getContent().add(run);
//        sdtblock.setSdtContent(contentBlock);
//
//        return sdtblock;
//    }
//
//    public static void addTableCellToTableRow(ObjectFactory factory, Tr tableRow, Tc tableCell) {
//
//        tableRow.getContent().add(tableCell);
//    }
//
//    public static Tc createTableCell(ObjectFactory factory, String content,
//            boolean bold, String fontSize, int width) {
//        if (content == null) {
//            content = " ";
//        }
//
//        Tc tableCell = factory.createTc();
//        P paragraph = factory.createP();
//        R runX = factory.createR();
//        org.docx4j.wml.RPr runProperties = factory.createRPr();
//
//        if (bold) {
//            addBoldStyle(runProperties);
//        }
//
//        if (fontSize != null && !fontSize.isEmpty()) {
//            setFontSize(runProperties, fontSize);
//        }
//
//        PPr paragraphProperties = factory.createPPr();
//        Jc justification = factory.createJc();
//        justification.setVal(JcEnumeration.LEFT);
//        paragraphProperties.setJc(justification);
//        PPrBase.Spacing spacing = new PPrBase.Spacing();
//        spacing.setAfter(BigInteger.ZERO);
//        paragraphProperties.setSpacing(spacing);
//        RFonts RunFont = new RFonts();
//        RunFont.setAscii("Arial");
//        RunFont.setHAnsi("Arial");
//        runProperties.setRFonts(RunFont);
//        runX.setRPr(runProperties);
//        paragraph.setPPr(paragraphProperties);
//        addContentToRun(runX, content);
//        paragraph.getContent().add(runX);
//        tableCell.getContent().add(paragraph);
//
//        CTVerticalJc vjc = new CTVerticalJc();
//        vjc.setVal(STVerticalJc.CENTER);
//        TcPr tcPr = factory.createTcPr();
//        tcPr.setVAlign(vjc);
//        tableCell.setTcPr(tcPr);
//        TblWidth tableWidth = new TblWidth();
//        tableWidth.setW(BigInteger.valueOf(width));
//        tcPr.setTcW(tableWidth);
//
//        return tableCell;
//    }
//
//    public static void setTblAllJcAlign(Tbl tbl, JcEnumeration jcType) {
//        if (jcType != null) {
//            setTblJcAlign(tbl, jcType);
//            List<Tr> trList = getTblAllTr(tbl);
//            for (Tr tr : trList) {
//                List<Tc> tcList = getTrAllCell(tr);
//                for (Tc tc : tcList) {
//                    setTcJcAlign(tc, jcType);
//                }
//            }
//        }
//    }
//
//    public static List<Tr> getTblAllTr(Tbl tbl) {
//        List<Object> objList = getAllElementFromObject(tbl, Tr.class);
//        List<Tr> trList = new ArrayList<Tr>();
//        if (objList == null) {
//            return trList;
//        }
//        for (Object obj : objList) {
//            if (obj instanceof Tr) {
//                Tr tr = (Tr) obj;
//                trList.add(tr);
//            }
//        }
//        return trList;
//
//    }
//
//    public static List<Tc> getTrAllCell(Tr tr) {
//        List<Object> objList = getAllElementFromObject(tr, Tc.class);
//        List<Tc> tcList = new ArrayList<Tc>();
//        if (objList == null) {
//            return tcList;
//        }
//        for (Object tcObj : objList) {
//            if (tcObj instanceof Tc) {
//                Tc objTc = (Tc) tcObj;
//                tcList.add(objTc);
//            }
//        }
//        return tcList;
//    }
//
//    public static void setTcJcAlign(Tc tc, JcEnumeration jcType) {
//        if (jcType != null) {
//            List<P> pList = getTcAllP(tc);
//            for (P p : pList) {
//                setParaJcAlign(p, jcType);
//            }
//        }
//    }
//
//    public static void setParaJcAlign(P paragraph, JcEnumeration hAlign) {
//        if (hAlign != null) {
//            PPr pprop = paragraph.getPPr();
//            if (pprop == null) {
//                pprop = new PPr();
//                paragraph.setPPr(pprop);
//            }
//            Jc align = new Jc();
//            align.setVal(hAlign);
//            pprop.setJc(align);
//        }
//    }
//
//    public static List<P> getTcAllP(Tc tc) {
//        List<Object> objList = getAllElementFromObject(tc, P.class);
//        List<P> pList = new ArrayList<P>();
//        if (objList == null) {
//            return pList;
//        }
//        for (Object obj : objList) {
//            if (obj instanceof P) {
//                P p = (P) obj;
//                pList.add(p);
//            }
//        }
//        return pList;
//    }
//
//    public static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
//        List<Object> result = new ArrayList<Object>();
//        if (obj instanceof JAXBElement) {
//            obj = ((JAXBElement<?>) obj).getValue();
//        }
//
//        if (obj.getClass().equals(toSearch)) {
//            result.add(obj);
//        } else if (obj instanceof ContentAccessor) {
//            List<?> children = ((ContentAccessor) obj).getContent();
//            for (Object child : children) {
//                result.addAll(getAllElementFromObject(child, toSearch));
//            }
//
//        }
//        return result;
//    }
//
//    public static void setTblJcAlign(Tbl tbl, JcEnumeration jcType) {
//        if (jcType != null) {
//            TblPr tblPr = new TblPr();
//            Jc jc = new Jc();
//            jc.setVal(jcType);
//            tblPr.setJc(jc);
//            tbl.setTblPr(tblPr);
//        }
//    }
//
//    public static TblPr getTblPr(Tbl table) {
//        return table.getTblPr();
//    }
//
//    public static void createDocxFile(WordprocessingMLPackage wordMLPackage, String fileName) throws Docx4JException {
//        wordMLPackage.save(new File(System.getProperty("user.dir"),
//                fileName + ".docx"));
//    }
//
//    public static void addUnderline(org.docx4j.wml.RPr runProperties) {
//        U underline = new U();
//        underline.setVal(UnderlineEnumeration.SINGLE);
//        runProperties.setU(underline);
//    }
//
//    public static void removeBoldStyle(org.docx4j.wml.RPr runProperties) {
//        runProperties.getB().setVal(false);
//    }
//
//    public static void removeThemeFontInformation(org.docx4j.wml.RPr runProperties) {
//        runProperties.getRFonts().setAsciiTheme(null);
//        runProperties.getRFonts().setHAnsiTheme(null);
//    }
//
//    public static void changeFontSize(org.docx4j.wml.RPr runProperties, int fontSize) {
//        HpsMeasure size = new HpsMeasure();
//        size.setVal(BigInteger.valueOf(fontSize));
//        runProperties.setSz(size);
//    }
//
//    public static void changeFontToArial(org.docx4j.wml.RPr runProperties) {
//        RFonts runFont = new RFonts();
//        runFont.setAscii("Arial");
//        runFont.setHAnsi("Arial");
//        runProperties.setRFonts(runFont);
//    }
//
//    public static org.docx4j.wml.RPr getRunPropertiesAndRemoveThemeInfo(Style style) {
//        // We only want to change some settings, so we get the existing run
//        // properties from the style.
//        org.docx4j.wml.RPr rpr = style.getRPr();
//        removeThemeFontInformation(rpr);
//        return rpr;
//    }
//
//    public static Relationship createFooterPart(WordprocessingMLPackage wordMLPackage, ObjectFactory factory) throws InvalidFormatException {
//        FooterPart footerPart = new FooterPart();
//        footerPart.setPackage(wordMLPackage);
//
//        footerPart.setJaxbElement(createFooter(factory, "Text"));
//
//        return wordMLPackage.getMainDocumentPart().addTargetPart(footerPart);
//    }
//
//    public static Ftr createFooter(ObjectFactory factory, String content) {
//        if (content == null) {
//            content = " ";
//        }
//
//        Ftr footer = factory.createFtr();
//        P paragraph = factory.createP();
//        R run = factory.createR();
//        Text text = new Text();
//        text.setValue(content);
//        run.getContent().add(text);
//        paragraph.getContent().add(run);
//        footer.getContent().add(paragraph);
//        return footer;
//    }
//
//    public static void createFooterReference(WordprocessingMLPackage wordMLPackage, ObjectFactory factory, Relationship relationship) {
//        List<SectionWrapper> sections
//                = wordMLPackage.getDocumentModel().getSections();
//
//        SectPr sectionProperties = sections.get(sections.size() - 1).getSectPr();
//        // There is always a section wrapper, but it might not contain a sectPr
//        if (sectionProperties == null) {
//            sectionProperties = factory.createSectPr();
//            wordMLPackage.getMainDocumentPart().addObject(sectionProperties);
//            sections.get(sections.size() - 1).setSectPr(sectionProperties);
//        }
//
//        FooterReference footerReference = factory.createFooterReference();
//        footerReference.setId(relationship.getId());
//        footerReference.setType(HdrFtrRef.DEFAULT);
//        sectionProperties.getEGHdrFtrReferences().add(footerReference);
//    }
//
//    public static void addTableCell(ObjectFactory factory, Tr tr, P paragraph) {
//        Tc tc1 = factory.createTc();
//        tc1.getContent().add(paragraph);
//        tr.getContent().add(tc1);
//    }
//
//    public static void addPageBreak(MainDocumentPart mainDocumentPart, ObjectFactory factory) {
//
//        Br breakObj = new Br();
//        breakObj.setType(STBrType.PAGE);
//
//        P paragraph = factory.createP();
//        paragraph.getContent().add(breakObj);
//        mainDocumentPart.getJaxbElement().getBody().getContent().add(paragraph);
//    }
//
//    public static SectPr getDocumentSeparator(WordprocessingMLPackage wordMLPackage) {
//        SectPr sectPr = wordMLPackage.getMainDocumentPart().getJaxbElement().getBody().getSectPr();
//        if (sectPr == null) {
//            // Maybe the last P already contains one?
//            // Presumably Word would reuse this.
//            List all = wordMLPackage.getMainDocumentPart().getContent();
//            Object last = all.get(all.size() - 1);
//            if (last instanceof P) {
//                if (((P) last).getPPr() != null && ((P) last).getPPr().getSectPr() != null) {
//                    sectPr = ((P) last).getPPr().getSectPr();
//                }
//            }
//        }
//
//        // <w:pgNumType w:start="1"/>
//        CTPageNumber pageNumber = sectPr.getPgNumType();
//        if (pageNumber == null) {
//            pageNumber = Context.getWmlObjectFactory().createCTPageNumber();
//            sectPr.setPgNumType(pageNumber);
//        }
//        pageNumber.setStart(BigInteger.ONE);
//        return sectPr;
//    }
//
//    public static void addContentToRun(R run, String content) {
//        
//        if (content != null) {
//            ObjectFactory factoryX = Context.getWmlObjectFactory();
//            String[] contentList = content.split("\n");
//            Br breakLine = factoryX.createBr();
//            if (contentList != null) {
//                for (int i = 0; i < contentList.length; i++) {
//                    Text text = factoryX.createText();
//                    text.setSpace("preserve");
//                    text.setValue(contentList[i]);
//                    run.getContent().add(text);
//                    run.getContent().add(breakLine);
//                }
//            } else {
//                Text text = factoryX.createText();
//                text.setValue(content);
//                run.getContent().add(text);
//            }
//        } 
//
//    }
}
