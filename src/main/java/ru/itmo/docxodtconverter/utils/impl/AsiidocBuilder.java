package ru.itmo.docxodtconverter.utils.impl;

import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;
import ru.itmo.docxodtconverter.utils.DocumentBuilder;

import java.util.List;

@Service
public class AsiidocBuilder implements DocumentBuilder {
    private void traversePictures(List<XWPFPicture> pictures) throws Exception {
        for (XWPFPicture picture : pictures) {
            System.out.println(picture);
            XWPFPictureData pictureData = picture.getPictureData();
            System.out.println(pictureData);
        }
    }

    private void traverseRunElements(List<IRunElement> runElements) throws Exception {
        for (IRunElement runElement : runElements) {
            if (runElement instanceof XWPFFieldRun) {
                XWPFFieldRun fieldRun = (XWPFFieldRun)runElement;
//                System.out.println(fieldRun.getClass().getName());
                System.out.println(fieldRun);
                traversePictures(fieldRun.getEmbeddedPictures());
            } else if (runElement instanceof XWPFHyperlinkRun) {
                XWPFHyperlinkRun hyperlinkRun = (XWPFHyperlinkRun)runElement;
//                System.out.println(hyperlinkRun.getClass().getName());
                System.out.println(hyperlinkRun);
                traversePictures(hyperlinkRun.getEmbeddedPictures());
            } else if (runElement instanceof XWPFRun) {
                XWPFRun run = (XWPFRun)runElement;
//                System.out.println(run.getClass().getName());
                System.out.println(run);
                traversePictures(run.getEmbeddedPictures());
            } else if (runElement instanceof XWPFSDT) {
                XWPFSDT sDT = (XWPFSDT)runElement;
                System.out.println(sDT);
                System.out.println(sDT.getContent());
                //ToDo: The SDT may have traversable content too.
            }
        }
    }

    private void traverseTableCells(List<ICell> tableICells) throws Exception {
        for (ICell tableICell : tableICells) {
            if (tableICell instanceof XWPFSDTCell) {
                XWPFSDTCell sDTCell = (XWPFSDTCell)tableICell;
                System.out.println(sDTCell);
                //ToDo: The SDTCell may have traversable content too.
            } else if (tableICell instanceof XWPFTableCell) {
                XWPFTableCell tableCell = (XWPFTableCell)tableICell;
                System.out.println(tableCell);
                traverseBodyElements(tableCell.getBodyElements());
            }
        }
    }

    private void traverseTableRows(List<XWPFTableRow> tableRows) throws Exception {
        for (XWPFTableRow tableRow : tableRows) {
            System.out.println(tableRow);
            traverseTableCells(tableRow.getTableICells());
        }
    }

    @Override
    public void traverseBodyElements(List<IBodyElement> bodyElements) throws Exception {
        for (IBodyElement bodyElement : bodyElements) {
            if (bodyElement instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph)bodyElement;
                System.out.println(paragraph);
                traverseRunElements(paragraph.getIRuns());
            } else if (bodyElement instanceof XWPFSDT) {
                XWPFSDT sDT = (XWPFSDT)bodyElement;
                System.out.println(sDT);
                System.out.println(sDT.getContent());
                //ToDo: The SDT may have traversable content too.
            } else if (bodyElement instanceof XWPFTable) {
                XWPFTable table = (XWPFTable)bodyElement;
                System.out.println(table);
                traverseTableRows(table.getRows());
            }
        }
    }
}
