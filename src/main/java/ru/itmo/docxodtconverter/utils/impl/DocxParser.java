package ru.itmo.docxodtconverter.utils.impl;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import ru.itmo.docxodtconverter.utils.DocumentParser;

import java.io.IOException;
import java.util.List;

@Service
public class DocxParser implements DocumentParser {
    @Override
    public List<IBodyElement> parse(MultipartFile file) throws InvalidFormatException, IOException {
        try {
            XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(file.getInputStream()));
            return docxFile.getBodyElements();
        } catch (InvalidFormatException e) {
            throw new InvalidFormatException("Invalid format of file");
        } catch (IOException e) {
            throw new IOException(e.getMessage());
        }
    }
}
