package ru.itmo.docxodtconverter.utils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

public interface DocumentParser {
    List<IBodyElement> parse(MultipartFile file) throws InvalidFormatException, IOException;
}
