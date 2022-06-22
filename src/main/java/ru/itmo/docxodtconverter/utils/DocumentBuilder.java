package ru.itmo.docxodtconverter.utils;

import org.apache.poi.xwpf.usermodel.IBodyElement;

import java.util.List;

public interface DocumentBuilder {
    void traverseBodyElements(List<IBodyElement> elements) throws Exception;
}
