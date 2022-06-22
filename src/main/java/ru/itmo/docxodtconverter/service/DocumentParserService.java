package ru.itmo.docxodtconverter.service;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.IBodyElement;

import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import ru.itmo.docxodtconverter.utils.DocumentBuilder;
import ru.itmo.docxodtconverter.utils.DocumentParser;

import java.io.IOException;
import java.util.List;

@Service
public class DocumentParserService {
    private final DocumentParser docxParser;
    private final DocumentBuilder asciidocBuilder;

    public DocumentParserService(DocumentParser docxParser, DocumentBuilder asciidocBuilder) {
        this.docxParser = docxParser;
        this.asciidocBuilder = asciidocBuilder;
    }

    public List<IBodyElement> parse(MultipartFile file) throws InvalidFormatException, IOException {
        try {
            return docxParser.parse(file);
        } catch (InvalidFormatException e) {
            throw new InvalidFormatException(e.getMessage());
        } catch (IOException e) {
            throw new IOException(e.getMessage());
        }
    }

    public void build(List<IBodyElement> elements) throws Exception {
        asciidocBuilder.traverseBodyElements(elements);
    }
}
