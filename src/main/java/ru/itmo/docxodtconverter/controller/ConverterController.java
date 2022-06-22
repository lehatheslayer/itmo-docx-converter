package ru.itmo.docxodtconverter.controller;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import ru.itmo.docxodtconverter.service.DocumentParserService;

import java.io.IOException;
import java.util.List;

@Controller
public class ConverterController {
    private final DocumentParserService documentParserService;

    public ConverterController(DocumentParserService documentParserService) {
        this.documentParserService = documentParserService;
    }

    @GetMapping("/")
    public String mainPage() {
        return "home";
    }

    @PostMapping("/upload")
    public String handleUploadFile(@RequestParam("file") MultipartFile file,
                                   @RequestParam("type") String type) {
        System.out.println(file.getContentType());
        System.out.println(type);

        try {
            List<IBodyElement> elements = documentParserService.parse(file);
            documentParserService.build(elements);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return "redirect:/";
    }
}
