package ru.itmo.docxodtconverter.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import ru.itmo.docxodtconverter.service.ParseService;

@Controller
public class ConverterController {
    private final ParseService documentParserService;

    public ConverterController(ParseService documentParserService) {
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
            documentParserService.parseToAscii(file);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return "redirect:/";
    }
}
