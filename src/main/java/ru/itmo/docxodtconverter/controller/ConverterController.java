package ru.itmo.docxodtconverter.controller;

import org.springframework.core.io.UrlResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import ru.itmo.docxodtconverter.service.ParseService;

import java.nio.file.Path;
import java.nio.file.Paths;

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
    public ResponseEntity<?> handleUploadFile(@RequestParam("file") MultipartFile file) throws Exception {
        UrlResource resource;
        Path path = Paths.get(ParseService.ASCIIDOC_FILE_NAME);
        try {
            this.documentParserService.parseToAscii(file);
            resource = new UrlResource(path.toUri());
        } catch (Exception e) {
            throw new Exception(e.getMessage());
        }

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + resource.getFilename() + "\"")
                .body(resource);
    }
}
