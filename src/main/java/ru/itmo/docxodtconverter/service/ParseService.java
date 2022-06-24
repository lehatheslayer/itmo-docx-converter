package ru.itmo.docxodtconverter.service;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * Сервис, который парсит Word-документы в формат ASCIIDoc
 */
@Service
public class ParseService {
    /**
     * Название выходного файла
     */
    public static final String ASCIIDOC_FILE_NAME = "output.adoc";
    public static final String ZIP_FILE_NAME = "result.zip";

    /**
     * Символы пробела и переходов на следующую строку
     */
    private static final String NEXT_LINE_SYMBOL = "\n";
    private static final String DOUBLE_NEXT_LINE_SYMBOL = "\n\n";
    private static final char SPACE_SYMBOL = ' ';

    /**
     * Префиксы заголовков
     */
    private static final String H1_PREFIX = "= ";
    private static final String H2_PREFIX = "== ";
    private static final String H3_PREFIX = "=== ";
    private static final String H4_PREFIX = "==== ";
    private static final String H5_PREFIX = "===== ";
    private static final String H6_PREFIX = "====== ";

    /**
     * Символы форматирования текста
     */
    private static final char BOLD_SYMBOL = '*';
    private static final char ITALIC_SYMBOL = '_';
    private static final char HIGHLIGHTED_SYMBOL = '#';

    /**
     * Префиксы списков
     */
    private static final String LIST_PREFIX = "- ";

    /**
     * Табличные префиксы
     */
    private static final String TABLE_PREFIX = "|===";
    private static final char TABLE_CELL_PREFIX = '|';

    /**
     * Префикс и суффикс для иллюстраций
     */
    private static final String IMAGE_PREFIX = "image::";
    private static final String IMAGE_SUFFIX = "[]";

    /**
     * Названия иллюстраций для добавления их в zip-архив
     */
    private final Set<String> pictures = new HashSet<>();

    /**
     * Метод парсит WORD-документы в ASCIIDoc формат
     *
     * @param file - WORD-документ
     * @throws InvalidFormatException - выбрасывается при неправильном формате файла
     * @throws IOException - выбрасывается при ошибках записи в файл
     */
    public void parseToAscii(MultipartFile file) throws InvalidFormatException, IOException {
        try (FileWriter writer = new FileWriter(ASCIIDOC_FILE_NAME, false)) {
            final XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(file.getInputStream()));
            final List<IBodyElement> bodyElements = docxFile.getBodyElements();

            for (IBodyElement bodyElement : bodyElements) {
                if (bodyElement.getElementType().equals(BodyElementType.PARAGRAPH)) {
                    writeText(writer, (XWPFParagraph) bodyElement);
                }

                if (bodyElement.getElementType().equals(BodyElementType.TABLE)) {
                    final List<List<String>> tableMatrix = readTable((XWPFTable) bodyElement);
                    writeTable(writer, tableMatrix);
                }
            }
        } catch (InvalidFormatException e) {
            throw new InvalidFormatException("Invalid format of file");
        } catch (IOException e) {
            throw new IOException(e.getMessage());
        } finally {
            madeZipArchive();
        }

    }

    /**
     * Метод, добавляющий текстовые данные в ASCIIDoc файл
     *
     * @param writer - объект класса FileWriter, который производит запись в файл
     * @param paragraph - представление параграфа WORD-документа
     * @throws IOException - выбрасывается при ошибках записи в файл
     */
    private void writeText(final FileWriter writer, final XWPFParagraph paragraph) throws IOException {
        final List<XWPFRun> runs = paragraph.getRuns();
        final StringBuilder sb = new StringBuilder();

        for (final XWPFRun run : runs) {
            String text = run.getText(0);
            if (text == null || text.equals(" ")) {
                final List<XWPFPicture> pictures = run.getEmbeddedPictures();
                for (final XWPFPicture picture : pictures) {
                    savePicture(picture);
                    writer.append(IMAGE_PREFIX)
                          .append(picture.getPictureData().getFileName())
                          .append(IMAGE_SUFFIX)
                          .append(DOUBLE_NEXT_LINE_SYMBOL);


                    this.pictures.add(picture.getPictureData().getFileName());
                }

                continue;
            }

            if (paragraph.getCTPPr().isSetNumPr()) {
                if (paragraph.getAlignment().getValue() == 2) {
                    sb.append(H2_PREFIX);
                } else {
                    sb.append(LIST_PREFIX);
                }
            } else if (paragraph.getAlignment().getValue() == 2) {
                sb.append(H2_PREFIX);
            }

            if (run.isBold()) {
                sb.append(BOLD_SYMBOL);
            }
            if (run.isItalic()) {
                sb.append(ITALIC_SYMBOL);
            }
            if (run.isHighlighted()) {
                sb.append(HIGHLIGHTED_SYMBOL);
            }

            if (run.isCapitalized()) {
                sb.append(text.toUpperCase(Locale.ROOT));
            } else {
                sb.append(text);
            }
            if (sb.charAt(sb.length() - 1) == SPACE_SYMBOL) {
                sb.deleteCharAt(sb.length() - 1);
            }

            if (run.isBold()) {
                sb.append(BOLD_SYMBOL);
            }
            if (run.isItalic()) {
                sb.append(ITALIC_SYMBOL);
            }
            if (run.isHighlighted()) {
                sb.append(HIGHLIGHTED_SYMBOL);
            }

            sb.append(SPACE_SYMBOL);
        }

        sb.append(DOUBLE_NEXT_LINE_SYMBOL);

        writer.write(sb.toString());
    }

    private void savePicture(XWPFPicture picture) throws IOException {
        ByteArrayInputStream bis = new ByteArrayInputStream(picture.getPictureData().getData());
        BufferedImage bImage = ImageIO.read(bis);

        ImageIO.write(bImage, "png", new File(picture.getPictureData().getFileName()));
    }

    /**
     * Метод парсит таблицы из WORD-документа в массив массивов из содержимого таблицы
     *
     * @param table - представление WORD-таблицы
     * @return - массив массивов из содержимого таблицы
     */
    private List<List<String>> readTable(XWPFTable table) {
        final List<List<String>> tableMatrix = new ArrayList<>();

        final List<XWPFTableRow> rows = table.getRows();
        for (int i = 0; i < rows.size(); i++) {
            List<XWPFTableCell> cells = rows.get(i).getTableCells();
            tableMatrix.add(new ArrayList<>());

            for (XWPFTableCell cell : cells) {
                tableMatrix.get(i).add(cell.getText());
            }
        }

        return tableMatrix;
    }

    /**
     * Метод записывает таблицу в ASCIIDoc-файл
     *
     * @param writer - объект класса FileWriter, который производит запись в файл
     * @param tableMatrix - массив массивов из содержимого таблицы
     * @throws IOException - выбрасывается при ошибках записи в файл
     */
    private void writeTable(FileWriter writer, List<List<String>> tableMatrix) throws IOException {
        writer.append(TABLE_PREFIX).append(NEXT_LINE_SYMBOL);

        for (final List<String> row : tableMatrix) {
            for (final String cell : row) {
                writer.append(TABLE_CELL_PREFIX).append(cell).append(SPACE_SYMBOL);
            }
            writer.append(NEXT_LINE_SYMBOL);
        }

        writer.append(TABLE_PREFIX);
    }

    /**
     * Создание Zip-архива и добавление в него файлов
     *
     * @throws IOException - выбрасывается при ошибках записи в файл
     */
    private void madeZipArchive() throws IOException {
        final FileOutputStream fos = new FileOutputStream(ZIP_FILE_NAME);
        final ZipOutputStream zipOut = new ZipOutputStream(fos);

        addPicturesToZip(zipOut);
        addAsciiDocToZip(zipOut);

        zipOut.close();
        fos.close();
    }

    /**
     * Добавление AsciiDoc-документа в архив
     *
     * @param zipOut - производит добавление файлов в архив
     * @throws IOException - выбрасывается при ошибках записи в файл
     */
    private void addAsciiDocToZip(final ZipOutputStream zipOut) throws IOException {
        final File fileToZip = new File(ASCIIDOC_FILE_NAME);
        final ZipEntry zipEntry = new ZipEntry(fileToZip.getName());
        zipOut.putNextEntry(zipEntry);

        Files.copy(fileToZip.toPath(), zipOut);
    }

    /**
     * Добавление иллюстраций в архив
     *
     * @param zipOut - производит добавление файлов в архив
     * @throws IOException - выбрасывается при ошибках записи в файл
     */
    private void addPicturesToZip(final ZipOutputStream zipOut) throws IOException {
        for (final String picture : pictures) {
            final File fileToZip = new File(picture);
            final ZipEntry zipEntry = new ZipEntry(fileToZip.getName());
            zipOut.putNextEntry(zipEntry);

            Files.copy(fileToZip.toPath(), zipOut);
        }
    }
}
