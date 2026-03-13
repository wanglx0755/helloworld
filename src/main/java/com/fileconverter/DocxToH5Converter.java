package com.fileconverter;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.LineSpacingRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

import java.io.IOException;
import java.io.InputStream;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.regex.Pattern;

public class DocxToH5Converter {
    private static final Pattern CN_SECTION_HEADING = Pattern.compile("^[一二三四五六七八九十百千]+、.+");
    private static final Pattern NUM_SECTION_HEADING = Pattern.compile("^\\d+(?:\\.\\d+)*[、.．].+");

    public static void main(String[] args) {
        if (args.length < 1) {
            System.err.println("用法: java -jar docx-to-h5-converter-1.0.0-jar-with-dependencies.jar <input.docx> [output.html]");
            System.exit(1);
        }

        Path input = Paths.get(args[0]).toAbsolutePath().normalize();
        if (!Files.exists(input) || !Files.isRegularFile(input)) {
            System.err.println("输入文件不存在: " + input);
            System.exit(2);
        }

        String inputFileName = input.getFileName().toString();
        if (!inputFileName.toLowerCase().endsWith(".docx")) {
            System.err.println("请输入 .docx 文件: " + input);
            System.exit(3);
        }

        Path output;
        if (args.length >= 2) {
            output = Paths.get(args[1]).toAbsolutePath().normalize();
        } else {
            String baseName = inputFileName.substring(0, inputFileName.length() - 5);
            output = input.getParent().resolve(baseName + ".html");
        }

        try {
            convertDocxToHtml(input, output);
            System.out.println("转换完成: " + output);
        } catch (Exception e) {
            System.err.println("转换失败: " + e.getMessage());
            e.printStackTrace(System.err);
            System.exit(4);
        }
    }

    private static void convertDocxToHtml(Path docxPath, Path htmlPath) throws IOException {
        try (InputStream is = Files.newInputStream(docxPath);
             XWPFDocument document = new XWPFDocument(is);
             Writer writer = Files.newBufferedWriter(htmlPath, StandardCharsets.UTF_8)) {

            StringBuilder html = new StringBuilder(8192);
            html.append("<!DOCTYPE html>\n");
            html.append("<html lang=\"zh-CN\">\n");
            html.append("<head>\n");
            html.append("  <meta charset=\"UTF-8\">\n");
            html.append("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n");
            html.append("  <title>").append(escapeHtml(docxPath.getFileName().toString())).append("</title>\n");
            html.append("  <style>\n");
            html.append("    body { max-width: 900px; margin: 40px auto; padding: 0 16px; line-height: 1.8; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif; color: #222; }\n");
            html.append("    p { margin: 0 0 1em 0; }\n");
            html.append("    h1, h2, h3, h4 { font-weight: 700; }\n");
            html.append("    table { border-collapse: collapse; margin: 1em 0; width: 100%; }\n");
            html.append("    th, td { border: 1px solid #bbb; padding: 8px; vertical-align: top; }\n");
            html.append("  </style>\n");
            html.append("</head>\n");
            html.append("<body>\n");

            for (IBodyElement element : document.getBodyElements()) {
                if (element instanceof XWPFParagraph) {
                    appendParagraph((XWPFParagraph) element, html);
                } else if (element instanceof XWPFTable) {
                    appendTable((XWPFTable) element, html);
                }
            }

            html.append("</body>\n");
            html.append("</html>\n");

            writer.write(html.toString());
        }
    }

    private static void appendParagraph(XWPFParagraph paragraph, StringBuilder html) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null || runs.isEmpty()) {
            html.append("<p>&nbsp;</p>\n");
            return;
        }

        String tag = resolveParagraphTag(paragraph);
        html.append("<").append(tag);

        StringBuilder style = new StringBuilder();
        int indentationLeft = paragraph.getIndentationLeft();
        if (indentationLeft > 0) {
            style.append("padding-left:").append(indentationLeft / 20).append("pt;");
        }
        int firstLineIndent = paragraph.getIndentationFirstLine();
        if (firstLineIndent > 0) {
            style.append("text-indent:").append(firstLineIndent / 20).append("pt;");
        }
        appendAlignmentStyle(paragraph, style);
        appendSpacingStyle(paragraph, style);
        if (isLikelySectionHeading(paragraph)) {
            style.append("font-weight:700;");
        }

        if (style.length() > 0) {
            html.append(" style=\"").append(style).append("\"");
        }

        html.append(">");

        for (XWPFRun run : runs) {
            appendRun(run, html);
        }

        html.append("</").append(tag).append(">\n");
    }

    private static boolean isLikelySectionHeading(XWPFParagraph paragraph) {
        String text = paragraph.getText();
        if (text == null) {
            return false;
        }
        String trimmed = text.trim();
        if (trimmed.isEmpty()) {
            return false;
        }
        return CN_SECTION_HEADING.matcher(trimmed).matches()
                || NUM_SECTION_HEADING.matcher(trimmed).matches();
    }

    private static void appendAlignmentStyle(XWPFParagraph paragraph, StringBuilder style) {
        ParagraphAlignment alignment = paragraph.getAlignment();
        if (alignment == null) {
            return;
        }
        switch (alignment) {
            case CENTER:
                style.append("text-align:center;");
                break;
            case RIGHT:
                style.append("text-align:right;");
                break;
            case BOTH:
                style.append("text-align:justify;");
                break;
            case DISTRIBUTE:
                style.append("text-align:justify;");
                break;
            default:
                break;
        }
    }

    private static void appendSpacingStyle(XWPFParagraph paragraph, StringBuilder style) {
        int beforeTwips = paragraph.getSpacingBefore();
        int afterTwips = paragraph.getSpacingAfter();
        if (beforeTwips > 0) {
            style.append("margin-top:").append(beforeTwips / 20).append("pt;");
        } else {
            int beforeLines = paragraph.getSpacingBeforeLines();
            if (beforeLines > 0) {
                style.append("margin-top:").append(beforeLines / 100.0).append("em;");
            }
        }
        if (afterTwips > 0) {
            style.append("margin-bottom:").append(afterTwips / 20).append("pt;");
        } else {
            int afterLines = paragraph.getSpacingAfterLines();
            if (afterLines > 0) {
                style.append("margin-bottom:").append(afterLines / 100.0).append("em;");
            }
        }

        double spacing = paragraph.getSpacingBetween();
        if (spacing > 0) {
            LineSpacingRule lineRule = paragraph.getSpacingLineRule();
            if (lineRule == LineSpacingRule.AUTO) {
                style.append("line-height:").append(spacing).append(";");
            } else {
                style.append("line-height:").append(spacing).append("pt;");
            }
        }
    }

    private static String resolveParagraphTag(XWPFParagraph paragraph) {
        String style = paragraph.getStyle();
        String normalizedStyle = style == null ? "" : style.replace(" ", "").toLowerCase();
        if (!normalizedStyle.isEmpty()) {
            if (normalizedStyle.contains("heading1") || normalizedStyle.contains("title") || normalizedStyle.contains("标题1")) {
                return "h1";
            }
            if (normalizedStyle.contains("heading2") || normalizedStyle.contains("标题2")) {
                return "h2";
            }
            if (normalizedStyle.contains("heading3") || normalizedStyle.contains("标题3")) {
                return "h3";
            }
            if (normalizedStyle.contains("heading4") || normalizedStyle.contains("标题4")) {
                return "h4";
            }
        }
        return "p";
    }

    private static void appendRun(XWPFRun run, StringBuilder html) {
        String text = getRunText(run);
        if (text.isEmpty()) {
            return;
        }

        String escaped = escapeHtml(text).replace("\n", "<br/>");

        boolean hasWrap = run.isBold() || run.isItalic() || run.getUnderline() != null
                || run.getColor() != null || run.getFontSize() > 0;

        if (!hasWrap) {
            html.append(escaped);
            return;
        }

        html.append("<span");

        StringBuilder style = new StringBuilder();
        if (run.isBold()) {
            style.append("font-weight:bold;");
        }
        if (run.isItalic()) {
            style.append("font-style:italic;");
        }
        if (run.getUnderline() != null && run.getUnderline().name().toLowerCase().contains("none") == false) {
            style.append("text-decoration:underline;");
        }
        if (run.getColor() != null) {
            style.append("color:#").append(run.getColor()).append(";");
        }
        if (run.getFontSize() > 0) {
            style.append("font-size:").append(run.getFontSize()).append("pt;");
        }

        if (style.length() > 0) {
            html.append(" style=\"").append(style).append("\"");
        }
        html.append(">").append(escaped).append("</span>");
    }

    private static String getRunText(XWPFRun run) {
        StringBuilder sb = new StringBuilder();
        List<CTText> textList = run.getCTR().getTList();
        if (textList != null && !textList.isEmpty()) {
            for (CTText ctText : textList) {
                if (ctText != null && ctText.getStringValue() != null) {
                    sb.append(ctText.getStringValue());
                }
            }
        } else {
            String fallback = run.toString();
            if (fallback != null) {
                sb.append(fallback);
            }
        }

        int carriageReturns = run.getCTR().sizeOfCrArray();
        for (int j = 0; j < carriageReturns; j++) {
            sb.append('\n');
        }

        int lineBreaks = run.getCTR().sizeOfBrArray();
        for (int j = 0; j < lineBreaks; j++) {
            sb.append('\n');
        }

        int tabs = run.getCTR().sizeOfTabArray();
        for (int j = 0; j < tabs; j++) {
            sb.append("    ");
        }

        return sb.toString();
    }

    private static void appendTable(XWPFTable table, StringBuilder html) {
        html.append("<table>\n");
        for (XWPFTableRow row : table.getRows()) {
            html.append("  <tr>");
            for (XWPFTableCell cell : row.getTableCells()) {
                html.append("<td>");
                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                if (paragraphs == null || paragraphs.isEmpty()) {
                    html.append("&nbsp;");
                } else {
                    for (XWPFParagraph p : paragraphs) {
                        appendParagraph(p, html);
                    }
                }
                html.append("</td>");
            }
            html.append("</tr>\n");
        }
        html.append("</table>\n");
    }

    private static String escapeHtml(String input) {
        StringBuilder sb = new StringBuilder(input.length() + 16);
        for (char c : input.toCharArray()) {
            switch (c) {
                case '&':
                    sb.append("&amp;");
                    break;
                case '<':
                    sb.append("&lt;");
                    break;
                case '>':
                    sb.append("&gt;");
                    break;
                case '"':
                    sb.append("&quot;");
                    break;
                case '\'':
                    sb.append("&#39;");
                    break;
                default:
                    sb.append(c);
            }
        }
        return sb.toString();
    }
}
