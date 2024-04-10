package io.comfortchauke.demo1;

import com.cloudmersive.client.ConvertWebApi;
import com.cloudmersive.client.invoker.ApiClient;
import com.cloudmersive.client.invoker.ApiException;
import com.cloudmersive.client.invoker.Configuration;
import com.cloudmersive.client.invoker.auth.ApiKeyAuth;
import com.cloudmersive.client.model.HtmlToOfficeRequest;
import com.steadystate.css.parser.CSSOMParser;
import lombok.SneakyThrows;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.JodConverter;
import org.jodconverter.local.office.LocalOfficeManager;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.io.File;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.w3c.dom.css.CSSRule;
import org.w3c.dom.css.CSSRuleList;
import org.w3c.dom.css.CSSStyleDeclaration;
import org.w3c.dom.css.CSSStyleRule;
import org.w3c.dom.css.CSSStyleSheet;
import org.w3c.css.sac.InputSource;

@SpringBootApplication
public class Demo1Application {

    @SneakyThrows
    public static void main(String[] args) {
        var ctx = SpringApplication.run(Demo1Application.class, args);
        File inputFile = new File("test.html");
        Path outputFile = Path.of("test.odt");
        Path outputFile1 = Path.of("test1.docx");

        String document = Files.readString(Path.of("test.html"));

        var documentInline = inlineCss(style, document);
        Files.write(Path.of("test1.html"), documentInline.getBytes());
//        inputRequest.setHtml(document);
        try {
            OfficeManager officeManager = LocalOfficeManager.install();
            officeManager.start();
            while(!officeManager.isRunning()){
                Thread.sleep(1000);
                System.out.println("Waiting for office to start");
            }
            JodConverter.convert(new ByteArrayInputStream(document.getBytes())).as(DefaultDocumentFormatRegistry.HTML)
                    .to(Path.of("test.odt").toFile()).as(DefaultDocumentFormatRegistry.ODT)
                    .execute();
            System.out.println("start1");
            Thread.sleep(5000);
            JodConverter.convert(new ByteArrayInputStream(documentInline.getBytes())).as(DefaultDocumentFormatRegistry.HTML)
                    .to(Path.of("test1.odt").toFile()).as(DefaultDocumentFormatRegistry.ODT)
                    .execute();
            System.out.println("start2");
        } catch (Exception e) {
            System.err.println("Exception when calling ConvertWebApi#convertWebHtmlToDocx");
            e.printStackTrace();
        }

//        try {
//            HtmlToOpenXMLConverter converter = new HtmlToOpenXMLConverter();
//            WordprocessingMLPackage wordDocument = converter.convert(Files.readString(Path.of("test.html")));
//            wordDocument.save(outputFile1.toFile());
//        } catch (InvalidFormatException e) {
//            e.printStackTrace();
//        }
        ctx.close();
    }


    private static String inlineCss(String css, String html) throws IOException {
        CSSOMParser parser = new CSSOMParser();
        CSSStyleSheet styleSheet = parser.parseStyleSheet(new InputSource(new StringReader(css)), null, null);
        final Document document = Jsoup.parse(html, "UTF-8");
        return inlineCss(styleSheet, document);
    }
        private static String inlineCss(CSSStyleSheet styleSheet, Document document) {
        final CSSRuleList rules = styleSheet.getCssRules();
        final Map<Element, Map<String, String>> elementStyles = new HashMap<>();

        /*
         * For each rule in the style sheet, find all HTML elements that
         * match based on its selector and store the style attributes in the
         * map with the selected element as the key.
         */
        for (int i = 0; i < rules.getLength(); i++) {
            final CSSRule rule = rules.item(i);
            if (rule instanceof CSSStyleRule) {
                final CSSStyleRule styleRule = (CSSStyleRule) rule;
                final String selector = styleRule.getSelectorText();

                // Ignore pseudo classes, as JSoup's selector cannot
                // handle
                // them.
                if (!selector.contains(":")) {
                    final Elements selectedElements = document.select(selector);
                    for (final Element selected : selectedElements) {
                        if (!elementStyles.containsKey(selected)) {
                            elementStyles.put(selected, new LinkedHashMap<String, String>());
                        }

                        final CSSStyleDeclaration styleDeclaration = styleRule.getStyle();

                        for (int j = 0; j < styleDeclaration.getLength(); j++) {
                            final String propertyName = styleDeclaration.item(j);
                            final String propertyValue = styleDeclaration.getPropertyValue(propertyName);
                            final Map<String, String> elementStyle = elementStyles.get(selected);
                            elementStyle.put(propertyName, propertyValue);
                        }

                    }
                }
            }
        }

        /*
         * Apply the style attributes to each element and remove the "class"
         * attribute.
         */
        for (final Map.Entry<Element, Map<String, String>> elementEntry : elementStyles.entrySet()) {
            final Element element = elementEntry.getKey();
            final StringBuilder builder = new StringBuilder();
            for (final Map.Entry<String, String> styleEntry : elementEntry.getValue().entrySet()) {
                builder.append(styleEntry.getKey()).append(":").append(styleEntry.getValue()).append(";");
            }
            builder.append(element.attr("style"));
            element.attr("style", builder.toString());
            element.removeAttr("class");
        }
        return document.html();
    }
    static String style= """
            body {
                 font-family: "Open Sans", sans-serif;
                 font-size: 13px;
               }
                          
               nlc {
                 font-family: "Arial Unicode MS";
               }
                          
               .handwritten {
                 font-family: Satisfy, "Open Sans", sans-serif;
               }
                          
               header {
                 display: block;
                 text-align: center;
               }
                          
               footer {
                 position: running(lastpagefooter);
                 font-family: "Open Sans", sans-serif;
                 font-size: 10px;
                 color: #a0a0a0;
               }
                          
               @page {
                 size: A4 portrait;
                 margin: 1.92cm;
                          
                 @bottom-center {
                   font-family: "Open Sans", sans-serif;
                   font-size: 10px;
                   color: #6f6f6f;
                   content: counter(page);
                 }
                          
                 @bottom-left {
                   content: element(lastpagefooter);
                 }
               }
                          
               h1, .title-document-type {
                 font-size: 18px;
                 text-align: center;
                 font-weight: bold;
                 margin: 3em 0 2em;
               }
                          
               h2, .page-title {
                 font-size: 14px;
                 text-align: center;
                 font-weight: bold;
                 margin: 2em 0;
                 text-transform: uppercase;
               }
                          
               h3 {
                 font-size: 13px;
                 font-weight: bold;
                 margin: 0;
                 text-transform: uppercase;
               }
                          
               a {
                 color: inherit;
                 text-decoration: none;
               }
                          
               table {
                 width: 100%;
                 border-collapse: collapse;
                 -fs-table-paginate: paginate;
                 border-spacing: 0;
               }
                          
               tr {
                 page-break-inside: avoid;
               }
                          
               th, td {
                 vertical-align: top;
                 border: 0.5px solid #D3D3D3;
                 padding: 10px;
               }
                          
               tr:first-child > th, tr:first-child > td {
                 border-top-width: 1px;
               }
                          
               tr:last-child > th, tr:last-child > td {
                 border-bottom-width: 1px;
               }
                          
               th:first-child, td:first-child {
                 border-left-width: 1px;
               }
                          
               th:last-child, td:last-child {
                 border-right-width: 1px;
               }
                          
               .wswb {
                 white-space: pre-line;
                 overflow-wrap: break-word;
                 word-wrap: break-word;
                 word-break: break-word;
               }
                          
               ul {
                 list-style-type: disc;
               }
                          
               p, li {
                 margin: 1em 0;
               }
                          
               ol > li {
                 display: table;
                 width: 100%;
               }
                          
               ol > li:before {
                 display: table-cell;
                 width: 32px;
                 padding-right: 8px;
                 white-space: nowrap;
               }
                          
               ol {
                 list-style-type: none;
                 margin: 0;
                 padding: 0;
                 counter-reset: nbcnt;
               }
                          
               ol > li:before {
                 counter-increment: nbcnt;
                 content: counters(nbcnt, ".") " ";
               }
                          
               ol ol ol {
                 counter-reset: lacnt;
               }
                          
               ol ol ol > li:before {
                 counter-increment: lacnt;
                 content: "(" counter(lacnt, lower-alpha) ")";
               }
                          
               ol ol ol ol {
                 counter-reset: lrcnt;
               }
                          
               ol ol ol ol > li:before {
                 counter-increment: lrcnt;
                 content: "(" counter(lrcnt, lower-roman) ")";
               }
                          
               ol.nb {
                 counter-reset: nbcnt;
               }
                          
               ol.nb > li:before {
                 counter-increment: nbcnt;
                 content: counters(nbcnt, ".") " ";
               }
                          
               ol.alpha {
                 counter-reset: lacnt;
               }
                          
               ol.alpha > li:before {
                 counter-increment: lacnt;
                 content: "(" counter(lacnt, lower-alpha) ")";
               }
                          
               ol.ua {
                 counter-reset: uacnt;
               }
                          
               ol.ua > li:before {
                 counter-increment: uacnt;
                 content: "(" counter(uacnt, upper-alpha) ")";
               }
                          
               ol.roman {
                 counter-reset: lrcnt;
               }
                          
               ol.roman > li:before {
                 counter-increment: lrcnt;
                 content: "(" counter(lrcnt, lower-roman) ")";
               }
                          
               .front-page .company-name {
                 font-size: 32px;
                 text-align: center;
                 font-weight: bold;
               }
                          
               .front-page .document-type {
                 text-align: center;
                 margin: 2em 0 2em;
                 font-size: 28px;
                 font-weight: bold;
               }
                          
               .front-page .parties {
                 text-align: center;
                 margin: 2em 0 2em;
                 font-size: 18px;
                 font-weight: bold;
               }
                          
               .front-page .logo {
                 width: 126px;
                 height: 126px;
                 margin: 140px 0 60px;
               }
                          
               .front-page .c13-logo {
                 width: 140px;
                 height: 78px;
                 margin: 140px 0 60px;
               }
                          
               .front-page .ckc-logo {
                 width: 160px;
                 height: 160px;
                 margin: 140px 0 60px;
               }
                          
               .front-page .no-logo {
                 height: 261px;
               }
                          
               .title-company-name {
                 font-size: 28px;
                 text-align: center;
                 font-weight: bold;
               }
                          
               .bold {
                 font-weight: bold;
               }
                          
               .italic {
                 font-style: italic;
               }
                          
               .r, .right {
                 text-align: right;
               }
                          
               .c, .center {
                 text-align: center;
               }
                          
               .capitalize {
                 text-transform: capitalize;
               }
                          
               .pb {
                 page-break-after: always;
               }
                          
               .pbia {
                 page-break-inside: avoid;
               }
                          
               .pl {
                 padding-left: 2em;
               }
                          
               .sectit {
                 text-transform: uppercase;
                 font-weight: bold;
               }
                          
               .signature-wrapper {
                 float: left;
                 width: 50%;
               }
                          
               .signature-container {
                 margin: auto;
                 width: 284px;
               }
                          
               .signature {
                 display: block;
                 width: 284px;
                 height: 110px;
                 border-bottom: 1px solid #000;
               }
                          
               .signature-name {
                 width: 284px;
                 height: 74px;
                 text-align: center;
                 font-weight: 600;
               }
                          
               .signature-onbehalf {
                 font-weight: 400;
               }
                          
               .signature-onbehalf-small {
                 font-size: 10px;
                 font-weight: 400;
               }
                          
               .signature-date {
                 font-size: 10px;
                 font-weight: 400;
                 color: #65738e;
               }
                          
               .signature > img {
                 width: 100%;
                 height: 100%;
               }
                          
               .signature > img.withSeal {
                 width: 60%;
               }
                          
               .signature-name-witness {
                 font-size: 8px;
               }
                          
               .signature-name-witness > .label {
                 float: left;
                 width: 28%;
                 text-align: right;
                 margin-right: 2%;
               }
                          
               .signature-name-witness > .value {
                 float: left;
                 width: 70%;
               }
                          
               .fl {
                 float: left;
               }
                          
               .w40 {
                 width: 40%;
               }
                          
               .w50 {
                 width: 50%;
               }
                          
               .mr30 {
                 margin-right: 30px;
               }
                          
               .clear {
                 clear: both;
               }
                          
               .box {
                 display: inline-block;
                 width: 100%;
                 min-width: 500px;
                 padding: 12px;
                 border-top: 0;
                 border-right: 0;
                 border-bottom: 1px solid #ccc;
                 border-left: 4px solid #000646;
               }
                          
               .important-notice {
                 border: 2px solid #000646;
                 padding: 12px;
                 margin: 20px;
               }
                          
               .line {
                 width: 130px;
                 border-bottom: 1px solid;
               }
                          
               .round-name {
                 background-color: #000646;
                 color: white;
                 font-weight: 600;
               }
                          
               .terms th {
                 width: 25%;
                 text-align: left;
               }
                          
               .action {
                 color: #02aba0;
                 font-weight: 600;
               }
                          
               .nowrap {
                 white-space: nowrap;
               }
                          
               .small {
                 font-size: 8px;
                 color: #6f6f6f;
               }
                          
               .signature-seal {
                 height: 100px;
                 width: 100px;
                 border-radius: 50%;
                 background-color: red;
                 display: inline-block;
                 margin-bottom: 10px;
               }
                          
                        
            """;

}
