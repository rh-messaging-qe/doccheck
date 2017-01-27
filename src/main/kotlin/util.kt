import com.sun.star.beans.PropertyValue
import com.sun.star.beans.XPropertySet
import com.sun.star.container.XNameContainer
import com.sun.star.document.XActionLockable
import com.sun.star.frame.XComponentLoader
import com.sun.star.frame.XModel
import com.sun.star.frame.XStorable
import com.sun.star.lang.XComponent
import com.sun.star.lang.XMultiServiceFactory
import com.sun.star.lang.XServiceInfo
import com.sun.star.text.XTextContent
import com.sun.star.xml.dom.XDocument
import org.xml.sax.InputSource
import java.io.*
import java.net.URL
import java.nio.file.Path
import java.nio.file.Paths
import java.util.*
import java.util.regex.Pattern
import javax.xml.parsers.DocumentBuilderFactory
import javax.xml.xpath.XPathFactory

fun withLockedUI(sheetdocument: XDocument, f: () -> Any) {
    val actionInterface = sheetdocument.query(XActionLockable::class.java)

    // lock all actions
    actionInterface.addActionLock()

    f()

    // remove all locks, the user see all changes
    actionInterface.removeActionLock()
}

fun Open(componentLoader: XComponentLoader, source: String): XComponent {
    val document = componentLoader.loadComponentFromURL(source, "_default", 0, emptyArray())
    return document
}

fun OpenHTML(componentLoader: XComponentLoader, source: String): XComponent {
    val propertyValues = arrayOf(
            PropertyValue().apply {
                Name = "FilterName"
                // Just a magic constant which works, found in
                // /usr/lib64/libreoffice/share/registry/writer.xcd
                Value = "HTML (StarWriter)"
            }
    )
    val document = componentLoader.loadComponentFromURL(source, "_default", 0, propertyValues)
    return document
}

fun embedAllImages(textDocument: XModel) {
    // http://dev.api.openoffice.narkive.com/tYCY7Kqs/insert-embedded-graphic-by-breaking-the-image-link
    // https://github.com/LibreOffice/core/blob/dcc92a7cb5aa1faa711c8da7f7d8ecee0a192c25/embeddedobj/test/Container1/EmbedContApp.java
    // http://stackoverflow.com/questions/5014901/export-images-and-graphics-in-doc-files-to-images-in-openoffice
    // http://140.211.11.67/en/forum/viewtopic.php?t=50114&p=228144

    // https://github.com/LibreOffice/noa-libre/wiki/convert_images_link_to_embed
    val multiServiceFactory = textDocument.query(XMultiServiceFactory::class.java)
    val graphicObjectsSupplier = textDocument.query(com.sun.star.text.XTextGraphicObjectsSupplier::class.java)
    val nameAccess = graphicObjectsSupplier.graphicObjects
    for (elementName in nameAccess.elementNames) {
        val xImageAny = nameAccess.getByName(elementName) as com.sun.star.uno.Any
        val xImage = xImageAny.`object` as XTextContent
        val xInfo = xImage.query(XServiceInfo::class.java)
        if (xInfo.supportsService("com.sun.star.text.TextGraphicObject")) {
            val xPropSet = xImage.query(XPropertySet::class.java)
            val displayName = xPropSet.getPropertyValue("LinkDisplayName").toString()
            val graphicURL = xPropSet.getPropertyValue("GraphicURL").toString()
            //only ones that are not embedded
            if (!graphicURL.contains("vnd.sun.")) {
                val xBitmapContainer = multiServiceFactory.createInstance("com.sun.star.drawing.BitmapTable").query(XNameContainer::class.java)
                if (!xBitmapContainer.hasByName(displayName)) {
                    xBitmapContainer.insertByName(displayName, graphicURL)
                    val newGraphicURL = xBitmapContainer.getByName(displayName).toString()
                    xPropSet.setPropertyValue("GraphicURL", newGraphicURL)
                }
            }
        }
    }
}

fun SaveAsFodt(document: XComponent, target: String) {
    val storable = document.query(XStorable::class.java)
    val propertyValues = arrayOf(
            PropertyValue().apply {
                Name = "Overwrite"
                Value = true
            },
            PropertyValue().apply {
                Name = "FilterName"
                Value = "OpenDocument Text Flat XML"
            }
    )
    storable.storeAsURL(target, propertyValues)
}

fun SaveAsPdf(document: Any, target: String) {
    val storable = document.query(XStorable::class.java)
    val propertyValues = arrayOf(
            PropertyValue().apply {
                Name = "Overwrite"
                Value = true
            },
            PropertyValue().apply {
                Name = "FilterName"
                Value = "writer_pdf_Export"
            }
    )
    storable.storeToURL(target, propertyValues)
}

fun mirrorDocPages(dir: File, urls: Array<String>) {
    val builder = ProcessBuilder()
            .command("wget", "--no-check-certificate", "-EHkKp", *urls)
            .directory(dir)
            .redirectOutput(File("/dev/stdout"))
            .redirectError(File("/dev/stderr"))
    val process = builder.start()
    process.waitFor()

    if (process.exitValue() != 0) {
        println("[FAIL] wget failed to download some files, see log above")
    }
}

fun downloadDocPagesFromXml(dir: Path, urls: Iterable<String>): List<Path> {
    return urls.map { url -> downloadDocPageFromXml(dir, URL(url)) }
}

fun downloadDocPageFromXml(dir: Path, url: URL): Path {
    val text = downloadPageXml(dir, url)
    //println(text)
    val page = parseDocPageXml(text)
    //println(page)
    val output = dir.resolve(Paths.get(url.path).fileName.toString() + ".html")
    output.toFile().writeText(
            """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Title</title>
</head>
<body>
${page.body}
</body>
</html>""")
    return output
}

fun parseDocPageId(scannable: InputStream): String {
    val pattern = Pattern.compile(
            """<meta name="revision" content="n_(\d+)_introducing-red-hat-jboss-a-mq-7_version_7.0-Beta_edition_1.0_release_(\d+)-revision_(\d+)" />""")
    // http://stackoverflow.com/questions/3013669/performing-regex-on-a-stream
    val scanner = Scanner(scannable)
    val s = scanner.findWithinHorizon(pattern, 0)
    if (s != null) {
        val m = pattern.matcher(s)
        m.find()
        val id = m.group(1)
        val release = m.group(2)
        val revision = m.group(3)
        return id
    }
    throw RuntimeException("Parsing id failed")
}

fun getUrl(domain: String, id: String): URL = URL("https://$domain/api/redhat_node/$id")

fun downloadPageXml(dir: Path, url: URL): String {
    val domain = url.host
    val id = parseDocPageId(url.openStream())
    val text = BufferedReader(InputStreamReader(getUrl(domain, id).openStream())).readText()
    dir.resolve("$id.xml").toFile().writeText(text)
    return text
}

fun parseDocPageXml(text: String): DocPage {
    val factory = DocumentBuilderFactory.newInstance()
    val builder = factory.newDocumentBuilder()
    val doc = builder.parse(InputSource(StringReader(text)))  // such ugly
    val xPathfactory = XPathFactory.newInstance()
    val xpath = xPathfactory.newXPath()

    val xVersion = xpath.compile("//result/vid")
    val xBody = xpath.compile("//result/body/en/item/value")

    val version = xVersion.evaluate(doc)
    val body = xBody.evaluate(doc)
    return DocPage(version, body)
}

data class DocPage(val version: String, val body: String)
