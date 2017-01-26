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
import java.io.File

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
