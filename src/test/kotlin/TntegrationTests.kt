import com.google.common.truth.Truth.assertThat
import com.sun.star.beans.XPropertySet
import com.sun.star.bridge.XUnoUrlResolver
import com.sun.star.comp.helper.Bootstrap
import com.sun.star.document.XScriptInvocationContext
import com.sun.star.frame.XComponentLoader
import com.sun.star.frame.XDesktop
import com.sun.star.frame.XModel
import com.sun.star.lang.XMultiServiceFactory
import com.sun.star.script.provider.XScriptContext
import com.sun.star.uno.XComponentContext
import org.junit.AfterClass
import org.junit.jupiter.api.Nested
import org.junit.jupiter.api.Test
import java.io.BufferedReader
import java.io.InputStreamReader
import java.net.URL
import java.nio.file.Files
import java.nio.file.Paths

fun startSOffice(): XScriptContext {
    val context = Bootstrap.bootstrap()
    val multiComponentFactory = context.serviceManager
    val desktop = multiComponentFactory.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    val scriptContext = desktop.query(XScriptContext::class.java)
    return scriptContext
}

fun printAvailableServices(multiServiceFactory: XMultiServiceFactory) {
    print(multiServiceFactory.availableServiceNames.joinToString(", "))
}

class SOfficeConnection {
    val multiServiceFactory: XMultiServiceFactory

    init {
        // http://www.openoffice.org/udk/java/man/AccessingOfficeFromRemote.html
        val localMultiServiceFactory = Bootstrap.createSimpleServiceManager()
        val objectUrlResolver = localMultiServiceFactory.createInstance("com.sun.star.bridge.UnoUrlResolver")
        val xurlresolver = objectUrlResolver.query(XUnoUrlResolver::class.java)
        val objectInitial = xurlresolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ServiceManager")
        multiServiceFactory = objectInitial.query(XMultiServiceFactory::class.java)
    }
}

class ScriptContext(val sofficeConnection: SOfficeConnection, val documentName: String) : XScriptContext {
    private val myDesktop by lazy {
        sofficeConnection.multiServiceFactory.createInstance("com.sun.star.frame.Desktop").query(XDesktop::class.java)
    }

    private val myDocument by lazy {
        val componentLoader = desktop.query(XComponentLoader::class.java)
        val path = Paths.get(documentName).toUri().toString()
        val document = Open(componentLoader, path)
        document.query(XModel::class.java)
    }

    private val myComponentContext by lazy {
        // https://wiki.openoffice.org/wiki/Documentation/DevGuide/ProUNO/Component_Context
        val propertySet = sofficeConnection.multiServiceFactory.query(XPropertySet::class.java)
        val oDefaultContext = propertySet.getPropertyValue("DefaultContext")
        oDefaultContext.query(XComponentContext::class.java)
    }

    override fun getDocument(): XModel = myDocument
    override fun getDesktop(): XDesktop = myDesktop
    override fun getComponentContext(): XComponentContext = myComponentContext
    override fun getInvocationContext(): XScriptInvocationContext = throw UnsupportedOperationException("not implemented")
}

class IntegrationTests {
    val sofficeConnection = SOfficeConnection()
    val scriptContext = ScriptContext(sofficeConnection, "src/test/tests.fods")

    @AfterClass
    fun tearDown() {
        scriptContext.document.dispose()
    }

    @Test fun testGetSettingsFromSettingsSheet() {
        val settings = getSettingsFromSettingsSheet(scriptContext, "Settings")
        assertThat(settings).isEqualTo(mapOf("firstKey" to "firstValue", "secondKey" to ""))
    }

    //    @Nested inner class diffDocumentCreationTest {
    @Test fun comparing_creates_diff_document() {
        val tmpdir = Files.createTempDirectory("")
        val doc = compareDocs(scriptContext, "document1.fodt", "document2.fodt")
        val target = tmpdir.resolve("_diff.fodt")
        val targetUriString = target.toUri().toString()
        SaveAsFodt(doc, targetUriString)
        SaveAsPdf(doc, targetUriString.substringBeforeLast('.') + ".pdf")
        doc.dispose()

        val text = target.toFile().readText(charset = Charsets.UTF_8)
        assertThat(text).containsMatch("""<text:p text:style-name="Title">Document <text:change-start text:change-id="\w+"/>1<text:change-end text:change-id="\w+"/><text:change-start text:change-id="\w+"/><text:span text:style-name="T1">2</text:span><text:change-end text:change-id="\w+"/></text:p>""")
        assertThat(text).containsMatch("""<text:p text:style-name="P1">This is document <text:change-start text:change-id="\w+"/>1<text:change-end text:change-id="\w+"/><text:change-start text:change-id="\w+"/>2<text:change-end text:change-id="\w+"/></text:p>""")

        tmpdir.toFile().deleteRecursively()
    }

    @Nested inner class userDefinedPropertyTest {
        private fun getUserProperty(key: String): String? = DownloadModuleImpl(scriptContext).getUserDefinedStringProperty(scriptContext.document, key)
        @Test fun defined_gives_value() = assertThat(getUserProperty("aTextProperty")).isEqualTo("aValue")
        @Test fun undefined_gives_null() = assertThat(getUserProperty("anUndefinedProperty")).isNull()
    }

    @Nested inner class guiTest {
        val gui = GUI(scriptContext)
        @Test fun testGui() {
            gui.showMessageBox("title", "message")
            gui.showErrorBox("title", "message")
            gui.showInputBox("title", "message", "default")
            gui.showQueryBox("title", "message", "default")
        }
    }

    @Test fun testDoConvert() {
        DownloadModuleImpl(scriptContext).doConvert(scriptContext, "blah.odt", "https://access.redhat.com/documentation/en/red-hat-jboss-a-mq/7.0-beta/single/introducing-red-hat-jboss-a-mq-7/", "001")
    }
}

class WriterIntegrationTests() {
    val sofficeConnection = SOfficeConnection()
    val scriptContext = ScriptContext(sofficeConnection, "src/test/tests.fodt")

    @Test fun testEmbeddedAllImages() {
        val textDocument = scriptContext.document
        embedAllImages(textDocument)
    }

    @Test fun testMirrorDocPages() {
        val urls = arrayOf(
                "https://example.org"
                // the following is more realistic, but takes way too long
                // "https://access.redhat.com/documentation/en/red-hat-jboss-a-mq/7.0-beta/single/introducing-red-hat-jboss-a-mq-7/"
        )
        val dir = Files.createTempDirectory("testdocmirror_")
        mirrorDocPages(dir.toFile(), urls)

        dir.toFile().deleteRecursively()
    }

    private fun readUrl(url: URL): String = BufferedReader(InputStreamReader(url.openStream())).readText()

    @Test fun testGetPageNum() {
        val url = URL("https://access.redhat.com/documentation/en/red-hat-jboss-a-mq/7.0-beta/single/introducing-red-hat-jboss-a-mq-7/")
        //println(readUrl(url))
        val id = parseDocPageId(url.openStream())
        assertThat(id).isEqualTo("2754751")
    }

    @Test fun testDownloadDocPageFromXml() {
        val dir = Files.createTempDirectory("testdoccheckxmls_")
        val url = URL("https://access.redhat.com/documentation/en/red-hat-jboss-a-mq/7.0-beta/single/introducing-red-hat-jboss-a-mq-7/")
        downloadDocPageFromXml(dir, url)

        dir.toFile().deleteRecursively()
    }
}
