import com.sun.star.script.provider.XScriptContext
import com.sun.star.text.XTextDocument
import com.sun.star.uno.UnoRuntime
import java.io.File
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

fun main(args: Array<String>) {
    println(args)
    val format = DateTimeFormatter.ISO_INSTANT.format(LocalDateTime.now())
    File("").toPath()
}

object HelloWorld {
    @JvmStatic fun printHW(xScriptContext: XScriptContext) {
        val xDocModel = xScriptContext.document

        // getting the text document object
        val xtextdocument = UnoRuntime.queryInterface(
                XTextDocument::class.java, xDocModel)

        xtextdocument.text

        val xText = xtextdocument.text
        val xTextRange = xText.end
        xTextRange.string = "Hello World (in Kotlin)"
    }

    @JvmStatic fun printkHW(xScriptContext: XScriptContext) {
        val xDocModel = xScriptContext.document

        // getting the text document object
        val xtextdocument = UnoRuntime.queryInterface(
                XTextDocument::class.java, xDocModel)

        xtextdocument.text

        val xText = xtextdocument.text
        val xTextRange = xText.end
        xTextRange.string = "Hello World (in Kotlin)"
    }
}