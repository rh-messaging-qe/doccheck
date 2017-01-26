import com.sun.star.beans.PropertyValue
import com.sun.star.frame.XDispatchHelper
import com.sun.star.frame.XDispatchProvider
import com.sun.star.lang.XComponent
import com.sun.star.lang.XMultiComponentFactory
import com.sun.star.script.provider.XScriptContext
import com.sun.star.sheet.XSpreadsheet
import com.sun.star.uno.UnoRuntime
import java.nio.file.Paths

/**
 * Extension function query simplifies calling UnoRuntime.queryInterface.
 *
 * The function takes a Java class, so a signature clash with UNO function
 * is unlikely, despite generic name. Naming it get would be too generic.
 */
fun <T> Any.query(anInterface: Class<T>): T {
    // TODO: taking KClass did not work for me
    return UnoRuntime.queryInterface(anInterface, this)
}

fun <T> withCatch(expression: () -> T): T? {
    return try {
        expression()
    } catch (_: Exception) {
        null
    }
}

object ComparisonModule {
    var scriptContext: XScriptContext? = null

    @JvmStatic fun CompareAllDocuments(xScriptContext: XScriptContext) {
        try {
            scriptContext = xScriptContext

            if (!promptUserForGoAhead()) {
                return
            }

            val sheets = xScriptContext.document.query(com.sun.star.sheet.XSpreadsheetDocument::class.java)
            val sheet = withCatch { sheets.sheets.getByName("Comparisons") }?.query(XSpreadsheet::class.java)

            if (sheet == null) {
                showErrorBox("Document does not have a 'Comparisons' sheet.")
                return
            }

            for (row in 2..1000) {
                val cellFinished = sheet.getCellByPosition(0, row)
                val cellOriginal = sheet.getCellByPosition(1, row)
                val cellNew = sheet.getCellByPosition(2, row)
                val cellDestination = sheet.getCellByPosition(3, row)

                if (cellFinished.formula != "Finished" && cellOriginal.formula != "" && cellNew.formula != "" && cellDestination.formula != "") {
                    val doc = compareDocs(xScriptContext, cellNew.formula, cellOriginal.formula)
                    val target = Paths.get(xScriptContext.document.url).resolve(cellDestination.formula).toString()
                    SaveAsFodt(doc, target)
                    SaveAsPdf(doc, target.substringBeforeLast('.') + ".pdf")
                    doc.dispose()
                    cellFinished.formula = "Finished"
                }

            }

            showMessageBox("Finished.")

        } catch(e: Throwable) {
            e.printStackTrace()
            throw e
        }
    }

    private fun showErrorBox(message: String) {
        val title = "Download documentation?"
        GUI(scriptContext!!).showErrorBox(title, message)
    }

    private fun showMessageBox(message: String) {
        val title = "Download documentation?"
        GUI(scriptContext!!).showErrorBox(title, message)
    }

    /**
     * Method promptUserForGoAhead displays a QueryBox to the user
     * @return true, if user wishes to go ahead
     */
    fun promptUserForGoAhead(): Boolean {
        val title = "Download documentation?"
        val message = "The spreadsheet will now download documentation files."
        return GUI(scriptContext!!).showQueryBox(title, message)
    }

}

//https://svn.apache.org/repos/asf/ofbiz/tags/REL-10.04.05/applications/content/src/org/ofbiz/content/openoffice/OpenOfficeServices.java
fun compareDocs(scriptContext: XScriptContext, newer: String, older: String): XComponent {
    val documentPath = Paths.get(scriptContext.document.url)

    val newerUrl = documentPath.resolveSibling(newer)
    val olderUrl = documentPath.resolveSibling(older)

    val componentLoader = scriptContext.desktop.currentFrame.query(com.sun.star.frame.XComponentLoader::class.java)
    val document = componentLoader.loadComponentFromURL(newerUrl.toString(), "_default", 0, arrayOf<PropertyValue>())

    val componentContext = scriptContext.componentContext
    val multiComponentFactory = componentContext.serviceManager.query(XMultiComponentFactory::class.java)
    val dispatchHelperObj = multiComponentFactory.createInstanceWithContext("com.sun.star.frame.DispatchHelper", scriptContext.componentContext)
    val dispatchHelper = dispatchHelperObj.query(XDispatchHelper::class.java)
    val dispatchProvider = scriptContext.desktop.currentFrame.query(XDispatchProvider::class.java)

    val propertyValues = arrayOf(
            with(PropertyValue()) {
                Name = "URL"
                Value = olderUrl.toString()
                this
            },
            with(PropertyValue()) {
                Name = "NoAcceptDialog"
                Value = true
                this
            }
    )
    dispatchHelper.executeDispatch(dispatchProvider, ".uno:CompareDocuments", "", 0, propertyValues)

    return document
}

