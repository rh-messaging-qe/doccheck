import com.sun.star.beans.UnknownPropertyException
import com.sun.star.beans.XPropertySet
import com.sun.star.frame.XModel
import com.sun.star.lang.XComponent
import com.sun.star.script.provider.XScriptContext
import com.sun.star.sheet.*
import com.sun.star.table.XCell
import java.net.URL
import java.nio.file.Files
import java.nio.file.Paths
import java.time.format.DateTimeFormatter
import java.time.format.DateTimeFormatterBuilder

object DownloadModule {
    @JvmStatic fun Convert(scriptContext: XScriptContext) {
        try {
            val gui = GUI(scriptContext)

            val version = gui.showInputBox("Please enter version (to be substituted for {version} substring):", "Enter Version Parameter", "")
            if (version.isNullOrEmpty()) {
                gui.showErrorBox("", "No version entered, exiting.")
                return
            }

            DownloadModuleImpl(scriptContext).Convert(version)

            gui.showMessageBox("", "Finished.")
        } catch (e: Exception) {
            e.printStackTrace()
            throw e
        }
    }
}

class DownloadModuleImpl(val scriptContext: XScriptContext) {
    val settings = getSettingsFromSettingsSheet(scriptContext, "Settings")
    val SKIP_DOWNLOAD = "skipDownload"
    val MIRROR_PATH = "mirrorPath"
    fun Convert(version: String) {
        val document = scriptContext.document.query(com.sun.star.sheet.XSpreadsheetDocument::class.java)
        val sheet = document.sheets.getByName("Files").query(XSpreadsheet::class.java)

        val dir = if (settings[MIRROR_PATH].isNullOrEmpty()) Files.createTempDirectory("docmirror_") else Paths.get(settings[MIRROR_PATH])
        val skipDownloads = settings[SKIP_DOWNLOAD] == "1" && !settings[MIRROR_PATH].isNullOrBlank()  // 1 meaning TRUE
        val downloader = DocPageDownloader(dir, skipDownloads)

        for (row in 4..1000) {
            val cellDownload = sheet.getCellByPosition(0, row)
            val cellURL = sheet.getCellByPosition(1, row)
            val cellFileName = sheet.getCellByPosition(2, row)

            if (cellDownload.formula != "Finished" && cellURL.formula.isNotBlank() && cellFileName.formula.isNotBlank()) {
                val page = downloader.downloadDocPage(URL(cellURL.formula))
                println("converting ${cellURL.formula} from body at $page")
                val target = doConvert(scriptContext, cellFileName.formula, page.toUri().toASCIIString(), version)

                cellDownload.formula = "Finished"
                updateDocuments(scriptContext, sheet, row, target, version)
            }
        }
    }

    private fun setCellBackgroundColor(cell: XCell, color: Int) {
        val cellDocProperties = cell.query(XPropertySet::class.java)
        cellDocProperties.setPropertyValue("CellBackColor", color)
    }

    fun updateDocuments(scriptContext: XScriptContext, sheet: XSpreadsheet, row: Int, target: String, version: String) {
        val sheets = scriptContext.document.query(com.sun.star.sheet.XSpreadsheetDocument::class.java).sheets
        val cellDoc = sheet.getCellByPosition(3, row)

        val versionSheet: XSpreadsheet
        if (!scriptContext.document.query(com.sun.star.sheet.XSpreadsheetDocument::class.java).sheets.hasByName(version)) {
            sheets.insertNewByName(version, 1000)
            versionSheet = sheets.getByName(version).query(XSpreadsheet::class.java)

            versionSheet.getCellByPosition(0, 0).formula = "Timestamp"
            versionSheet.getCellByPosition(1, 0).formula = "URL"
            versionSheet.getCellByPosition(2, 0).formula = "File Name Template"
            versionSheet.getCellByPosition(3, 0).formula = "Stored as"
        } else {
            versionSheet = sheets.getByName(version).query(XSpreadsheet::class.java)
        }

        setCellBackgroundColor(cellDoc, 0x000000)

        if (cellDoc.formula != target) {
            val cellRangeAddress = sheet.getCellRangeByPosition(3, row, 3, row)
            val cellRangeMovement = sheet.query(XCellRangeMovement::class.java)
            cellRangeMovement.insertCells(cellRangeAddress.query(XCellRangeAddressable::class.java).rangeAddress, CellInsertMode.RIGHT)

            setCellBackgroundColor(cellDoc, 0xedad53)
            cellDoc.formula = target

            copyRowToEndOfComparisonSheet(versionSheet, sheets.getByName("Comparisons"), row, version)
        }
        copyRowToEndOfVersionSheet(sheet, versionSheet, row)
    }

    private fun copyRowToEndOfVersionSheet(sheet: XSpreadsheet, versionSheet: XSpreadsheet, row: Int) {
        val range = sheet.getCellRangeByPosition(1, row, 3, row)

        var firstEmpty = 0
        while (!versionSheet.getCellByPosition(1, firstEmpty).formula.isBlank()) {
            firstEmpty++
        }

        val now = java.time.LocalDateTime.now()
        val formatter = DateTimeFormatterBuilder()
                .append(DateTimeFormatter.ISO_LOCAL_DATE)
                .appendLiteral(' ')
                .append(DateTimeFormatter.ISO_LOCAL_TIME)
                .toFormatter()
        versionSheet.getCellByPosition(1, firstEmpty).formula = now.format(formatter)
        val cellRangeMovement = sheet.query(XCellRangeMovement::class.java)
        cellRangeMovement.copyRange(versionSheet.getCellByPosition(1, firstEmpty).query(XCellAddressable::class.java).cellAddress, range.query(XCellRangeAddressable::class.java).rangeAddress)
    }

    fun copyRowToEndOfComparisonSheet(comparisonSheet: XSpreadsheet, byName: Any?, row: Int, version: String) {
        if (comparisonSheet.getCellByPosition(3, row).formula.isBlank()) {
            return
        }
        if (comparisonSheet.getCellByPosition(4, row).formula.isBlank()) {
            return
        }

        val origName = comparisonSheet.getCellByPosition(4, row).formula
        val newName = comparisonSheet.getCellByPosition(3, row).formula

        if (lastSuffix(origName, "/") == lastSuffix(newName, "/")) {
            return
        }

        var i = 0
        do {
            i++
        } while (comparisonSheet.getCellByPosition(1, i).formula.isBlank())

        comparisonSheet.getCellByPosition(1, i).formula = origName
        comparisonSheet.getCellByPosition(2, i).formula = newName
        comparisonSheet.getCellByPosition(3, i).formula = combineFileNamesToCompared(origName, newName)

        throw UnsupportedOperationException("not implemented") //To change body of created functions use File | Settings | File Templates.
    }

    private fun combineFileNamesToCompared(origName: String?, newName: String?): String? {
//        preffixUpToLast(newName, "/") + "comparison/" + combineFileNames(/)
        throw UnsupportedOperationException("not implemented") //To change body of created functions use File | Settings | File Templates.
    }

    private fun lastSuffix(origName: String?, s: String): Any {
        throw UnsupportedOperationException("not implemented") //To change body of created functions use File | Settings | File Templates.
    }

    fun findEmptyColInRow(sheet: XSpreadsheet, row: Int, col: Int): Int {
        var c = col
        while (sheet.getCellByPosition(c, row).formula.isNotBlank()) {
            c++
        }
        return c
    }

    fun getUserDefinedStringProperty(doc: XComponent, key: String): String? {
        try {
            val documentPropertiesSupplier = doc.query(com.sun.star.document.XDocumentPropertiesSupplier::class.java)
            val value = documentPropertiesSupplier.documentProperties.userDefinedProperties.query(XPropertySet::class.java).getPropertyValue(key)
            return value as String
        } catch(e: UnknownPropertyException) {
            return null
        }
    }

    fun doConvert(scriptContext: XScriptContext, fileName: String, url: String, version: String): String {
        if (fileName.isBlank() || url.isBlank()) {
            return ""
        }

        val componentLoader = scriptContext.desktop.currentFrame.query(com.sun.star.frame.XComponentLoader::class.java)
        val doc = OpenHTML(componentLoader, url)
        embedAllImages(doc.query(XModel::class.java))
        val res = "somefile" //"$pkg$rev"
        val outputUrl = Paths.get(scriptContext.document.url).resolveSibling(res)

        SaveAsFodt(doc, outputUrl.toString() + ".fodt")
        SaveAsPdf(doc, outputUrl.toString() + ".pdf")
        doc.dispose()

        return res
    }
}
