import com.sun.star.awt.*
import com.sun.star.beans.XMultiPropertySet
import com.sun.star.lang.XMultiComponentFactory
import com.sun.star.script.provider.XScriptContext
import com.sun.star.uno.UnoRuntime

fun commonPrefixUpToDash(x: String, y: String): Int {
    return 42
}

/*
Function CommonPrefixUpToDash(x As String, y As String) As Integer
'	If IsMissing(x) Or IsArray(x) Then
'		CommonPrefix = ""
'		Exit Function
'	End If

	Dim MaxLength
	Dim I
	Dim LastDash

	MaxLength = Len(x)
	If MaxLength > Len(y) Then MaxLength = Len(y)

	I = 0
	LastDash = 0
	While I < MaxLength - 1 And Mid(x, I + 1, 1) = Mid(y, I + 1, 1)
		I = I + 1
		if (Mid(x, I, 1) = "-") Then LastDash = I
	Wend

	CommonPrefixUpToDash = LastDash
End Function

*/

fun CombineFileNames(x: String, y: String, infix: String): String {
    return "42"
}

/*

Function CombineFileNames(x As String, y As String, Infix As String) As String
	Dim I
	Dim J
	I = CommonPrefixUpToDash(x,y)
'	J = InStr(I, x, ".") - 1	' Find the file extension and strip it
	J = regexi(x, "(\.[^.]+)$", 1)	' Find the file extension and strip it
	if J = -1 Then J = I
	CombineFileNames = Left(x, J) & Infix & Right(y, Len(y) - I)
End Function


 */

class GUI(scriptContext: XScriptContext) {
    val parentWindow = scriptContext.document.currentController.frame.containerWindow
    val parentWindowPeer = UnoRuntime.queryInterface(XWindowPeer::class.java, parentWindow)
    val messageBoxFactory = UnoRuntime.queryInterface(com.sun.star.awt.XMessageBoxFactory::class.java, parentWindowPeer.toolkit)

    val componentContext = scriptContext.componentContext
    val multiComponentFactory = componentContext.serviceManager.query(XMultiComponentFactory::class.java)

    fun showErrorBox(title: String, message: String) {
        val messageBox = messageBoxFactory.createMessageBox(parentWindowPeer, MessageBoxType.ERRORBOX, MessageBoxButtons.BUTTONS_OK + MessageBoxButtons.DEFAULT_BUTTON_OK, title, message)
        messageBox.execute()
    }

    fun showMessageBox(title: String, message: String) {
        val messageBox = messageBoxFactory.createMessageBox(parentWindowPeer, MessageBoxType.INFOBOX, MessageBoxButtons.BUTTONS_OK + MessageBoxButtons.DEFAULT_BUTTON_OK, title, message)
        messageBox.execute()
    }

    fun showQueryBox(title: String, message: String, default: String = ""): Boolean {
        val messageBox = messageBoxFactory.createMessageBox(parentWindowPeer, MessageBoxType.QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL + MessageBoxButtons.DEFAULT_BUTTON_OK, title, message)
        val result = messageBox.execute()
        return result.toInt() == MessageBoxButtons.BUTTONS_OK
    }

    fun showInputBox(s: String, s1: String, s2: String): String {
        // https://wiki.openoffice.org/wiki/Python/Transfer_from_Basic_to_Python#Input_Box
        // http://api.libreoffice.org/examples/DevelopersGuide/GUI/UnoDialogSample.java
        // https://sourceforge.net/projects/libreoffice-java-inputbox/files/
        val dialog: UnoDialogSample = UnoDialogSample(componentContext, multiComponentFactory)
        try {
            dialog.initialize(arrayOf("Height", "Moveable", "Name", "PositionX", "PositionY", "Step", "TabIndex", "Title", "Width"),
                    arrayOf(Integer.valueOf(380), java.lang.Boolean.TRUE, "MyTestDialog", Integer.valueOf(102), Integer.valueOf(41), Integer.valueOf(0), java.lang.Short.valueOf(0.toShort()), "OpenOffice", Integer.valueOf(380)))
            val oFTHeaderModel = dialog.m_xMSFDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
            val xFTHeaderModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet::class.java, oFTHeaderModel)
            xFTHeaderModelMPSet.setPropertyValues(
                    arrayOf("Height", "Label", "Name", "PositionX", "PositionY", "Width"),
                    arrayOf(Integer.valueOf(8), "This code-sample demonstrates how to create various controls in a dialog", "HeaderLabel", Integer.valueOf(106), Integer.valueOf(6), Integer.valueOf(300)))
            // add the model to the NameContainer of the dialog model
            dialog.m_xDlgModelNameContainer.insertByName("Headerlabel", oFTHeaderModel)
            dialog.insertFixedText(dialog, 106, 18, 100, 0, "My ~Label")
            val edit = dialog.insertEditField(dialog, dialog, 106, 72, 60)
            dialog.insertButton(dialog, 106, 320, 50, "~OK", PushButtonType.OK_value.toShort())
            dialog.insertButton(dialog, 156, 320, 50, "~Cancel", PushButtonType.CANCEL_value.toShort())
//        oUnoDialogSample.createWindowPeer()
//        oUnoDialogSample.addRoadmap()
//        oUnoDialogSample.insertRoadmapItem(0, true, "Introduction", 1)
//        oUnoDialogSample.insertRoadmapItem(1, true, "Documents", 2)
            dialog.xDialog = UnoRuntime.queryInterface(XDialog::class.java, dialog.m_xDialogControl)
            if (dialog.executeDialog() != 0.toShort()) {
                return edit.text
            }
        } catch(e: Exception) {
            System.err.println(e.toString() + e.message)
            e.printStackTrace()
        } finally {
            //make sure always to dispose the component and free the memory!
            if (dialog.m_xComponent != null) {
                dialog.m_xComponent.dispose()
            }
        }

        return ""
//
//    val WIDTH = 600
//    val HORI_MARGIN = VERT_MARGIN = 8
//    val BUTTON_WIDTH = 100
//    val BUTTON_HEIGHT = 26
//    val HORI_SEP = VERT_SEP = 8
//    val LABEL_HEIGHT = BUTTON_HEIGHT * 2 + 5
//    val EDIT_HEIGHT = 24
//    val HEIGHT = VERT_MARGIN * 2 + LABEL_HEIGHT + VERT_SEP + EDIT_HEIGHT
////    import uno
////    from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE
////    from com.sun.star.awt.PushButtonType import OK, CANCEL
////    from com.sun.star.util.MeasureUnit import TWIP
////    ctx = uno.getComponentContext()
//    val create = fun (name: String): Any {
//        return ctx.getServiceManager().createInstanceWithContext(name, ctx)
//    }
//
//    val dialog = create("com.sun.star.awt.UnoControlDialog")
//    val dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
//     dialog.setModel(dialog_model)
//    dialog.setVisible(False)
//    dialog.setTitle(title)
//    dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)
//    fun add(name, type, x_, y_, width_, height_, props): Any {
//        model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + type + "Model")
//        dialog_model.insertByName(name, model)
//        control = dialog.getControl(name)
//        control.setPosSize(x_, y_, width_, height_, POSSIZE)
//        for key, value in props.items():
//        setattr(model, key, value)
//    }
//    val label_width = WIDTH - BUTTON_WIDTH - HORI_SEP - HORI_MARGIN * 2
//    add("label", "FixedText", HORI_MARGIN, VERT_MARGIN, label_width, LABEL_HEIGHT,
//    {"Label": str(message), "NoLabel": True})
//    add("btn_ok", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN,
//    BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})
//    add("btn_cancel", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN + BUTTON_HEIGHT + 5,
//    BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": CANCEL})
//    add("edit", "Edit", HORI_MARGIN, LABEL_HEIGHT + VERT_MARGIN + VERT_SEP,
//    WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(default)})
//    val frame = create("com.sun.star.frame.Desktop").getCurrentFrame()
//    val window = frame?.getContainerWindow()
//    dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
//    if (x != null && y != null) {
//        val ps = dialog.convertSizeToPixel(uno.createUnoStruct("com.sun.star.awt.Size", x, y), TWIP)
//        val _x, _y = ps.Width, ps.Height
//    } else if (window) {
//        ps = window.getPosSize()
//        _x = ps.Width / 2 - WIDTH / 2
//        _y = ps.Height / 2 - HEIGHT / 2
//    }
//        dialog.setPosSize(_x, _y, 0, 0, POS)
//        val edit = dialog.getControl("edit")
//    edit.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(default))))
//    edit.setFocus()
//    val ret = edit.getModel().Text if dialog.execute() else ""
//    dialog.dispose()
//    return ret
    }
}