import com.sun.star.awt.PushButtonType;
import com.sun.star.awt.XDialog;
import com.sun.star.beans.XMultiPropertySet;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.Time;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class FirstJUnit5Test {
    @Test
    void myFirstTest() {
        assertEquals(2, 1 + 1);
    }

    //    @Test
    void UnoDialogSampleMain() {
        UnoDialogSample oUnoDialogSample = null;

        try {
            XComponentContext xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();
            if (xContext != null)
                System.out.println("Connected to a running office ...");
            XMultiComponentFactory xMCF = xContext.getServiceManager();
            oUnoDialogSample = new UnoDialogSample(xContext, xMCF);
            oUnoDialogSample.initialize(new String[]{"Height", "Moveable", "Name", "PositionX", "PositionY", "Step", "TabIndex", "Title", "Width"},
                    new Object[]{Integer.valueOf(380), Boolean.TRUE, "MyTestDialog", Integer.valueOf(102), Integer.valueOf(41), Integer.valueOf(0), Short.valueOf((short) 0), "OpenOffice", Integer.valueOf(380)});
            Object oFTHeaderModel = oUnoDialogSample.m_xMSFDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel");
            XMultiPropertySet xFTHeaderModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oFTHeaderModel);
            xFTHeaderModelMPSet.setPropertyValues(
                    new String[]{"Height", "Label", "Name", "PositionX", "PositionY", "Width"},
                    new Object[]{Integer.valueOf(8), "This code-sample demonstrates how to create various controls in a dialog", "HeaderLabel", Integer.valueOf(106), Integer.valueOf(6), Integer.valueOf(300)});
            // add the model to the NameContainer of the dialog model
            oUnoDialogSample.m_xDlgModelNameContainer.insertByName("Headerlabel", oFTHeaderModel);
            oUnoDialogSample.insertFixedText(oUnoDialogSample, 106, 18, 100, 0, "My ~Label");
            oUnoDialogSample.insertCurrencyField(oUnoDialogSample, 106, 30, 60);
            oUnoDialogSample.insertProgressBar(106, 44, 100, 100);
            oUnoDialogSample.insertHorizontalFixedLine(106, 58, 100, "My FixedLine");
            oUnoDialogSample.insertEditField(oUnoDialogSample, oUnoDialogSample, 106, 72, 60);
            oUnoDialogSample.insertTimeField(106, 96, 50, new Time(0, (short) 0, (short) 0, (short) 10, false), new Time((short) 0, (short) 0, (short) 0, (short) 0, false), new Time((short) 0, (short) 0, (short) 0, (short) 17, false));
            oUnoDialogSample.insertDateField(oUnoDialogSample, 166, 96, 50);
            oUnoDialogSample.insertGroupBox(102, 124, 70, 100);
            oUnoDialogSample.insertPatternField(106, 136, 50);
            oUnoDialogSample.insertNumericField(106, 152, 50, 0.0, 1000.0, 500.0, 100.0, (short) 1);
            oUnoDialogSample.insertCheckBox(oUnoDialogSample, 106, 168, 150);
            oUnoDialogSample.insertRadioButtonGroup((short) 50, 130, 200, 150);
            oUnoDialogSample.insertListBox(106, 230, 50, 0, new String[]{"First Item", "Second Item"});
            oUnoDialogSample.insertComboBox(106, 250, 50);
            oUnoDialogSample.insertFormattedField(oUnoDialogSample, 106, 270, 100);
            oUnoDialogSample.insertVerticalScrollBar(oUnoDialogSample, 230, 230, 52);
            oUnoDialogSample.insertFileControl(oUnoDialogSample, 106, 290, 200);
            oUnoDialogSample.insertButton(oUnoDialogSample, 106, 320, 50, "~Close dialog", (short) PushButtonType.OK_value);
            oUnoDialogSample.createWindowPeer();
//            oUnoDialogSample.addRoadmap();
//            oUnoDialogSample.insertRoadmapItem(0, true, "Introduction", 1);
//            oUnoDialogSample.insertRoadmapItem(1, true, "Documents", 2);
            oUnoDialogSample.xDialog = UnoRuntime.queryInterface(XDialog.class, oUnoDialogSample.m_xDialogControl);
            oUnoDialogSample.executeDialog();
        } catch (Exception e) {
            System.err.println(e + e.getMessage());
            e.printStackTrace();
        } finally {
            //make sure always to dispose the component and free the memory!
            if (oUnoDialogSample != null) {
                if (oUnoDialogSample.m_xComponent != null) {
                    oUnoDialogSample.m_xComponent.dispose();
                }
            }
        }
    }
}