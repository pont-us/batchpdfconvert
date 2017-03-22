package net.talvi.batchpdfconvert;

import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 *
 * @author pont
 */
public class BatchPdfConvert {

    public static XDesktop getDesktop() {
        try {
            final XComponentContext context = Bootstrap.bootstrap();
            final XMultiComponentFactory mcf = context.getServiceManager();
            if (mcf != null) {
                System.out.println("Connected to LibreOffice.");

                final Object desktop = mcf.createInstanceWithContext(
                        "com.sun.star.frame.Desktop", context);
                return UnoRuntime.queryInterface(XDesktop.class, desktop);
            } else {
                System.out.println("Connection failed!");
                System.exit(1);
            }
        } catch (BootstrapException | Exception e) {
            e.printStackTrace(System.err);
            System.exit(1);
        }
        System.exit(1);
        return null;
    }
    
    private static PropertyValue makePropVal(String name, Object value) {
        final PropertyValue pv = new PropertyValue();
        pv.Name = name;
        pv.Value = value;
        return pv;
    }

    protected static XComponent createDocument(XDesktop xDesktop,
            String docType) {
        final String urlString = "private:factory/" + docType;
        final PropertyValue emptyArgs[] = new PropertyValue[0];
        try {
            final XComponentLoader xComponentLoader =
                    UnoRuntime.queryInterface(XComponentLoader.class, xDesktop);   
            return xComponentLoader.loadComponentFromURL(
                    urlString, "_blank", 0, emptyArgs);
        } catch(com.sun.star.io.IOException ex) {
            ex.printStackTrace(System.err);
            System.exit(1);
        }
        return null;
    }
    
    public static XTextDocument createTextDocument(XDesktop desktop) {
        final XComponent component = createDocument(desktop, "swriter");
        return UnoRuntime.queryInterface(XTextDocument.class, component);
    }
    
    public static void main(String[] argv) {
	final XDesktop desktop = getDesktop();
        final XTextDocument textDoc = createTextDocument(desktop);
        
        textDoc.getText().setString("Hello world!");
 
	final XStorable storable = (XStorable)
	    UnoRuntime.queryInterface(XStorable.class, textDoc);

	System.out.println("xStorable: " + storable);
 
	final String outputUrlString = "file:///home/pont/exported.pdf";
        
        final PropertyValue[] filterData = {
                makePropVal("Watermark", "Wibble")
        };
        
        // See http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XStorable.html
	final PropertyValue[] saveOptions = {
            makePropVal("FilterName", "writer_pdf_Export"),  // or impress_pdf_Export, etc.
            makePropVal("FilterData", filterData)
        };
        
	try {
	    storable.storeToURL(outputUrlString, saveOptions);
	} catch (com.sun.star.io.IOException ex) {
	    ex.printStackTrace(System.err);
	    return;
	}
    }    
}
