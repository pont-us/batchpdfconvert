package net.talvi.batchpdfconvert;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.presentation.XPresentation;
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
    
    public static XComponent loadDocument(String urlString, XDesktop desktop) {
        try {
        com.sun.star.frame.XComponentLoader xCompLoader =
                UnoRuntime.queryInterface(
                 com.sun.star.frame.XComponentLoader.class, desktop);
        return xCompLoader.loadComponentFromURL(
                urlString, "_blank", 0, new com.sun.star.beans.PropertyValue[0]);
        } catch (IOException ex) {
            ex.printStackTrace(System.err);
            System.exit(1);
        }
        return null;
    }
    
    public static void loadAndConvertOdp() {
        final XDesktop desktop = getDesktop();
        final XComponent xComp = loadDocument("file:///home/pont/Untitled 1.odp", desktop);
        final XPresentation odpDoc = UnoRuntime.queryInterface(XPresentation.class, xComp);
        
        XDrawPagesSupplier xDPS = UnoRuntime.queryInterface(
                XDrawPagesSupplier.class, xComp);
        XDrawPages xDPn = xDPS.getDrawPages();
        com.sun.star.container.XIndexAccess xDPi =
                UnoRuntime.queryInterface(
                com.sun.star.container.XIndexAccess.class, xDPn);
        
        int nPages = xDPn.getCount();
        for (int i=0; i<nPages; i++) {
            try {
                XDrawPage page = UnoRuntime.queryInterface(
                com.sun.star.drawing.XDrawPage.class, xDPi.getByIndex(i));
                int nSomethings = page.getCount();
                XIndexAccess dpia = UnoRuntime.queryInterface(XIndexAccess.class, page);
                for (int j=0; j<nSomethings; j++) {
                    XShape shape = UnoRuntime.queryInterface(XShape.class, dpia.getByIndex(j));
                    System.out.println("Shape: "+shape.getShapeType());
                    // com.sun.star.presentation.TitleTextShape
                    // com.sun.star.presentation.SubtitleShape
                    //if ("com.sun.star.presentation.TitleTextShape".equals(shape.getShapeType())) {
                        XPropertySet shapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, shape);
                        shapeProps.setPropertyValue("Shadow", Boolean.FALSE);
                    //}
                }
                
            } catch (IndexOutOfBoundsException | WrappedTargetException |
                    UnknownPropertyException | PropertyVetoException | IllegalArgumentException ex) {
                ex.printStackTrace(System.err);
                System.exit(1);
            
            }
        }
        
        final XStorable storable = (XStorable)
	    UnoRuntime.queryInterface(XStorable.class, xComp);

	System.out.println("xStorable: " + storable);
 
	final String outputUrlString = "file:///home/pont/exported.pdf";
        
        final PropertyValue[] filterData = {
            //makePropVal("ReduceImageResolution", Boolean.TRUE),
            //makePropVal("MaxImageResolution", 75),
            makePropVal("UseLosslessCompression", Boolean.FALSE),
            makePropVal("Quality", 10)
           //     makePropVal("Watermark", "Wibble")
        };
        
        // See http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XStorable.html
	final PropertyValue[] saveOptions = {
            makePropVal("FilterName", "impress_pdf_Export"), 
            makePropVal("FilterData", filterData)
        };
        
	try {
	    storable.storeToURL(outputUrlString, saveOptions);
	} catch (com.sun.star.io.IOException ex) {
	    ex.printStackTrace(System.err);
	    return;
	}
    }
    
    public static void createAndConvertText() {
        	final XDesktop desktop = getDesktop();
        final XTextDocument textDoc = createTextDocument(desktop);
        
        textDoc.getText().setString("Hello world!");
        
        //final XComponent xComp = loadDocument("file:///home/pont/Untitled 1.odt", desktop);
        //final XTextDocument textDoc = UnoRuntime.queryInterface(XTextDocument.class, xComp);
 
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
    
    public static void main(String[] argv) {
        loadAndConvertOdp();
    }    
}
