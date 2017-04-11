package net.talvi.batchpdfconvert;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
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
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.XCloseable;
import static com.sun.star.uno.UnoRuntime.queryInterface;
import com.sun.star.util.CloseVetoException;

import java.io.File;
import java.util.logging.Level;
import java.util.logging.Logger;

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
                return queryInterface(XDesktop.class, desktop);
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

    protected static XComponent createDocument(XDesktop desktop,
            String docType) {
        final String urlString = "private:factory/" + docType;
        final PropertyValue emptyArgs[] = new PropertyValue[0];
        try {
            final XComponentLoader xComponentLoader =
                    queryInterface(XComponentLoader.class, desktop);   
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
        return queryInterface(XTextDocument.class, component);
    }
    
    public static XComponent loadDocument(String urlString, XDesktop desktop) {
        try {
            final XComponentLoader xCompLoader =
                    queryInterface(XComponentLoader.class, desktop);
            return xCompLoader.loadComponentFromURL(
                    urlString, "_blank", 0, new PropertyValue[0]);
        } catch (IOException ex) {
            ex.printStackTrace(System.err);
            System.exit(1);
        }
        return null;
    }
    
    public static void closeDocument(XComponent component) {
        // Check document functionality
        final XModel model = (XModel) queryInterface(XModel.class, component);
        
        System.out.println("Model: "+model);
        
        if (model != null) {
            final XComponent disposable = (XComponent) queryInterface(
                        XComponent.class, model);
            // Proper document so close if possible.
            final XCloseable xCloseable =
                    (XCloseable)queryInterface(
                            XCloseable.class,model);
            
            if (xCloseable != null) {
                System.out.println("Closeable.");
                try {
                    // https://www.openoffice.org/api/docs/common/ref/com/sun/star/util/XCloseable.html
                    // https://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Closing_Documents

                    // Entirely correct behaviour would be a little complicated 
                    // to implement here. Immediate close can't be guaranteed:
                    // something may throw a closeVeto indicating that it needs
                    // to finish something before close (and that something may
                    // take an unbounded time). The flag we pass to the close
                    // method controls who takes ownership of the object, but
                    // that doesn't affect our responsibility to wait until 
                    // nobody is vetoing the close -- if we have ownership we
                    // just get the additional responsibility to "try again at
                    // a later time" to do the close.
                    
                    // Here we give up ownership, which means that if anyone
                    // throws a veto exception, they take on the responsibility
                    // for closing the document. In that case, we *should*
                    // probably wait nicely until they close it. For now, we
                    // just trust that it will happen -- a veto seems very
                    // unlikely anyway in a conversion process where we're
                    // just loading, saving, and closing a document.
                    
                    xCloseable.close(true);
                } catch(CloseVetoException closeVeto) {
                    System.out.println("Veto!");
                    // Attempting to force things with a dispose here would 
                    // probably be unwise: the wiki says " Calling dispose() on
                    // an XCloseable might lead to deadlocks or crash the entire
                    // application."
                }
            } else { // If we can't close, we dispose.
                System.out.println("Disposing.");
                disposable.dispose();
            }
        }
    }
    
    public static void loadAndConvertOdp(String inputPath, String outputPath) {
        final XDesktop desktop = getDesktop();
        String inputUrl = null;
        String outputUrl = null;
        try {
            inputUrl = (new File(inputPath)).toURI().toURL().toString();
            outputUrl = (new File(outputPath)).toURI().toURL().toString();
        } catch (java.io.IOException ex) {
            Logger.getLogger(BatchPdfConvert.class.getName()).log(Level.SEVERE, null, ex);
            System.exit(1);
        }
        final XComponent xComp = loadDocument(inputUrl, desktop);
        final XPresentation odpDoc = queryInterface(XPresentation.class, xComp);
        
        final XDrawPagesSupplier xDPS = queryInterface(XDrawPagesSupplier.class, xComp);
        final XDrawPages xDPn = xDPS.getDrawPages();
        final XIndexAccess xDPi = queryInterface(XIndexAccess.class, xDPn);
        
        for (int i=0; i<xDPn.getCount(); i++) {
            try {
                final XDrawPage page = queryInterface(XDrawPage.class, xDPi.getByIndex(i));
                final XIndexAccess dpia = queryInterface(XIndexAccess.class, page);
                for (int j=0; j<page.getCount(); j++) {
                    XShape shape = queryInterface(XShape.class, dpia.getByIndex(j));
                    // System.out.println("Shape: "+shape.getShapeType());
                    // com.sun.star.presentation.TitleTextShape
                    // com.sun.star.presentation.SubtitleShape
                    //if ("com.sun.star.presentation.TitleTextShape".equals(shape.getShapeType())) {
                    final XPropertySet shapeProps = (XPropertySet)
                            queryInterface(XPropertySet.class, shape);
                    final XPropertySetInfo psi = shapeProps.getPropertySetInfo();
                    if (psi.hasPropertyByName("Shadow")) {
                        shapeProps.setPropertyValue("Shadow", Boolean.FALSE);
                    }
                    //}
                }
                
            } catch (IndexOutOfBoundsException | WrappedTargetException |
                    UnknownPropertyException | PropertyVetoException | IllegalArgumentException ex) {
                ex.printStackTrace(System.err);
                System.exit(1);
            }
        }
        
        final XStorable storable = (XStorable)
                queryInterface(XStorable.class, xComp);

	System.out.println("xStorable: " + storable);
         
        final PropertyValue[] filterData = {
            // Image resolution settings don't seem to affect file size
            // in practice.
            // makePropVal("ReduceImageResolution", Boolean.TRUE),
            // makePropVal("MaxImageResolution", 75),
            makePropVal("UseLosslessCompression", Boolean.FALSE),
            makePropVal("Quality", 80)
        };
        
        // See http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XStorable.html
	final PropertyValue[] saveOptions = {
            makePropVal("FilterName", "impress_pdf_Export"), 
            makePropVal("FilterData", filterData)
        };
        
	try {
	    storable.storeToURL(outputUrl, saveOptions);
	} catch (com.sun.star.io.IOException ex) {
	    ex.printStackTrace(System.err);
	}
        
        closeDocument(xComp);
        System.out.println("Done.");
    }
    
    public static void createAndConvertText() {
        final XDesktop desktop = getDesktop();
        final XTextDocument textDoc = createTextDocument(desktop);
        
        textDoc.getText().setString("Hello world!");
        
        //final XComponent xComp = loadDocument("file:///home/pont/Untitled 1.odt", desktop);
        //final XTextDocument textDoc = queryInterface(XTextDocument.class, xComp);
 
	final XStorable storable = (XStorable)
	    queryInterface(XStorable.class, textDoc);

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
        loadAndConvertOdp(argv[0], argv[1]);
        System.exit(0);
    }    
}
