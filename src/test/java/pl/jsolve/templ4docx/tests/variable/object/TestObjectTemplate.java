package pl.jsolve.templ4docx.tests.variable.object;

import static org.junit.Assert.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import org.junit.Test;

import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.core.VariablePattern;
import pl.jsolve.templ4docx.tests.variable.object.model.Obj01;
import pl.jsolve.templ4docx.tests.variable.object.model.Obj02;
import pl.jsolve.templ4docx.tests.variable.object.model.Obj03;
import pl.jsolve.templ4docx.variable.ObjectVariable;
import pl.jsolve.templ4docx.variable.TextVariable;
import pl.jsolve.templ4docx.variable.Variables;

/**
 * @author indvd00m (gotoindvdum[at]gmail[dot]com)
 *
 */
public class TestObjectTemplate extends AbstractVariableObjectTest {

    @Test
    public void test() throws IOException {
        String templateFileName = "variable/object/object-template";
        InputStream is = getClass().getClassLoader().getResourceAsStream(templateFileName + ".docx");

        Docx docx = new Docx(is);
        is.close();
        docx.setVariablePattern(new VariablePattern("#{", "}"));

        Variables var = new Variables();
        var.addObjectVariable(new ObjectVariable("#{var01}", new Obj01(), docx.getVariablePattern()));
        var.addObjectVariable(new ObjectVariable("#{var02}", new Obj02(), docx.getVariablePattern()));
        var.addObjectVariable(new ObjectVariable("#{var03}", new Obj03(), docx.getVariablePattern()));

        var.addTextVariable(new TextVariable("#{header}", "Injected header text"));
        var.addTextVariable(new TextVariable("#{footer}", "Injected footer text"));
        
        docx.fillTemplate(var);

        String tmpPath = System.getProperty("java.io.tmpdir");
        String processedPath = String.format("%s%s%s", tmpPath, File.separator,
                templateFileName + "-processed" + ".docx");

        docx.save(processedPath);

        String text = docx.readTextContent();
        String expected =        		
        		"Header text here - Injected header text – what happens to this text.\n"
        		+ "This is test simple template with three variables: value01, value02, value03."
        		+ " This is nested values of variables: field1Value, field2Value, field3Value. "
        		+ "And more: field11Value.\nFooter text - "
        		+ "Injected footer text";
       
        assertEquals(
                expected,
                text.trim());
    }

}
