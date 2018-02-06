package pl.jsolve.templ4docx.utils;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.core.VariablePattern;

/**
 * Utility class for various XWPF related objects.
 * 
 * @author Michael A. Angelo
 * @since 2018-02-06
 *
 */
public class XWPFUtils {

	/**
	 *
     * XWPFParagraph.runs are erroneously getting create with variable pattern prefix, key-name, and then suffix
     * all in separate ArrayList positions.
     * 
	 * These individual entries need to be concatenated and set onto 1 position with the paragraph runs.
	 * The two run entries containing only the prefix and suffix should be removed.
	 * 
	 * Note: This only seems to happen within the headers and footers.
	 * 
	 * Example:
	 * 
	 * Original runs = ['some text', '#{', 'token_key', '}', 'some other text']
	 * 
	 * New runs = ['some text', '#{token_key}', 'some other text']
	 * 
	 * There are other scenarios currently not covered:
	 * 
	 * ['#', '{', 'token_key', '}', 'other text']
	 * ['#{', 'token_key}', 'other text']
	 * 
	 * @param doc
	 */
	private static void fixRuns(List<XWPFParagraph> paragraphs, String prefix, String suffix) {
		
        for (XWPFParagraph paragraph : paragraphs) {
        	List<XWPFRun> toRemove = new ArrayList<XWPFRun>();
        	List<XWPFRun> runs = paragraph.getRuns();
        	for (int i = 0; i < runs.size(); i++) {
        		XWPFRun run = runs.get(i);
        		
        		// If the run text is only the prefix, then we have this situation:
        		// run[ndx] = [prefix]
        		// run[ndx + 1] = token
        		// run[ndx + 2] = [suffix]
        		
        		if (run.text().equals(prefix)) {
        			// The next run is the token within the prefix and suffix.
        			XWPFRun nextRun = runs.get(i + 1);
        			
        			// Put the prefix + token + suffix all on the same line.
        			String nextText = prefix + nextRun.text() + suffix;
        			nextRun.setText(nextText, 0);
        			
        			// Remove the runs with only the prefix and the suffix.
        			XWPFRun currentPlus2 = runs.get(i + 2);
        			toRemove.add(run);
        			toRemove.add(currentPlus2);        			 
        		}
        	}
        	
            // Remove all runs that were marked for removal.
        	for (XWPFRun runToRemove : toRemove) {
        		int runNdxToRemove = runs.indexOf(runToRemove);        		
        		paragraph.removeRun(runNdxToRemove);
        	}
        	
        }
	}

	public static void fixDocumentRuns(Docx doc) {
		fixHeaderRuns(doc);
		fixParagraphRuns(doc);
		fixFooterRuns(doc);
	}
	
	public static void fixHeaderRuns(Docx doc) {
		XWPFDocument document = doc.getXWPFDocument();
		VariablePattern vp = doc.getVariablePattern();
	
		for (XWPFHeader header : document.getHeaderList()) {
			fixRuns(header.getParagraphs(), vp.getOriginalPrefix(), vp.getOriginalSuffix()); 
		}
	}

	public static void fixParagraphRuns(Docx doc) {
		XWPFDocument document = doc.getXWPFDocument();
		VariablePattern vp = doc.getVariablePattern();
	
		fixRuns(document.getParagraphs(), vp.getOriginalPrefix(), vp.getOriginalSuffix());		
	}
	
	public static void fixFooterRuns(Docx doc) {
		XWPFDocument document = doc.getXWPFDocument();
		VariablePattern vp = doc.getVariablePattern();
	
		for (XWPFFooter footer : document.getFooterList()) {
			fixRuns(footer.getParagraphs(), vp.getOriginalPrefix(), vp.getOriginalSuffix());
		}		
	}
}
