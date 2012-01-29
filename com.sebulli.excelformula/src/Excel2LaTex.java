/*
 * 
 *  ExcelFormula
 *  Copyright (C) 2012  Gerd Bartelt
 *
 *  This program is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation, either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program.  If not, see <http://www.gnu.org/licenses/>.
 *   
 */

import java.util.ArrayList;

/**
 * Converts an excel formula to a latex one
 * 
 * @author Gerd Bartelt
 *
 */
public class Excel2LaTex {

	// Working String with the formula
	private String s;
	
	// Counts the start "«" and end tags "»"
	// The Java Editor MUST be set to UTF-8 !!
	private int tags = 0;

	/**
	 * Constructor
	 */
	public Excel2LaTex () {
	}

	/**
	 * Checks, whether the next character is part of the actual variable, or
	 * whether it is a new operator.
	 * 
	 * @param c
	 * 		The next character
	 * @param op
	 * 		The operator
	 * @return
	 * 		False, if it is still the actual variable name
	 */
	private boolean noVar(char c, String op) {

		// Count the start and end tags
		if (c=='«')
			tags++;
		if (c=='»')
			tags--;
		
		// Exit with false, if we are between a start and end tag
		if (tags != 0)
			return false;

		// Operators indicate the end of a variable
		switch (c) {
		case '+':
		case '-':
		case '/':
		case '*':
		case '=':
		//case ' ':
		case '{':
		case '}':
		case '\\':
				return true;

		// if the division operator is used, the "^" operator 
		// does not separate a block
		case '^':
				return (!op.equals("/"));
				
		default: 
				return false;
		}
	}
	
	/**
	 * Checks, whether a strings starts with a bracket and
	 * the bracket is not closed
	 * 
	 * @param s
	 * 		The string to check
	 * @return
	 * 		True, if it starts with an open bracket
	 */
	private boolean startsWithOpenBracket(String s) {

		// Exit, if the string does not start with a bracket
		if (!s.startsWith("("))
			return false;

		// Count the brackets
		int brackets = 0;
		
		for (int i = 0; i< s.length(); i++) {

			// Get the next character
			char c = s.charAt(i);
		
			// Count up on every "("
			if ( c == '(')
				brackets ++;
			
			// Count down on every ")"
			if ( c == ')') {
				brackets --;
				if (brackets <= 0) {
					return false;
				}
			}
		}
		
		// Return the result
		return true;
	}
	
	/**
	 * Checks, whether a strings ends with a bracket and
	 * the bracket is not closed
	 * 
	 * @param s
	 * 		The string to check
	 * @return
	 * 		True, if it ends with an open bracket
	 */
	private boolean endsWithOpenBracket(String s) {

		// Exit, if the string does not start with a bracket
		if (!s.endsWith(")"))
			return false;

		// Count the brackets
		int brackets = 0;

		for (int i = s.length()-1; i>=0; i--) {

			// Get the next character
			char c = s.charAt(i);

			// Count up on every ")"
			if ( c == ')')
				brackets ++;
			
			// Count down on every "("
			if ( c == '(') {
				brackets --;
				if (brackets <= 0) {
					return false;
				}
			}
		}

		// Return the result
		return true;
	}
	
	
	/**
	 * Remove unused brackets
	 * 
	 * @param s
	 * 		The string with brackets
	 * 
	 * @return
	 * 		The optimized string
	 */
	private String removeBrackets(String s) {

		// Count the open brackets
		int brackets = 0;
		boolean found = false;
		
		for (int i = 0; i< s.length(); i++) {

			// Get the next character
			char c = s.charAt(i);
			
			// Count up on every "("
			if ( c == '(') {
				brackets ++;
				found = true;
			}

			// Count down on every ")"
			if ( c == ')') {
				brackets --;
				found = true;
				
				// Exit, if more brackets are closes than opened
				if (brackets < 0) {
					return s;
				}
			}
		}

		// Exit with the original string, if more brackets are opened
		// than closed
		if (brackets != 0)
			return s;
		
		// Exit with the original string, if no brackets are found
		if (!found)
			return s;
		
		// Remove the leading and trailing brackets 
		String st = s.trim();
		if (st.startsWith("(") && st.endsWith(")"))
			return st.substring(1,st.length()-1);
		
		// Exit with the original string
		return s;
	}
	
	/**
	 * Converts an operator to a latex operator
	 * 
	 * @param op
	 * 		The operator to replace
	 * @param latexOp
	 * 		The latex operator
	 * @return
	 * 		True, if at least one operator was replaced
	 */
	private boolean convertOperator(String op, String latexOp) {
		
		// Position of the operator
		int pos = s.indexOf(op);

		// Not found, exit with false
		if (pos < 0)
			return false;
		
		// Count the brackets
		int brackets;
		
		// String before the operator
		String pre ="";
		brackets = 0; 
		
		// Reset the open tag counter
		tags = 0;
		
		// Search all characters before the operator
		for (int i = pos-1; i>= 0; i--) {
			
			// Get the next character
			char c = s.charAt(i);
			
			// Count up on every ")"
			if ( c == ')')
				brackets ++;

			// Count down on every "("
			if ( c == '(') {
				brackets --;
			}
			
			// If all open brackets are closed, check for a character
			// That is not part of the variable oder block
			if (brackets == 0) {
				if (noVar(c, op)) {
					break;
				}
			}
			
			// Exit, if more brackets are closed than opened
			if (brackets < 0) {
					break;
			}

			// Collect the characters to a string
			pre = c + pre; 
				
		}
		
		
		// Reset the bracket counter
		brackets = 0; 

		// String after the operator
		String post ="";
		tags = 0;

		// Search all characters after the operator
		for (int i = pos+1; i < s.length(); i++) {

			// Get the next character
			char c = s.charAt(i);

			// Count up on every "("
			if ( c == '(')
				brackets ++;

			// Count down on every ")"
			if ( c == ')') {
				brackets --;
			}
			
			// If all open brackets are closed, check for a character
			// That is not part of the variable oder block
			if (brackets == 0) {
				if (noVar(c,op)) {
					break;
				}
			}

			// Exit, if more brackets are closed than opened
			if (brackets < 0) {
					break;
			}
			
			// Collect the characters to a string
			post = post + c;
		}

		// Generate the string to replace
		String target = pre + op + post;

		// Remove the leading and trailing brackets
		if (startsWithOpenBracket(pre) && endsWithOpenBracket(post)) {
			pre = pre.substring(1, pre.length());
			post = post.substring(0, post.length()-1);
		}

		// Remove the leading bracket
		if (startsWithOpenBracket(pre)) {
			pre = pre.substring(1, pre.length());
			target = target.substring(1, target.length());
		}

		// Remove the trailing bracket
		if (endsWithOpenBracket(post)) {
			post = post.substring(0, post.length()-1);
			target = target.substring(0, target.length()-1);
		}
		
		// Remove leading and trailing brackets
		pre = removeBrackets(pre);
		post = removeBrackets(post);

		// Generate the replacement
		String replacement = "«" + latexOp + " {" + pre+ "} {" + post + "}»";
		
		// Convert ^(1/x) to square roots
		if (op.equals("^")) {
			
			// Search for the "1/"
			if (post.startsWith("1/")) {
				
				// Replace it by a square root
				s= s.replace(target, "$$tmp$$");
				target = "$$tmp$$";
				replacement = "« \\sqrt [" + post.substring(2) + "]{" + pre+ "}»";
			}
			else
				return false;	
		}

		// Replace it
		s= s.replace(target, replacement);
		return true;
	}
	
	/**
	 * Convert an excel function to a latex one
	 *  
	 * @param op
	 * 		The Excel function name
	 * @param latexOp
	 * 		The name of the latex function
	 * @param keepBrackets
	 * 		false, if unnecessary brackets are removed 
	 * @return
	 * 		True, if at least one function was replaced
	 */
	private boolean convertFunction(String op, String latexOp, boolean keepBrackets) {

		// Get the start position of the function name
		int pos = s.indexOf(op+"(");

		// Get the end position of the function name
		int pos2 = s.indexOf(op+"(") + op.length()+1;

		// Not found, exit with false
		if (pos < 0)
			return false;
		
		// Count the open brackets
		int brackets = 1;
		
		// Collect all parameters
		ArrayList<String> params = new ArrayList<String>();
		params.add("");
		
		// String with all parameters
		String param = "";
		int parami=0;
		
		// Search all characters after the funtion name
		for (int i = pos2; i < s.length(); i++) {
			
			// Get the next character
			char c = s.charAt(i);

			// Count up on every "("
			if ( c == '(')
				brackets ++;

			// Count down on every ")"
			if ( c == ')') {
				brackets --;

				// If all brackets are closed  exit the loop
				if (brackets <= 0) {
					break;
				}
			}
			
			// split the parameters
			if ((brackets == 1) && (c==';')) {

				// Count the function parameters
				parami++;
				params.add("");
			}
			else {
				// Set the parameter name
				params.set(parami, params.get(parami) + c); 
			}
			
			// Generate also the parameter string with all parameters
			param = param + c;
				
		}
		
		// Generate the text to replace
		String target = op+"(" +  param + ")";
		
		// Add brackets
		if (keepBrackets)
			param = "(" + param +")";

		// Generate the replacement with start and end tags
		String replacement = "«" + latexOp + " {" + param+ "}»";

		// The ABS function is replaced by a | operator
		if (latexOp.equals("|"))
			replacement = "«|" + param+ "|»";
		
		// Convert logical operations AND and OR
		if (op.equals("AND") || op.equals("OR")) {
			replacement = "«(";
			for (String par : params) {
				replacement += par+ " " + latexOp + " ";
			}
			replacement = replacement.substring(0, replacement.length()-latexOp.length()-1);
			replacement += ")»";
			
		}
		
		// Convert the IF function to a latex case
		if (op.equals("IF") && (params.size()==3)) {
			replacement = "« \\begin{cases}";
			replacement += params.get(1)+ " & \\text { if } ";
			replacement += params.get(0) + ",\\\\";
			
			replacement += params.get(2)+ " & \\text { other cases }";
			replacement += "\\end{cases}»";
		}

		// Convert the SUM function
		if (op.equals("SUM") ) {
			String[] sumparams = param.split(":");
			if (sumparams.length == 2) {
				replacement = "« \\sum_ ";
				replacement += "{" + sumparams[0] + "}";
				replacement += "^";
				replacement += "{" + sumparams[1] + "}";
//				replacement += " {\\dots}";
				replacement += " {}";
				replacement += " »";
			}
		}

		// Convert the EXP funtion
		if (op.equals("EXP") ) {
			replacement = "« e^";
			replacement += "{" + param + "}";
			replacement += " »";
		}

		// Replace the text
		s = s.replace(target, replacement);
		return true;
	}
	

	/**
	 * Generate a prefix to enlarge brackets
	 * 
	 * @param size
	 * 		The size from 0 (normal) to 5 (largest)
	 * @return
	 * 		The latex tag
	 */
	private String getPraefix(int size) {

		// Return the latex tag depending on the size
		switch (size) {
		case 1: return "";
		case 2: return "\\big";
		case 3: return "\\Big";
		case 4: return "\\bigg";
		case 5: return "\\Bigg";
		}
		return "\\Bigg ";
	}
	
	/**
	 * Adds prefix \\big ... to a bracket, depending on its position
	 * 
	 * @param s
	 * 		The string part to convert
	 * @param maxdeep
	 * 		The maximum deep of the brackets
	 * @return
	 * 		The converted string
	 */
	private String formatBracketsPart(String s, int maxdeep) {
		
		int i;
		char c;
		
		// The deep
		int deep = maxdeep;
		
		// Temporary string
		String part ="";

		// Get all characters
		for (i=0; i<s.length();i++) {
			
			// Get the next character
			c = s.charAt(i);
			
			// Add the prefix before opening brackets
			if (c == '(') {
				part += getPraefix(deep) + c ;
				deep--;
			}

			// Add the prefix before closing brackets
			else if (c == ')') {
				deep++;
				part += getPraefix(deep) + c ;
			}
			
			// Add no prefix
			else {
				part += c;
			}
			
		}
		
		// Return the converted string
		return part;
	}
	
	
	/**
	 * Search all brackets, split the string into blocks and
	 * count the maximum bracket deep of this block
	 *  
	 * @param s
	 * 		The string to convert
	 * @return
	 * 		The converted string
	 */
	private String formatBrackets(String s) {
		
		int i;
		char c;
		
		// Count the bracket deep
		int deep = 0;
		
		// Part of the string
		String part ="";
		
		// The maximum deep
		int maxdeep = 0;
		
		// The converted string
		String snew = "";
		
		// Get all characters of the string
		for (i=0; i<s.length();i++) {
			
			// Get the next character
			c = s.charAt(i);
			part += c;
			
			// If an open bracket is detected, increase the bracket deep
			if (c == '(') {
				deep++;
				
				// Check the maximum deep
				if (deep > maxdeep) {
					maxdeep = deep;
				}
			}

			// If an closing bracket is detected, decrease the bracket deep
			if (c == ')') {
				deep--;
				
				// If all brackets are closed, than format this part
				if (deep ==0 ) {
					snew += formatBracketsPart(part,maxdeep);

					// Reset maximum deep and part string for the next block
					maxdeep = 0;
					part = "";
				}
			}
			
		}
		
		// Return the converted string
		return snew + part;
	}
	
	/**
	 * Convert an excel string to a latex string 
	 * 
	 * @param excelString
	 * 		The string to convert
	 * @return
	 * 		The converted string
	 */
	public String convert (String excelString) {

		s = excelString;

		// Some tests
		//s = "1/ (3+4)";
		// s = "(1*1/2+4^2/4^3)^2";
		// s = "(1/2)^2";
		// s = "A_1+3^(9-1)+SUMME(A1:I3)";
		// s = "SUMME(A1:B12)+12+WURZEL(199)+EXP(1/T)+(1+(1+(1)))";
		// s = "A_1^(1/3)+1/(1+x_2)+SUMME(A1:C6)+SIN(2*PI())+ABS(EXP(T/T_N))+WENN(x<10;0;10)";
		// s = "(SUMME(A1:B2))^(1/3)";

		
		// Convert German function names
		s= s.replace("ADRESSE(", "ADDRESS(");
		s= s.replace("INDIREKT(", "INDIRECT(");
		s= s.replace("WURZEL(", "SQRT(");
		s= s.replace("SUMME(", "SUM(");
		s= s.replace("WENN(", "IF(");
		s= s.replace("UND(", "AND(");
		s= s.replace("ODER(", "OR(");
		
		// Convert the ABS function to an ABS operator
		while (convertFunction("ABS", "|", false));

		// Convert logical functions
		while (convertFunction("AND", "\\wedge", false));
		while (convertFunction("OR", "\\vee", false));
		while (convertFunction("IF", "", false));

		// Convert functions
		while (convertFunction("SQRT", "\\sqrt", false));
		while (convertFunction("EXP", "", false));
		while (convertFunction("SIN", "\\sin", true));
		while (convertFunction("COS", "\\cos", true));
		while (convertFunction("TAN", "\\tan", true));
		while (convertFunction("SINH", "\\sinh", true));
		while (convertFunction("COSH", "\\cosh", true));
		while (convertFunction("TANH", "\\tanh", true));
		while (convertFunction("ARCSIN", "\\arcsin", true));
		while (convertFunction("ARCCOS", "\\arccos", true));
		while (convertFunction("ARCTAN", "\\arctan", true));

		// Convert functions
		while (convertFunction("LN", "\\ln", true));
		while (convertFunction("LG", "\\lg", true));
		while (convertFunction("LOG", "\\log", true));

		// Convert functions
		while (convertFunction("MIN", "\\min", true));
		while (convertFunction("MAX", "\\max", true));

		// Convert operators
		while (convertOperator("^", "\\sqrt"));
		while (convertFunction("^", "^",  false));

		// Convert SUM function
		while (convertFunction("SUM", "\\sum", false));

		// Convert operators
		while (convertOperator("/", "\\frac"));
		
		// Replace some special characters and texts
		s = s.replace(":", " \\dots ");
		s = s.replace("*PI()", " \\pi ");
		s = s.replace("PI()", " \\pi ");
		s = s.replace("*", " \\cdot ");

		// Format the brackets
		s = formatBrackets(s);
		
		// Some funtions that are not supported
		if (s.contains("INDIRECT(") || 
			s.contains("ADDRESS(") )
			s= "\\text{willst mich testen ?? }";

		// Remove the start and end tags
		s = s.replace('«',' ');
    	s = s.replace('»',' ');

		return s;
	}
}


