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

import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;

import javax.swing.JApplet;
import javax.swing.JButton;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;


/**
 * ExcelFormula Applet
 * Converts an excel formula to a mathematical equation
 * 
 * @author Gerd Bartelt
 */
public class ExcelFormula extends JApplet implements ActionListener {
	
	private static final long serialVersionUID = -5281147244707576899L;

	// The main panel
	private ImagePanel mainPanel;
	
	// Text field with the excel formula
	private JTextField excelFormulaField;
	
	// Invisible text field with the LatTex formula
	private JTextField latexFormulaField;
	
	// Button to start the conversion of the excel formula
	private JButton excelButton;
	// Button to display the latex formula
	private JButton latexButton;
	
	// The converter
	private Excel2LaTex excel2LaTex =  new Excel2LaTex();
	
	// Control that displays the equation
	private Equation equation;

    
    /**
     * Create the GUI. For thread safety, this method should
     * be invoked from the event-dispatching thread.
     */
    private void createGUI() {

    	// Creates a main panel with a background image
    	mainPanel = new ImagePanel (this.getImage (this.getCodeBase(), "pics/xl.png"));
    	setContentPane(mainPanel); 

        //The Excel formula field
        excelFormulaField = new JTextField();
        excelFormulaField.setBorder(javax.swing.BorderFactory.createEmptyBorder());
        excelFormulaField.addActionListener (this);
        excelFormulaField.setText("= A_1^(1/3)+1/(1+x_2)+SUMME(A1:C6)+SIN(2*PI())+ABS(EXP(T/T_N))+WENN(x<10;0;10)");
        excelFormulaField.addKeyListener(new KeyListener(){

        	// Key listener for the ENTER key
			@Override
			public void keyPressed(KeyEvent arg0) {
				
				// ENTER key
				if (arg0.getKeyCode() == 0x0a)
				{
					// Do the calculation
					calc();
				}
			}

			@Override
			public void keyReleased(KeyEvent arg0) {
			}

			@Override
			public void keyTyped(KeyEvent arg0) {
			}
        });
        
        // Set the size and position 
        excelFormulaField.setBounds(200, 117, 650, 20);

        //The invisible LaTex formula field
        latexFormulaField = new JTextField();
        latexFormulaField.setBounds(5, 30, 400, 25);
        
        //The invisible button to start the conversion
        excelButton = new JButton (">>");
        excelButton.addActionListener (this);
        excelButton.setBounds(450, 5, 100, 25);
        
        // The invisible button to display the result
        latexButton = new JButton ("latex");
        latexButton.addActionListener (this);
        latexButton.setBounds(450, 30, 100, 25);

        // Create the equation control
        equation = new Equation(this); 
        equation.setLocation(54, 194);
        
        mainPanel.add(excelFormulaField);

        // Some invisible controls (only for debugging)
//        mainPanel.add(latexFormulaField);
//        mainPanel.add(excelButton);
//        mainPanel.add(latexButton);

        mainPanel.add(equation);
        mainPanel.setBackground (Color.white);

        // Do a first conversion of the demo formula
        calc();

    }

    /**
     * Do the conversion and display the result
     */
    private void calc() {

    	// Convert the excel formula
    	String result = excel2LaTex.convert(excelFormulaField.getText());
    	
    	// Display the result in a text field
    	latexFormulaField.setText(result);
    	
    	// Display the result as mathematical equation
    	equation.setFormula(result);
    	this.repaint();
    }
    
    
    /**
     * Called when this applet is loaded into the browser
     */
    public void init() {

    	// Execute a job on the event-dispatching thread:
        // Creating this applet's GUI.
        try {
            SwingUtilities.invokeAndWait(new Runnable() {
                public void run() {
                    createGUI();
                }
            });
        } catch (Exception e) { 
            System.err.println("createGUI didn't successfully complete");
        }

    }

    /**
     * Perform the action events
     */
    public void actionPerformed(ActionEvent e) {
    	String action;
    	action = e.getActionCommand ();
    	
    	// The button ">>" was pressed
	    if (action.equals(">>")) {
	    	String result = excel2LaTex.convert(excelFormulaField.getText());
	    	latexFormulaField.setText(result);
	    	equation.setFormula(result);
	    }
	    
	    // The button "latex" was pressed
	    if (action.equals("latex")) {
	    	equation.setFormula(latexFormulaField.getText());
	    }

    }

    /**
     * Start the applet
     */
    public void start() {
    }

    /**
     * Stop the applet
     */
    public void stop() {
    }
    
    /**
     * The applet information.
     */
    public String getAppletInfo() {
        return "Title: Excel equation converter v1.0, 2012-01-29\n"
               + "Author: Gerd Bartelt\n"
               + "Converts excel formulas to mathematical equations";
    }

}