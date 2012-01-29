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

import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;

import javax.swing.JApplet;
import javax.swing.JComponent;

import org.scilab.forge.jlatexmath.TeXConstants;
import org.scilab.forge.jlatexmath.TeXFormula;
import org.scilab.forge.jlatexmath.TeXIcon;

/**
 * Equation control to display the result as a LaTex equation
 * 
 * @author Gerd Bartelt
 *
 */
public class Equation extends JComponent {

	private static final long serialVersionUID = 2039288774668284129L;
	
	// The excel formula
	private String math = "";
	
	// The converted formula
	TeXFormula fomule;
	
	// The icon with the equation
	TeXIcon ti;
	
	// Reference to the applet
	JApplet app;

	/**
	 * Constructor 
	 * creates the control
	 * 
	 * @param app
	 * 		The applet
	 */
    public Equation(JApplet app) {
        super ();
        this.app = app;
        this.setSize(752, 416);
    }
    
    /**
     * Converts a LaTex formula to an icon
     * 
     * @param formula
     * 		The LaTex formula
     */
    public void setFormula (String formula) {
    	math = formula;
    	
    	// Create the icon with a size of 25
        fomule = new TeXFormula(math);
        ti = fomule.createTeXIcon(TeXConstants.STYLE_DISPLAY, 25);
    }
    
    /**
     * Paint the component
     * 
     * @param g
     * 		Reference to the graphics
     */
    public void paintComponent(Graphics g) {
        super.paintComponent(g);
        Graphics2D g2D = (Graphics2D)g;

        
        // Do it only, if the icon is valid
        if (ti != null) {
            try {
            	// Use a image buffer
                BufferedImage image = new BufferedImage(752, 416, BufferedImage.TYPE_INT_ARGB);
                Graphics2D g2 = image.createGraphics();
                
                // Paint the equation
                ti.paintIcon(this, g2, 0, 0);
                g2D.drawImage(image,0,0 , app);

                g2.dispose();
            } catch (Exception e) {
            }
        }

        g2D.dispose();
        
    }

}

