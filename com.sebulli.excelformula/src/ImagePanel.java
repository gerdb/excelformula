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

import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Image;

import javax.swing.JPanel;

/**
 * Creates a JPanel with a background image
 * 
 * @author Gerd Bartelt
 *
 */
public class ImagePanel extends JPanel{

	private static final long serialVersionUID = 1003478673077301230L;
	
	// The image
	private Image img;


	/**
	 * Constructor
	 * Creates an ImagePanel from an image
	 * @param img
	 */
	public ImagePanel(Image img) {
		this.img = img;
		Dimension size = new Dimension(img.getWidth(null), img.getHeight(null));
		setPreferredSize(size);
		setMinimumSize(size);
		setMaximumSize(size);
		setSize(size);
		setLayout(null);
	}

	/**
	 * Paint the component
	 */
	public void paintComponent(Graphics g) {
		g.drawImage(img, 0, 0, null);
	}

}
