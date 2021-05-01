/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ewb.qa.tdd.GUI;

import java.awt.*;
import java.awt.event.*;

/**
 *
 * @author JPE61800
 */
public class TDDMouseListener implements MouseListener {
    //addMouseListener(this);
    private int flagIndex;
    
    public void mouseClicked(MouseEvent e){
        setFlagIndex(1);
    }

    @Override
    public void mousePressed(MouseEvent e) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void mouseReleased(MouseEvent e) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void mouseEntered(MouseEvent e) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void mouseExited(MouseEvent e) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
    
    public void setFlagIndex(int givenIndex){
        flagIndex = givenIndex;
    }
    
    public int getFlagIndex(){
        return flagIndex;
    }
}
