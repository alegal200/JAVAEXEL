package be.arma;


import javax.swing.*;
import java.awt.*;

public class start {

    public static JFrame frame = new JFrame("App");;

    public static void main(String[] args) {


        frame.setContentPane(new App().panelMain);
        frame.setTitle("Arma - Matrice Manager");
        frame.setSize(700, 400);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setIconImage(Toolkit.getDefaultToolkit().getImage(start.class.getResource("/images/logo.png")));
        frame.pack();
        frame.setVisible(true);
        frame.setLayout(new BorderLayout());




    }
        


    }

