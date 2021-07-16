package be.arma;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;

public class App extends JFrame {

    // ------------
    // Déclaration de toutes les infterfaces
    // ------------

    public JPanel panelMain;
    private JPanel EditMatrice;
    private JPanel CreateMatrice;
    private JCheckBox PromacoCheck;
    private JButton CreateFournisseurButton;
    private JTextField FieldFournisseur;
    private JTextField FieldPath;
    private JPanel Parameter;
    private JLabel LabelFournisseur;
    private JLabel LabelPath;
    private JLabel NombreElements;
    private JPanel CreateMatriceFolder;
    private JProgressBar CreateProgressBar;
    private JLabel CreateProgressBarLabel;
    private JTextField FieldMatricesCreate;
    private JButton EditFournisseurButton;
    private JButton CreateMatricesButton;
    private JPanel EditMatriceFolder;
    private JButton EditMatricesButton;
    private JProgressBar EditProgressBar;
    private JLabel EditProgressBarLabel;
    private JCheckBox checkFamilleMatriceCheckBox;
    private JCheckBox DateMatriceCheckBox;
    private JTextField FieldMatricesEdit;
    private JTextField LienDuFichierEdit;
    private JTextField NomFournisseurEdit;
    private JTextField LienDeCopieEdit;
    private JCheckBox TVACodeMatriceCheckBox;
    private JCheckBox PolitiqueTarifMatriceCheckBox;
    private JCheckBox FournisseurMatriceCheckBox;
    private JCheckBox recyclageHTVMatriceCheckBox;
    private JTextField FieldCopieMatricesEdit;
    private JCheckBox DateMatricesCheckBox;
    private JCheckBox FournisseurMatricesCheckBox;
    private JCheckBox TVACodeMatricesCheckBox;
    private JCheckBox PolitiqueTarifMatricesCheckBox;
    private JCheckBox recyclageHTVMatricesCheckBox;
    private JLabel test;
    private JList list1;

    public int i;


    // ------------
    // Déclaration du HandlerThread
    // ------------

    public App() {

        CreateFournisseurButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                //NombreElements.setText("Nombre d'élèments : ");
                String Fournisseur = FieldFournisseur.getText();
                String path = FieldPath.getText();
                //creationExcel ex = new  creationExcel(path,Fournisseur,"//", "");

                famille fam = new famille("\\\\192.168.1.200\\export\\GEEK\\PLV.xlsx");

                fam.createPLV();
                boolean promaco;
                if(PromacoCheck.isSelected()){
                     promaco = true;
                }else{
                     promaco = false;
                }
                creationExcel ex = new creationExcel(path, Fournisseur, "\\\\192.168.1.200\\export\\DL\\cedric\\DL negoce\\FINI - fournisseurs\\matricevide.xlsx", promaco);
                FileWriter myWriter = null;
                try {
                    myWriter = new FileWriter(path + "\\famillerror.txt");
                } catch (IOException ioException) {
                    ioException.printStackTrace();
                }
                fam.info(ex.getFamList(), Fournisseur);

                ArrayList<String> erreurfamille = fam.getErreurList();
                try {
                    for (int w = 0; w < erreurfamille.size(); w++) {
                        myWriter.write("Erreur " + w + " : " + erreurfamille.get(w) + "\n");
                    }

                } catch (IOException error) {
                    JOptionPane.showMessageDialog(null, error);
                }

                try {
                    myWriter.close();
                } catch (IOException ioException) {
                    ioException.printStackTrace();
                }
            }
        });


        EditFournisseurButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                    modification mod = new modification(LienDuFichierEdit.getText(), NomFournisseurEdit.getText(), LienDeCopieEdit.getText());

                    if(DateMatriceCheckBox.isSelected()){
                        mod.setBdate(true);
                    }
                    if(FournisseurMatriceCheckBox.isSelected()){
                        mod.setBfourn(true);
                    }
                    if(TVACodeMatriceCheckBox.isSelected()){
                        mod.setBTVAcode(true);
                    }
                    if(PolitiqueTarifMatriceCheckBox.isSelected()){
                        mod.setBPolitiqueTarif(true);
                    }
                    if(recyclageHTVMatriceCheckBox.isSelected()){
                    mod.setBRecyclageHTV(true);
                    }
                    mod.modifmoicastp();
            }
        });
        CreateMatricesButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String Magasin;

                boolean promaco;
                if(PromacoCheck.isSelected()){
                    promaco = true;
                }else{
                    promaco = false;
                }
                String path = FieldMatricesCreate.getText();
                File directoryPath = new File(FieldMatricesCreate.getText());
                String contents[] = directoryPath.list();

                Thread threadcreate = new Thread(() ->{
                    try{
                        FileWriter myWriter = new FileWriter(path + "\\famillerror.txt");

                        for(int i=0; i<contents.length; i++) {
                            CreateProgressBarLabel.setText(i + 1  + "/" + contents.length +" fait");
                            double a = i+1;
                            double b = contents.length;
                            double pourcent = ( a / b) * 100;
                            CreateProgressBar.setValue((int) pourcent);
                            famille fam = new famille("\\\\192.168.1.200\\export\\GEEK\\PLV.xlsx");

                            fam.createPLV();

                            creationExcel ex = new creationExcel(path +"\\" + contents[i],contents[i] ,"\\\\192.168.1.200\\export\\DL\\cedric\\DL negoce\\FINI - fournisseurs\\matricevide.xlsx", promaco);

                            fam.info(ex.getFamList(),contents[i]);

                            ArrayList<String> erreurfamille = fam.getErreurList();

                            try {
                                for (int w = 0; w < erreurfamille.size(); w++) {
                                    myWriter.write("Erreur " + w + " : " + erreurfamille.get(w) + "\n");
                                }

                            }catch (IOException error){
                                JOptionPane.showMessageDialog(null, error);
                            }
                        }
                        myWriter.close();
                    }catch (Exception ex){
                        JOptionPane.showMessageDialog(null, ex);
                    }
                });

                threadcreate.start();
            }
        });
        EditMatricesButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                File directoryPath = new File(FieldMatricesEdit.getText());
                String contents[] = directoryPath.list();

                if(DateMatriceCheckBox.isSelected()){
                    Boolean checkDescription = true;
                }else{
                    Boolean checkDescription = false;
                }



                    Thread threadedit = new Thread(() ->{
                       try{
                           for(int i=0; i<contents.length; i++) {
                               EditProgressBarLabel.setText(i + 1 + "/" + contents.length + " fait");
                               double a = i + 1;
                               double b = contents.length;
                               double pourcent = (a / b) * 100;
                               EditProgressBar.setValue((int) pourcent);

                               modification mod = new modification(FieldMatricesEdit.getText() + "\\" + contents[i] + "\\" + contents[i] + ".xlsx", contents[i], FieldCopieMatricesEdit.getText());

                               if(DateMatricesCheckBox.isSelected()){
                                   mod.setBdate(true);
                               }
                               if(FournisseurMatricesCheckBox.isSelected()){
                                   mod.setBfourn(true);
                               }
                               if(TVACodeMatricesCheckBox.isSelected()){
                                   mod.setBTVAcode(true);
                               }
                               if(PolitiqueTarifMatricesCheckBox.isSelected()){
                                   mod.setBPolitiqueTarif(true);
                               }
                               if(recyclageHTVMatricesCheckBox.isSelected()){
                                   mod.setBRecyclageHTV(true);
                               }
                               mod.modifmoicastp();
                           }
                       }
                       catch (Exception ex){
                    JOptionPane.showMessageDialog(null, ex);
                }
                    });

                    threadedit.start();

            }
        });
    }

}
