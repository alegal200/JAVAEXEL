
package be.arma;


import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;

import java.io.IOException;
import java.util.ArrayList;


public class famille {


    private ArrayList<String> fam1 ;
    private ArrayList<String> fam2 ;
    private ArrayList<String> fam3 ;
    private ArrayList<String> fam4 ;
    private ArrayList<String> erreur ;
    private FileInputStream file1 ;
    private Workbook workbook1;
    private Sheet sheet1 ;



    public famille(String lienPLV )  {

        fam1 = new ArrayList<String>();
        fam2 = new ArrayList<String>();
        fam3 = new ArrayList<String>();
        fam4 = new ArrayList<String>();
        erreur = new ArrayList<String>();

        try{
            file1 = new FileInputStream(new File(lienPLV));
            workbook1 = new  XSSFWorkbook(file1);
            sheet1 = workbook1.getSheetAt(0);

        }catch (Exception e1 ){
            JOptionPane.showMessageDialog(null,"err n000f"+ e1);
        }



    }

    public void createPLV(){

        for (int i = 0; i <= sheet1.getLastRowNum() ; i++) {
            if(sheet1.getRow(i).getCell(0) == null)
                break;
            addarray(fam1, i ,3 );
            addarray(fam2, i ,4 );
            addarray(fam3, i ,5 );
            addarray(fam4, i ,6 );

        }







        try {
            file1.close();
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null,"err n00ff"+ e);
        }

    }

    private void addarray(ArrayList<String> fam, int i, int i1) {


        if(sheet1.getRow(i) == null)
            return;
        if(sheet1.getRow(i).getCell(i1)== null|| sheet1.getRow(i).getCell(i1).getStringCellValue().equals("")  )
            return;
        fam.add( sheet1.getRow(i).getCell(i1).getStringCellValue() ) ;


    }


    public void affchefamerr(){

        for (int i = 0; i < erreur.size(); i++) {
            System.out.println("val "+i+" :"+erreur.get(i));
        }
    }

    public  boolean IsonFam(String s){
        if(s.length() == 6 )
            return  fam4.contains(s);
        if(s.length() == 5 )
            return  fam3.contains(s);
        if(s.length() == 4)
            return  fam2.contains(s);
        if(s.length() == 3)
            return  fam1.contains(s);
        return false ;
    }

    public void info(ArrayList<String> arrayy , String nomFourn){
        for (int i = 0; i < arrayy.size(); i++) {
            if(IsonFam(arrayy.get(i))== false )
                erreur.add(nomFourn+" fam : "+arrayy.get(i)) ;
        }
    }
    public ArrayList<String> getErreurList(){
        return  erreur ;
    }


}

