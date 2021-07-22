package be.arma;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;


public class modification {

    private boolean Bdate ;
    private boolean Bfourn ;
    private boolean BTVAcode ;
    private boolean BPolitiqueTarif ;
    private boolean BRecyclageHTV ;
    private boolean BmoveIntrastat ;
    private boolean BcleanEAN ;

    private  final String LienMat ;
    private  final String NomFourn ;
    private  final String LienColler ;
    private final ArrayList<String> ErreurList ;

    public modification(String lienMat ,  String nomFourn ,String lienColler ){


         LienMat = lienMat ;
         NomFourn = nomFourn ;
         LienColler = lienColler ;
         Bdate = false ;
         Bfourn = false ;
         BTVAcode = false ;
         BPolitiqueTarif = false ;
         BRecyclageHTV = false ;
         BmoveIntrastat = false ;
         BcleanEAN  = false ;

         ErreurList = new ArrayList<>();

    }

    public void setBdate(boolean b ){
        Bdate =  b ;
    }
    public void setBfourn(boolean b ){
        Bfourn =  b ;
    }
    public void setBTVAcode(boolean b ){
        BTVAcode =  b ;
    }
    public void setBPolitiqueTarif(boolean b ){
        BPolitiqueTarif =  b ;
    }
    public void setBRecyclageHTV(boolean b ){
        BRecyclageHTV =  b ;
    }
    public void setBmoveIntrastat(boolean b ){
        BmoveIntrastat=  b ;
    }
    public void setBcleanEAN(boolean b ){
        BcleanEAN=  b ;
    }
    public ArrayList<String> getError() {
        return ErreurList;
    }








    public void modifmoicastp() {

        // ouverture

        FileInputStream filematvide ;
        Workbook workbook ;

        try {
            filematvide = new FileInputStream(LienMat);

            workbook  = new XSSFWorkbook(filematvide);


        }catch(Exception e1){
            JOptionPane.showMessageDialog(null,"err n0003"+ e1);
            return;
        }

        //

        // modification
        modifs(workbook);







        // fermeture

        try{




            FileOutputStream fos = new FileOutputStream(LienColler+"\\"+NomFourn) ;

            workbook.write(fos);

            fos.flush();
            fos.close();

            filematvide.close();


        }catch (Exception e){
            JOptionPane.showMessageDialog(null, "erreur d ecriture");
        }






    }

    private void modifs(Workbook workbook) {

        Sheet sheeteArticle = workbook.getSheet("ARTICLE");
        // Sheet sheeteStock = workbookecriture.getSheetAt(9);


        for (int i = 17; i < sheeteArticle.getLastRowNum(); i++) {
            if(sheeteArticle.getRow(i).getCell(0)==null || sheeteArticle.getRow(i).getCell(0).getStringCellValue().equals("")  )
                break;
            try{


                if( Bdate){

                    pastval(sheeteArticle ,i , 78,"20210101");
                    pastval(sheeteArticle ,i , 112,"20210101");

                }


                if(Bfourn){

                    Cell ced3e = sheeteArticle.getRow(i).getCell(64) ;
                    if(    ced3e != null && ced3e.getStringCellValue().equals("'-----")     )
                        pastval(sheeteArticle ,i , 64, ""  );


                }


                if(BmoveIntrastat){

                     if( sheeteArticle.getRow(i).getCell(132).getStringCellValue() !=null )
                         pastval(sheeteArticle ,i ,18 , sheeteArticle.getRow(i).getCell(132).getStringCellValue() );

                }

                if ( BcleanEAN ) {
                    if (sheeteArticle.getRow(i) != null){
                        Cell ced3e = sheeteArticle.getRow(i).getCell(56);
                    if (ced3e != null) {
                        StringBuilder tmp = new StringBuilder(ced3e.getStringCellValue());
                         if ( ! tmp.toString().equals("") ) {
                            try {

                               double d ;
                                d = Double.parseDouble(tmp.toString());

                                if (tmp.length() != 13  && tmp.length() != 8 ){
                                    if( tmp.length() < 8 ){
                                        if(tmp.length()==1){
                                            ErreurList.add("code bare erreur nbr    ; " + NomFourn + "    ; ln " + i + " ref ;  "+  sheeteArticle.getRow(i).getCell(0).getStringCellValue() + "   ;   code article  ; "+sheeteArticle.getRow(i).getCell(2).getStringCellValue() +"  ;   ean   ; "+ sheeteArticle.getRow(i).getCell(56).getStringCellValue());
                                            pastval(sheeteArticle , i ,56 , "");
                                        }
                                       else{
                                            while (tmp.length() != 8 ){
                                                tmp.insert(0, "0");
                                            }
                                        }
                                        if(tmp.length() == 8)
                                        pastval(sheeteArticle , i ,56 , tmp.toString());
                                    }else{
                                        if(tmp.length() < 13 ){
                                            while (tmp.length()!= 13 )
                                                tmp.insert(0, "0");

                                            pastval(sheeteArticle , i ,56 , tmp.toString());
                                        }else{
                                            ErreurList.add("code bare erreur nbr    ; " + NomFourn + "    ; ln " + i + " ref ;  "+  sheeteArticle.getRow(i).getCell(0).getStringCellValue() + "   ;   code article  ; "+sheeteArticle.getRow(i).getCell(2).getStringCellValue() +"  ;   ean   ; "+ sheeteArticle.getRow(i).getCell(56).getStringCellValue());
                                            pastval(sheeteArticle , i ,56 , "");
                                        }

                                    }

                                }
                           } catch (Exception e) {
                                ErreurList.add("code bare erreur nbr    ; " + NomFourn + "    ; ln " + i + " ref ;  "+  sheeteArticle.getRow(i).getCell(0).getStringCellValue() + "   ;   code article  ; "+sheeteArticle.getRow(i).getCell(2).getStringCellValue() +"  ;   ean   ; "+ sheeteArticle.getRow(i).getCell(56).getStringCellValue());
                                pastval(sheeteArticle , i ,56 , "");
                            }

                        }

                    }
                }

                }






                if(BTVAcode){

                    pastval(sheeteArticle , i,  19 ,  sheeteArticle.getRow(i).getCell(19).getStringCellValue().substring(0,1)  );

                }



                if( BRecyclageHTV ){

                    if( sheeteArticle.getRow(i).getCell(23).getStringCellValue().contains("-")   )
                        pastval(sheeteArticle , i,  23 ,  sheeteArticle.getRow(i).getCell(23).getStringCellValue().substring(1,  sheeteArticle.getRow(i).getCell(23).getStringCellValue().length() )  );


                }
                if( BTVAcode ){
                    pastval(sheeteArticle , i,  75 , "00");
                }














            }catch (Exception e){
                System.out.println("erreur général l "+i+" fourn : "+NomFourn+" msg :"+e.getMessage());
            }

        }

        // vider la derniere colone
        if(BmoveIntrastat ){

            for (int i = 0; i < sheeteArticle.getLastRowNum(); i++){
                if(sheeteArticle.getRow(i) != null )
                    if(sheeteArticle.getRow(i).getCell(132) != null )
                         sheeteArticle.getRow(i).getCell(132).setCellValue("");
            }
        }





    }

    private void pastval(Sheet sheetp , int i, int i1, String s) {

        if(sheetp.getRow(i) == null){
            sheetp.createRow(i);
        }


        Cell c = sheetp.getRow(i).getCell(i1) ;
        if(c == null ){
            c =  sheetp.getRow(i).createCell(i1);
        }
        c.setCellValue(s);


    }




}
