package be.arma;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;

public class creationExcel {

    private  final String LienMats ;
    private  final String NomFourn ;
    private  final String LienMatVide ;
    private  final ArrayList<String> famList ;
    private  boolean promaco;
    private  final  ArrayList<String> ErrList ;

    public creationExcel(String lienMats , String nomFourn , String lienMatVide , boolean promacoo ){
        LienMats = lienMats ;
        NomFourn = nomFourn ;
        LienMatVide = lienMatVide ;
        famList = new ArrayList<>();
        ErrList = new ArrayList<>();
        promaco = false ;
        promaco = promacoo;
        this.fabrique();
    }
    private void fabrique(){


        // datas


        Workbook workbookecriture ;
        FileInputStream file1 ;
        FileInputStream file2 ;
        FileInputStream file3 ;
        FileInputStream filematvide ;
        Workbook workbook1 ;
        Workbook workbook2  ;
        Workbook workbook3  ;


        // ouverture fichiers


        try {
            file1 = new FileInputStream(LienMats + "\\1.xls");
            file2 = new FileInputStream(LienMats + "\\2.xls");
            file3 = new FileInputStream(LienMats + "\\3.xls");

            workbook1 = new HSSFWorkbook(file1);
            workbook2 = new HSSFWorkbook(file2);
            workbook3 = new HSSFWorkbook(file3);

        }catch(Exception e1){
            JOptionPane.showMessageDialog(null,"err n0001"+ e1);
            return;
        }


        // copier coller matrcie vide





        try{

            var source = new File(LienMatVide);
            var dest = new File(LienMats+"\\matvide.xlsx");


            Files.copy(  source.toPath(), dest.toPath() , StandardCopyOption.REPLACE_EXISTING);


        }
        catch (Exception e) {
            JOptionPane.showMessageDialog(null,"err 0002"+ e);
            return;
        }



        // ouverture matrice vide copier




        try {
            filematvide = new FileInputStream(LienMats + "\\matvide.xlsx");

            workbookecriture  = new XSSFWorkbook(filematvide);


        }catch(Exception e1){
            JOptionPane.showMessageDialog(null,"err n0003"+ e1);
            return;
        }


        workbookecriture = copiecollemoica( workbook1,  workbook2,  workbook3,  workbookecriture) ;

        remplir(workbookecriture);


        // sauvegarde



        try{

            FileOutputStream fos = null ;

            if(famList.size() > 1 || ErrList.size() > 1 )
                fos = new FileOutputStream(LienMats+"\\"+NomFourn+".xlsx") ;
            else
                fos = new FileOutputStream(LienMats+"\\vide.xlsx") ;
            workbookecriture.write(fos);

            fos.flush();
            fos.close();

            file1.close();
            file2.close();
            file3.close();
            file3.close();

        }catch (Exception e){
            JOptionPane.showMessageDialog(null, "erreur d ecriture");
        }








    }

    private void remplir(Workbook workbookecriture ) {


        Sheet sheeteArticle = workbookecriture.getSheet("ARTICLE");
        Sheet sheeteStock = workbookecriture.getSheetAt(9);



        for (int i = 17; i <= sheeteArticle.getLastRowNum(); i++) {
            if(sheeteArticle.getRow(i).getCell(0)==null || sheeteArticle.getRow(i).getCell(0).getStringCellValue().equals("")  )
                break;
            try{




                pastval(sheeteArticle ,i , 3,"M/SES");
                pastval(sheeteArticle ,i , 27,"PCE");
                pastval(sheeteArticle ,i , 28,"PCE");
                pastval(sheeteArticle ,i , 29,"N");
                pastval(sheeteArticle ,i , 30,"PCE");
                pastval(sheeteArticle ,i , 31,"N");
                pastval(sheeteArticle ,i , 32,"PCE");
                pastval(sheeteArticle ,i , 33,"PCE");
                pastval(sheeteArticle ,i , 34,"PCE");
                pastval(sheeteArticle ,i , 35,"PCE");
                pastval(sheeteArticle ,i , 36,"PCE");
                pastval(sheeteArticle ,i , 37,"PCE");

                pastval(sheeteArticle ,i , 47,"QT");

                pastval(sheeteArticle ,i , 49,"N");

                pastval(sheeteArticle ,i , 75,"00");
                pastval(sheeteArticle ,i , 76,"O");
                pastval(sheeteArticle ,i , 77,"O");
                pastval(sheeteArticle ,i , 78,"20210101");
                pastval(sheeteArticle ,i , 79,"0");

                pastval(sheeteArticle ,i , 112,"20210101");
                pastval(sheeteArticle ,i , 113,"1");
                pastval(sheeteArticle ,i , 114,"0");
                pastval(sheeteArticle ,i , 115,"C2");


                // web export

                Cell c = sheeteArticle.getRow(i).getCell(127) ;
                if(    c != null && c.getStringCellValue().equals("A")     )
                    pastval(sheeteArticle ,i , 127,"O");
                else
                    pastval(sheeteArticle ,i , 127,"N");



                // familles


                try {


                    if ( !promaco ) {



                        Cell cacti = sheeteArticle.getRow(i).getCell(12);
                        Cell cfam = sheeteArticle.getRow(i).getCell(13);
                        if (cacti != null && cfam != null) {
                            String famillecomplet = cacti.getStringCellValue().substring(0, 2) + cfam.getStringCellValue();

                            pastval(sheeteArticle, i, 12, cacti.getStringCellValue().substring(0, 2));
                            try {
                                pastval(sheeteArticle, i, 13, famillecomplet.substring(0, 3));
                                famList.add(famillecomplet.substring(0, 3));
                            } catch (Exception e) {
                                //
                            }
                            try {
                                pastval(sheeteArticle, i, 14, famillecomplet.substring(0, 4));
                                famList.add(famillecomplet.substring(0, 4));
                            } catch (Exception e) {
                                //
                            }
                            try {
                                pastval(sheeteArticle, i, 15, famillecomplet.substring(0, 5));
                                famList.add(famillecomplet.substring(0, 5));
                            } catch (Exception e) {
                                //
                            }
                            try {
                                pastval(sheeteArticle, i, 16, famillecomplet.substring(0, 6));
                                famList.add(famillecomplet.substring(0, 6));
                            } catch (Exception e) {
                                //
                            }


                        }

                    } else {

                        // promaco
                        Cell cacti = sheeteArticle.getRow(i).getCell(12);
                        if (cacti != null)
                            pastval(sheeteArticle, i, 12, cacti.getStringCellValue().substring(0, 2));

                    }

                }catch (Exception e ){
                    ErrList.add("erreur famille : "+NomFourn+"a la ligne n "+i) ;
                }



                // fournisseur
/*
                if(  sheeteArticle.getRow(i).getCell(64)==null || sheeteArticle.getRow(i).getCell(64).getStringCellValue().equals("") )
                    pastval(sheeteArticle ,i , 64, "'-----"  );
*/
                // description + recherche

                if(  sheeteArticle.getRow(i).getCell(0) != null ||  ! sheeteArticle.getRow(i).getCell(64).getStringCellValue().equals("")   ){

                    String Libelle =  sheeteArticle.getRow(i).getCell(0).getStringCellValue() ;
                    String recherche ;
                    recherche = Libelle.split(" ")[0] ;
                    if( recherche.length() < 3  )
                        recherche = Libelle.split(" ")[1] ;

                    try{
                        double f = Double.parseDouble(recherche) ;
                        recherche = Libelle.split(" ")[1] ;
                        f = Double.parseDouble(recherche) ;
                        recherche = Libelle.split(" ")[2] ;

                    }catch (Exception e){
                        //
                    }
                    if(recherche == null )
                        recherche = "" ;
                    pastval(sheeteArticle,i,4,recherche  );




                    if(Libelle.length() > 40 ){

                        String fin = "" ;
                        int lastnbr = Libelle.lastIndexOf(" ") ;
                        String  temp = Libelle ;
                        while (lastnbr > 40 ){

                            temp = temp.substring(0, lastnbr -1 );
                            lastnbr = temp.lastIndexOf(" ");
                            if(  ! temp.contains(" ")  ){
                                lastnbr = 39 ;
                            }

                        }
                        int taille = Libelle.length() - lastnbr ;
                        if(lastnbr < 0 )
                            lastnbr = 0 ;
                        if(taille > 0 ){
                            fin =Libelle.substring( lastnbr ,Libelle.length()) ;
                            pastval(sheeteArticle,i,6,fin);
                        }
                        pastval(sheeteArticle,i,5,Libelle.substring(0,lastnbr)  );

                    }
                    else{
                        pastval(sheeteArticle,i,5,Libelle  );
                    }


                }




                // colone w type de taxe


                Cell ced3e = sheeteArticle.getRow(i).getCell(23) ;
                if(    ced3e != null && ! ced3e.getStringCellValue().equals("0.0")     )
                    pastval(sheeteArticle,i ,22 ,"ED3E");

                // feuille stock
                Cell cedEe = sheeteArticle.getRow(i).getCell(129) ;
                if(    cedEe != null && cedEe.getStringCellValue().equals("A")     )
                    pastval(sheeteStock ,i-12 ,12 ,"O");




                pastval(sheeteStock, i-12 , 1 , sheeteArticle.getRow(i).getCell(2).getStringCellValue() );
                if(promaco)
                    pastval(sheeteStock, i-12 , 3 , "20");
                else
                pastval(sheeteStock, i-12 , 3 , "10");
                pastval(sheeteStock, i-12 , 4 , "O");


                // conversion des longeurs

                try {

                    if (sheeteArticle.getRow(i).getCell(39) != null && ! sheeteArticle.getRow(i).getCell(39).getStringCellValue().equals("0")  &&  ! sheeteArticle.getRow(i).getCell(39).getStringCellValue().equals("0.0")  && ! sheeteArticle.getRow(i).getCell(39).getStringCellValue().isEmpty() ) {
                        double d = Double.parseDouble(sheeteArticle.getRow(i).getCell(39).getStringCellValue()) * 1000;
                        sheeteArticle.getRow(i).getCell(39).setCellValue(d);

                    }
                    if (sheeteArticle.getRow(i).getCell(40) != null &&  ! sheeteArticle.getRow(i).getCell(40).getStringCellValue().equals("0")  && ! sheeteArticle.getRow(i).getCell(40).getStringCellValue().equals("0.0")   &&  ! sheeteArticle.getRow(i).getCell(40).getStringCellValue().isEmpty()  ) {
                        double d = Double.parseDouble(sheeteArticle.getRow(i).getCell(40).getStringCellValue()) * 1000;
                        sheeteArticle.getRow(i).getCell(40).setCellValue(d);
                    }
                    if (sheeteArticle.getRow(i).getCell(41) != null &&! sheeteArticle.getRow(i).getCell(41).getStringCellValue().equals("0")  &&   !sheeteArticle.getRow(i).getCell(41).getStringCellValue().equals("0.0")   && !  sheeteArticle.getRow(i).getCell(40).getStringCellValue().isEmpty()  ) {
                        double d = Double.parseDouble(sheeteArticle.getRow(i).getCell(41).getStringCellValue()) * 1000;
                        sheeteArticle.getRow(i).getCell(41).setCellValue(d);
                    }
                    if (sheeteArticle.getRow(i).getCell(42) != null && !sheeteArticle.getRow(i).getCell(42).getStringCellValue().equals("0")   &&! sheeteArticle.getRow(i).getCell(40).getStringCellValue().isEmpty()   ) {
                        try{
                            double d = Double.parseDouble(sheeteArticle.getRow(i).getCell(42).getStringCellValue()) * 1000;
                            sheeteArticle.getRow(i).getCell(42).setCellValue(d);
                        }catch (Exception e){
                        //
                        }

                    }


                }catch (Exception e ){
                    // System.out.println("err n 0088"+e.getMessage() );
                }



                // CONVERSION DES TAUX TVA
                pastval(sheeteArticle , i,  19 ,  sheeteArticle.getRow(i).getCell(19).getStringCellValue().substring(0,1)  );


                // retirer le - dans la colone x

                if( sheeteArticle.getRow(i).getCell(23).getStringCellValue().contains("-")   )
                    pastval(sheeteArticle , i,  23 ,  sheeteArticle.getRow(i).getCell(23).getStringCellValue().substring(1,  sheeteArticle.getRow(i).getCell(23).getStringCellValue().length() )  );

                // code bare verification + modification



                if (sheeteArticle.getRow(i) != null){
                    Cell ced3 = sheeteArticle.getRow(i).getCell(56);
                    if (ced3 != null) {
                        StringBuilder tmp = new StringBuilder(ced3.getStringCellValue());
                        if ( ! tmp.toString().equals("") ) {
                            try {

                                double d ;
                                d = Double.parseDouble(tmp.toString());

                                if (tmp.length() != 13  && tmp.length() != 8 ){
                                    if( tmp.length() < 8 ){
                                        if(tmp.length()==1){
                                            ErrList.add("code bare erreur nbr    ; " + NomFourn + "    ; ln " + i + " ref ;  "+  sheeteArticle.getRow(i).getCell(0).getStringCellValue() + "   ;   code article  ; "+sheeteArticle.getRow(i).getCell(2).getStringCellValue() +"  ;   ean   ; "+ sheeteArticle.getRow(i).getCell(56).getStringCellValue());
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
                                            ErrList.add("code bare erreur nbr    ; " + NomFourn + "    ; ln " + i + " ref ;  "+  sheeteArticle.getRow(i).getCell(0).getStringCellValue() + "   ;   code article  ; "+sheeteArticle.getRow(i).getCell(2).getStringCellValue() +"  ;   ean   ; "+ sheeteArticle.getRow(i).getCell(56).getStringCellValue());
                                            pastval(sheeteArticle , i ,56 , "");
                                        }

                                    }

                                }
                            } catch (Exception e) {
                                ErrList.add("code bare erreur nbr    ; " + NomFourn + "    ; ln " + i + " ref ;  "+  sheeteArticle.getRow(i).getCell(0).getStringCellValue() + "   ;   code article  ; "+sheeteArticle.getRow(i).getCell(2).getStringCellValue() +"  ;   ean   ; "+ sheeteArticle.getRow(i).getCell(56).getStringCellValue());
                                pastval(sheeteArticle , i ,56 , "");
                            }

                        }

                    }
                }





                // intra stat

                if(  sheeteArticle.getRow(i) != null && sheeteArticle.getRow(i).getCell(132) !=null && sheeteArticle.getRow(i).getCell(132).getStringCellValue() !=null )
                    pastval(sheeteArticle ,i ,18 , sheeteArticle.getRow(i).getCell(132).getStringCellValue() );














            }catch (Exception e){
                System.out.println("erreur"+e.getMessage()  );
            }

        }

        // intrastat

        for (int i = 0; i < sheeteArticle.getLastRowNum(); i++){
            if(sheeteArticle.getRow(i) != null )
                if(sheeteArticle.getRow(i).getCell(132) != null )
                    sheeteArticle.getRow(i).getCell(132).setCellValue("");
        }




    }

    private void pastval(Sheet sheetp, int i, int i1, String s) {

        if(sheetp.getRow(i) == null){
            sheetp.createRow(i);
        }


        Cell c = sheetp.getRow(i).getCell(i1) ;
        if(c == null ){
            c =  sheetp.getRow(i).createCell(i1);
        }
        c.setCellValue(s);


    }


    private Workbook copiecollemoica(Workbook workbook1, Workbook workbook2, Workbook workbook3, Workbook workbookecriture) {

        Sheet sheet1 = workbook1.getSheetAt(0);
        Sheet sheet2 = workbook2.getSheetAt(0);
        Sheet sheet3 = workbook3.getSheetAt(0);
        Sheet sheetecriture = workbookecriture.getSheet("ARTICLE") ;




        int numrow = sheet1.getLastRowNum() ;
        int debutrow = 16 ;

        // 1 er copier

        for (int i = 0; i < numrow; i++) {

            for (int j = 0; j < sheet1.getRow(i).getLastCellNum(); j++) { // recup la dernier case de la ligne
                if( sheet1.getRow(i) != null && sheet1.getRow(i).getCell(j) != null ) {
                    Cell c = null;
                    try {
                        c = sheetecriture.getRow(debutrow+i).getCell(j) ;
                        if(sheetecriture.getRow(debutrow+i) == null)
                            sheetecriture.createRow(debutrow+i);
                        if(sheetecriture.getRow(debutrow+i).getCell(j) == null )
                            c = sheetecriture.getRow(debutrow+i).createCell(j);
                    }catch (Exception e){
                        if(sheetecriture.getRow(debutrow+i) == null)
                            sheetecriture.createRow(debutrow+i);
                        if(sheetecriture.getRow(debutrow+i).getCell(j) == null )
                            c = sheetecriture.getRow(debutrow+i).createCell(j);


                    }
                  //  if(c != null)
                        c.setCellValue( sheet1.getRow(i).getCell(j).toString()     );


                }


            }


        }



        // copier 2 eme

        numrow = sheet2.getLastRowNum() ;
        int ajoutp2 = 43 ;




        for (int i = 0; i < numrow; i++) {

            for (int j = 0 ; j < sheet2.getRow(i).getLastCellNum()  ; j++) { // recup la dernier case de la ligne

                if( sheet2.getRow(i) != null && sheet2.getRow(i).getCell(j ) != null ) {
                    Cell c = null;
                    try {
                        c = sheetecriture.getRow(debutrow+i).getCell(j+ajoutp2) ;
                        if(sheetecriture.getRow(debutrow+i) == null)
                            sheetecriture.createRow(debutrow+i);
                        if(sheetecriture.getRow(debutrow+i).getCell(j+ajoutp2) == null )
                            c = sheetecriture.getRow(debutrow+i).createCell(j+ajoutp2);
                    }catch (Exception e){
                        if(sheetecriture.getRow(debutrow+i) == null)
                            sheetecriture.createRow(debutrow+i);
                        if(sheetecriture.getRow(debutrow+i).getCell(j+ajoutp2) == null )
                            c = sheetecriture.getRow(debutrow+i).createCell(j+ajoutp2);

                    }

                        c.setCellValue( sheet2.getRow(i).getCell(j).toString()  );  //

                }


            }


        }



        // copier 3 eme

        numrow = sheet2.getLastRowNum() ;
        int ajoutp3 = 88 ;




        for (int i = 0; i < numrow; i++) {

            for (int j = 0 ; j < sheet3.getRow(i).getLastCellNum()  ; j++) { // recup la dernier case de la ligne

                if( sheet3.getRow(i) != null && sheet3.getRow(i).getCell(j ) != null ) {
                    Cell c = null;
                    try {
                        c = sheetecriture.getRow(debutrow+i).getCell(j+ajoutp3) ;
                        if(sheetecriture.getRow(debutrow+i) == null)
                            sheetecriture.createRow(debutrow+i);
                        if(sheetecriture.getRow(debutrow+i).getCell(j+ajoutp3) == null )
                            c = sheetecriture.getRow(debutrow+i).createCell(j+ajoutp3);
                    }catch (Exception e){
                        if(sheetecriture.getRow(debutrow+i) == null)
                            sheetecriture.createRow(debutrow+i);
                        if(sheetecriture.getRow(debutrow+i).getCell(j+ajoutp3) == null )
                            c = sheetecriture.getRow(debutrow+i).createCell(j+ajoutp3);

                    }

                        c.setCellValue( sheet3.getRow(i).getCell(j).toString()  );  //

                }


            }


        }






        return  workbookecriture ;



    }



    public  ArrayList<String> getFamList(){
        return  famList ;
    }
    public  ArrayList<String> getErrList(){ return  ErrList ; }


















}
