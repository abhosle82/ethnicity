package enablement.transformers.hackathon.ethnicity.poi;

import enablement.transformers.hackathon.ethnicity.entities.CompanyDiversityInfo;
import enablement.transformers.hackathon.ethnicity.entities.EthnicityBean;
import enablement.transformers.hackathon.ethnicity.entities.LeaderDiversityInfo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;

public class ReadFirstNameLastName {
    public static List<EthnicityBean> readDiversityAndInclusionData(){
        CompanyDiversityInfo objCDI = null;
        LeaderDiversityInfo objLDI1 = null;
        LeaderDiversityInfo objLDI2 = null;

        List<EthnicityBean> lsEB = new ArrayList<>();
        try
        {
            FileInputStream file = new FileInputStream(new File("Hackathon_Data_MinorityWomenOwned_2022 v1.xlsx"));


            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            int i=0;
            String strValue = null;
            int iColumnIndex = 0;
            String strBusinessContact1 = null;
            int iLenBusinessContact1 = 0;
            String strBusinessContact2 = null;
            int iLenBusinessContact2 = 0;
            String strBusinessContactName = null;
            String strFLName[] = null;
            int iIndexVal = 0;
            EthnicityBean bean = null;
            while (rowIterator.hasNext())
            {
                bean = new EthnicityBean();
                objCDI = new CompanyDiversityInfo();
                objLDI1 = new LeaderDiversityInfo();
                objLDI2 = new LeaderDiversityInfo();

                i++;
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                if(row.getRowNum()==0){
                    continue;
                }
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    iColumnIndex =cell.getColumnIndex();

                    if(cell.getCellType() == CellType.NUMERIC){
                        strValue = NumberToTextConverter.toText(cell.getNumericCellValue()) ;
                    }else{
                        strValue=cell.getStringCellValue() ;
                    }
                    if(iColumnIndex == 0){
                        bean.setDunsNumber(strValue);
                    }else if(iColumnIndex == 8){
                        objLDI1.setName(strValue);
                    }else if(iColumnIndex == 9){
                        objLDI2.setName(strValue);
                    }else if(iColumnIndex == 11){
                        if(objLDI1.getName()!=null){
                            strBusinessContact1 = objLDI1.getName();
                            iLenBusinessContact1 = strBusinessContact1.trim().length();
                        }
                        if(objLDI2.getName()!=null){
                            strBusinessContact2 = objLDI2.getName();
                            iLenBusinessContact2 = strBusinessContact2.trim().length();
                        }
                        if(iLenBusinessContact1 != 0 && iLenBusinessContact2 !=0){
                            strBusinessContactName = strBusinessContact1;
                        }else if(iLenBusinessContact1 != 0 && iLenBusinessContact2 ==0){
                            strBusinessContactName = strBusinessContact1;
                        }else if(iLenBusinessContact1 == 0 && iLenBusinessContact2 !=0){
                            strBusinessContactName = strBusinessContact2;
                        }else if(iLenBusinessContact1 == 0 && iLenBusinessContact2 ==0){
                            strBusinessContactName = null;
                        }
                        if(strBusinessContactName!=null){
                            iIndexVal = strBusinessContactName.lastIndexOf("-");
                            //System.out.println("Index Of :::"+iIndexVal);
                            if(iIndexVal!=-1) {
                                strBusinessContactName=strBusinessContactName.substring(0, iIndexVal);
                                strBusinessContactName=strBusinessContactName.replace("-"," ");

                                if("804660298".equals(bean.getDunsNumber())){
                                    System.out.println("strBusinessContactName::::"+strBusinessContactName);
                                }
                               // System.out.println("Index::::"+iIndexVal+" |"+strBusinessContactName+"|");
                                strFLName = strBusinessContactName.trim().split("\\s+");
                                //System.out.println(strFLName.length);

                                if(strFLName!=null && strFLName.length == 5){
                                    if(strFLName[0].length()==1){
                                        /*System.out.println("First Name:"+ strFLName[1]);
                                        System.out.println("Last Name:"+  strFLName[2]);*/
                                        bean.setFirstName( strFLName[1]);
                                        bean.setLastName( strFLName[4]);
                                    }else{
                                        /*System.out.println("First Name:"+ strFLName[0]);
                                        System.out.println("Last Name:"+  strFLName[2]);*/
                                        bean.setFirstName( strFLName[0]);
                                        bean.setLastName( strFLName[4]);
                                    }

                                }if(strFLName!=null && strFLName.length == 4){
                                    if(strFLName[0].length()==1){
                                        /*System.out.println("First Name:"+ strFLName[1]);
                                        System.out.println("Last Name:"+  strFLName[2]);*/
                                        bean.setFirstName( strFLName[1]);
                                        bean.setLastName( strFLName[3]);
                                    }else{
                                        /*System.out.println("First Name:"+ strFLName[0]);
                                        System.out.println("Last Name:"+  strFLName[2]);*/
                                        bean.setFirstName( strFLName[0]);
                                        bean.setLastName( strFLName[3]);
                                    }

                                }else if(strFLName!=null && strFLName.length == 3){
                                    if(strFLName[0].length()==1){
                                        /*System.out.println("First Name:"+ strFLName[1]);
                                        System.out.println("Last Name:"+  strFLName[2]);*/
                                        bean.setFirstName( strFLName[1]);
                                        bean.setLastName( strFLName[2]);
                                    }else{
                                        /*System.out.println("First Name:"+ strFLName[0]);
                                        System.out.println("Last Name:"+  strFLName[2]);*/
                                        bean.setFirstName( strFLName[0]);
                                        bean.setLastName( strFLName[2]);
                                    }

                                }else if(strFLName!=null && strFLName.length == 2){
                                   /* System.out.println("First Name:"+ strFLName[0]);
                                    System.out.println("Last Name:"+  strFLName[1]);*/
                                    bean.setFirstName( strFLName[0]);
                                    bean.setLastName( strFLName[1]);
                                }
                            }
                        }
                        if(bean.getFirstName()!=null && bean.getFirstName().trim().length()>0 && bean.getLastName()!=null && bean.getLastName().trim().length() > 0) {
                            lsEB.add(bean);
                        }
                    }
                }
                //System.out.println("#################### DONE ##########################");
            }

            file.close();
        }catch (Exception e){
            e.printStackTrace();
        }
        return lsEB;
    }
}
