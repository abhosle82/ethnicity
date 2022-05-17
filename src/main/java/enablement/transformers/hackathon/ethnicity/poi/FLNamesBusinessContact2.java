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
import java.util.Iterator;
import java.util.List;

public class FLNamesBusinessContact2 {
    public static List<EthnicityBean> readDiversityAndInclusionData(){
        CompanyDiversityInfo objCDI = null;
        LeaderDiversityInfo objLDI = null;


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
            int iLenBusinessContact = 0;
            String strBusinessContactName = null;
            String strFLName[] = null;
            int iIndexVal = 0;
            EthnicityBean bean = null;
            while (rowIterator.hasNext())
            {
                bean = new EthnicityBean();
                objCDI = new CompanyDiversityInfo();
                objLDI = new LeaderDiversityInfo();


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

                    }else if(iColumnIndex == 9){
                        objLDI.setName(strValue);
                    }else if(iColumnIndex == 11){
                        if(objLDI.getName()!=null){
                            strBusinessContactName = objLDI.getName();
                            iLenBusinessContact = strBusinessContactName.trim().length();
                        }
                        if(strBusinessContactName!=null && iLenBusinessContact > 0){
                            iIndexVal = strBusinessContactName.lastIndexOf("-");
                            if(iIndexVal!=-1) {
                                strBusinessContactName=strBusinessContactName.substring(0, iIndexVal);
                                strBusinessContactName=strBusinessContactName.replace("-"," ");

                                System.out.println("Index::::"+iIndexVal+" |"+strBusinessContactName+"|");
                                strFLName = strBusinessContactName.trim().split("\\s+");

                                if(strFLName!=null && strFLName.length == 5){
                                    if(strFLName[0].length()==1){
                                        bean.setFirstName( strFLName[1]);
                                        bean.setLastName( strFLName[4]);
                                    }else{
                                        bean.setFirstName( strFLName[0]);
                                        bean.setLastName( strFLName[4]);
                                    }

                                }if(strFLName!=null && strFLName.length == 4){
                                    if(strFLName[0].length()==1){
                                        bean.setFirstName( strFLName[1]);
                                        bean.setLastName( strFLName[3]);
                                    }else{
                                        bean.setFirstName( strFLName[0]);
                                        bean.setLastName( strFLName[3]);
                                    }

                                }else if(strFLName!=null && strFLName.length == 3){
                                    if(strFLName[0].length()==1){
                                        bean.setFirstName( strFLName[1]);
                                        bean.setLastName( strFLName[2]);
                                    }else{
                                        bean.setFirstName( strFLName[0]);
                                        bean.setLastName( strFLName[2]);
                                    }

                                }else if(strFLName!=null && strFLName.length == 2){
                                    bean.setFirstName( strFLName[0]);
                                    bean.setLastName( strFLName[1]);
                                }
                            }
                        }
                        if(bean.getFirstName()!=null && bean.getFirstName().trim().length()>0 && bean.getLastName()!=null && bean.getLastName().trim().length() > 0) {
                            System.out.println(bean.getFirstName()+"|"+bean.getLastName());
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
