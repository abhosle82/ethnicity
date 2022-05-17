package enablement.transformers.hackathon.ethnicity.poi;

import enablement.transformers.hackathon.ethnicity.entities.EthnicityBean;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

public class WriteFirstNameLastName {
        public static void writeDiversityAndInclusionData(List<EthnicityBean> lsEB){//Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("DiversityInclusionData1");

        //This data needs to be written (Object[])
        Map<String, Object[]> data = new LinkedHashMap<>();
            data.put("1", new Object[] {"DunsNum", "FirstName", "LastName"});
            int iRowCount=1;
            for(EthnicityBean bean: lsEB){
                ++iRowCount;
                data.put(String.valueOf(iRowCount), new Object[] {bean.getDunsNumber(), bean.getFirstName(), bean.getLastName()});
            }

        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset){
                Row row = sheet.createRow(rownum++);
                Object [] objArr = data.get(key);
                int cellnum = 0;
                for (Object obj : objArr){
                    Cell cell = row.createCell(cellnum++);
                    if(obj instanceof String)
                        cell.setCellValue((String)obj);
                    else if(obj instanceof Integer)
                        cell.setCellValue((Integer)obj);
                }
        }
        try{
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("hackathon_flname_Leader_2.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
