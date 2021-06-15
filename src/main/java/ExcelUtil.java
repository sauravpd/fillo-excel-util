import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

import java.net.URL;

/**
 * Excel Util
 */
public class ExcelUtil
{
    private static ExcelUtil excelUtil;
    private Connection connection;
    private Fillo fillo;

    private ExcelUtil() {

    }

    public static ExcelUtil getInstance() {
        if (excelUtil == null) {
            excelUtil = new ExcelUtil();
        }
        return excelUtil;
    }

    /**
     * get Records from sheet
     * @param sheetName
     * @return
     */
    public Recordset getRecordsBySheetName(String sheetName) {
        Recordset records = null;
        try {
            String query = "Select * from " + sheetName;
            records = connection.executeQuery(query);
        } catch (Exception e) {
            System.out.println("Failed to retrieve records from sheet" + e.getMessage());
        }
        return records;
    }

    /**
     * update Record
     * @param sheetName
     * @param id
     * @param salary
     */
    public void updateRecord(String sheetName,String id,String salary){
        int records =0;
        try {
            String query = "Update "+sheetName+" set salary="+salary+" where EmpId="+id;
            records = connection.executeUpdate(query);
        } catch (Exception e) {
            System.out.println("Failed to update records "+ e.getMessage());
        }
        if(records>0){
            System.out.println("record updated successfully , row "+records+" affected");
        }
    }

    /**
     * delete Record
     * @param sheetName
     * @param id
     */
    public void deleteRecord(String sheetName,String id){
        int records =0;
        try {
            String query = "Delete from "+sheetName+" where EmpId="+id;
            records = connection.executeUpdate(query);
        } catch (Exception e) {
            System.out.println("Failed to delete records "+ e.getMessage());
        }
        if(records>0){
            System.out.println("record deleted successfully , row "+records+" affected");
        }
    }

    /**
     * get Connection object
     * @param filepath
     * @return
     */
    public void setConnection(String filepath) {
        try {
            fillo = new Fillo();
            URL filename = ExcelUtil.class.getResource("testdata.xlsx");
            connection = fillo.getConnection(filename.getFile());
        } catch (Exception e) {
            System.out.println("Failed to connect " + e.getMessage());
        }
    }

    public void closeConnection(){
        if(connection!=null){
            connection.close();
        }
    }
}
