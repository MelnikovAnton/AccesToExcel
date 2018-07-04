package ru.melnikov.accesstoexcel;

import org.apache.commons.dbutils.QueryRunner;
import org.apache.commons.dbutils.handlers.ArrayListHandler;
import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPClient;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Properties;

public class Action {

    private static final Logger log = Logger.getLogger(Action.class);

    private final Properties props;

    private static final String[]HEADER= {"Код УФПС-1","Индекс","СКУП ID","Название ОПС","Индекс почтамта","Название ПТ","Адрес","Класс ОПС","Тип ОПС","ЛВС (ПКТ)","Loopback","резерв PtP /30 (CISCO2811-RVPN)","PtP /30 (RVPN-PE)","GRE tunnel1 /резерв /30","GRE tunnel2 /резерв /30","ПКД","Резерв ПКД"};

    public Action(Properties props) {
        this.props=props;
    }

    public boolean downloadFile(String localPath){
        String server = props.getProperty("ftp.server");
        int port = Integer.parseInt(props.getProperty("ftp.port"));
        String user = props.getProperty("ftp.user");
        String pass = props.getProperty("ftp.pass");

        boolean success=false;

        FTPClient ftpClient = new FTPClient();

        try {
            ftpClient.connect(server, port);
            ftpClient.login(user, pass);
            ftpClient.enterLocalPassiveMode();
            ftpClient.setFileType(FTP.BINARY_FILE_TYPE);

            String remoteFile1 = props.getProperty("remoteFileName");
            log.debug("Downloading file from server " + server+ ":" + port + remoteFile1);
            File downloadFile1 = new File(localPath);
            OutputStream outputStream1 = new BufferedOutputStream(new FileOutputStream(downloadFile1));
            success = ftpClient.retrieveFile(remoteFile1, outputStream1);
            outputStream1.close();

            if (success) {
                log.info("File has been downloaded successfully.");
            }

        }catch (IOException e) {
            e.printStackTrace();
        }finally {
            try {
                if (ftpClient.isConnected()) {
                    ftpClient.logout();
                    ftpClient.disconnect();
                }
            } catch (IOException ex) {
                ex.printStackTrace();
            }
            return success;
        }
    }

    public void createXLSX(String filename){
        try {

            Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");

            Connection conn = DriverManager.getConnection("jdbc:ucanaccess://".concat(props.getProperty("localDBpath")));

            String sql = "SELECT * FROM ОПС";

            Object[] arr =  new QueryRunner()
                    .query(conn, sql, new ArrayListHandler())
                    .stream()
                    .map(array -> new OPS(array))
                    .toArray();


            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Лист1");

            XSSFCellStyle style = workbook.createCellStyle();
            style.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());

            XSSFRow row = sheet.createRow(0);
            for(int i=0;i<HEADER.length;i++){
                XSSFCell cell = row.createCell(i, CellType.STRING);
                cell.setCellValue(HEADER[i]);
            }

            for (int i=0;i<arr.length;i++) {
                OPS o= (OPS) arr[i];
                o.getRow(sheet.createRow(i+1));
            }

            workbook.write(new FileOutputStream(filename));
            workbook.close();
        }catch (IOException | SQLException | ClassNotFoundException e) {
            e.printStackTrace();
        }
    }
}
