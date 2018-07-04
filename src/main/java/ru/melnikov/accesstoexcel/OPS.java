package ru.melnikov.accesstoexcel;

import org.apache.commons.lang.math.NumberUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class OPS {

    private Integer code;
    private String index;
    private Integer skupID;
    private String OpsName;
    private String postIndex;
    private String PostName;
    private String address;
    private String opsCalss;
    private String opsType;
    private String lvs;
    private String loopback;
    private String PtPreserv;
    private String RVPN_PE;
    private String gre_tun1;
    private String gre_tun2;
    private String pdk;
    private String pdkRes;



//    Код УФПС-1	Индекс	СКУП ID	Название ОПС	Индекс почтамта	Название ПТ	Адрес	Класс ОПС	Тип ОПС	ЛВС (ПКТ)	Loopback	резерв PtP /30 (CISCO2811-RVPN)	PtP /30 (RVPN-PE)	GRE tunnel1 /резерв /30	GRE tunnel2 /резерв /30	ПКД	Резерв ПКД

    public XSSFRow getRow(XSSFRow row){
        row.createCell(0,CellType.NUMERIC).setCellValue(code);

        if (NumberUtils.isNumber(this.getIndex())) row.createCell(1,CellType.NUMERIC).setCellValue(Integer.parseInt(index));
            else row.createCell(1,CellType.STRING).setCellValue(index);

        row.createCell(2,CellType.NUMERIC).setCellValue((skupID==null)?0:skupID);
        row.createCell(3,CellType.STRING).setCellValue(OpsName);
        row.createCell(4,CellType.STRING).setCellValue(postIndex);
        row.createCell(5,CellType.STRING).setCellValue(PostName);
        row.createCell(6,CellType.STRING).setCellValue(address);
        row.createCell(7,CellType.STRING).setCellValue(opsCalss);
        row.createCell(8,CellType.STRING).setCellValue(opsType);
        row.createCell(9,CellType.STRING).setCellValue(lvs);
        row.createCell(10,CellType.STRING).setCellValue(loopback);
        row.createCell(11,CellType.STRING).setCellValue(PtPreserv);
        row.createCell(12,CellType.STRING).setCellValue(RVPN_PE);
        row.createCell(13,CellType.STRING).setCellValue(gre_tun1);
        row.createCell(14,CellType.STRING).setCellValue(gre_tun2);
        row.createCell(15,CellType.STRING).setCellValue(pdk);
        row.createCell(16,CellType.STRING).setCellValue(pdkRes);




       return row;
    }


    public OPS(Integer code, String index, Integer skupID, String opsName, String postIndex, String postName, String address, String opsCalss, String opsType, String lvs, String loopback, String ptPreserv, String RVPN_PE, String gre_tun1, String gre_tun2, String pdk, String pdkRes) {
        this.code = code;
        this.index = index;
        this.skupID = skupID;
        OpsName = opsName;
        this.postIndex = postIndex;
        PostName = postName;
        this.address = address;
        this.opsCalss = opsCalss;
        this.opsType = opsType;
        this.lvs = lvs;
        this.loopback = loopback;
        PtPreserv = ptPreserv;
        this.RVPN_PE = RVPN_PE;
        this.gre_tun1 = gre_tun1;
        this.gre_tun2 = gre_tun2;
        this.pdk = pdk;
        this.pdkRes = pdkRes;
    }

    public OPS(Object[] array) {

    this.code = (Integer) array[0];
    this.index = ((String) array[1]).trim();
    this.skupID = (Integer) array[2];
    OpsName = (String) array[3];
    this.postIndex = (String) array[4];
    PostName = (String) array[5];
    this.address = (String) array[6];
    this.opsCalss = (String) array[7];
    this.opsType = (String) array[8];
    this.lvs = (String) array[9];
    this.loopback = (String) array[10];
    PtPreserv = (String) array[11];
    this.RVPN_PE = (String) array[12];
    this.gre_tun1 = (String) array[13];
    this.gre_tun2 = (String) array[14];
    this.pdk = (String) array[15];
    this.pdkRes = (String) array[16];

    }


    public Integer getCode() {
        return code;
    }

    public String getIndex() {
        return index;
    }

    public Integer getSkupID() {
        return skupID;
    }

    public String getOpsName() {
        return OpsName;
    }

    public String getPostIndex() {
        return postIndex;
    }

    public String getPostName() {
        return PostName;
    }

    public String getAddress() {
        return address;
    }

    public String getOpsCalss() {
        return opsCalss;
    }

    public String getOpsType() {
        return opsType;
    }

    public String getLvs() {
        return lvs;
    }

    public String getLoopback() {
        return loopback;
    }

    public String getPtPreserv() {
        return PtPreserv;
    }

    public String getRVPN_PE() {
        return RVPN_PE;
    }

    public String getGre_tun1() {
        return gre_tun1;
    }

    public String getGre_tun2() {
        return gre_tun2;
    }

    public String getPdk() {
        return pdk;
    }

    public String getPdkRes() {
        return pdkRes;
    }
}
