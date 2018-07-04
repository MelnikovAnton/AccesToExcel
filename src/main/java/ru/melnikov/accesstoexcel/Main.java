package ru.melnikov.accesstoexcel;

import org.apache.log4j.Logger;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class Main {

    private static final Logger log = Logger.getLogger(Main.class);

    private static final Properties props = new Properties();

    public static void main(String[] args) throws IOException {

        log.info("Start init...");
        try (InputStream inputStream = ClassLoader.getSystemResourceAsStream("app.properties")) {
            props.load(inputStream);
        }

        Action act = new Action(props);
        if (act.downloadFile(props.getProperty("localDBpath"))) {
            log.info("File downloaded successfull. Starting create EXCEL document ");
            act.createXLSX(props.getProperty("localExcelPath"));
        }
            else log.warn("Cannot download remote file");
        log.info("Stop application...");
    }
}
