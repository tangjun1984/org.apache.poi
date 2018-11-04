package com.demo1984s.poi.demo;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.net.URL;

public class App {

    /**
     *
     * @param args
     */
    public static void main(String[] args) {
        String password = "123456";
        URL url = App.class.getClassLoader().getResource("XSSFWorkbook.xlsx");
        System.out.println(url.getFile());
        // XSSFWorkbook
        encryptXSSFWorkbook(new File(url.getFile()), password);
        // HSSFWorkbook
        url = App.class.getClassLoader().getResource("HSSFWorkbook.xls");
        System.out.println(url.getFile());
        encryptHSSFWorkbook(new File(url.getFile()), password);
    }

    /**
     * 给 xls 文件加密码
     *
     * @param excelFile
     * @param password
     */
    private static void encryptHSSFWorkbook(File excelFile, String password) {
        try {
            Biff8EncryptionKey.setCurrentUserPassword(password);
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(new FileInputStream(excelFile));
            hssfWorkbook.writeProtectWorkbook(Biff8EncryptionKey.getCurrentUserPassword(), "bb");
            hssfWorkbook.unwriteProtectWorkbook();
            FileOutputStream fos = new FileOutputStream(excelFile);
            hssfWorkbook.write(fos);
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 给 xlsx 文件加密码
     *
     * @param excelFile
     * @param password
     */
    private static void encryptXSSFWorkbook(File excelFile, String password) {
        try {
            POIFSFileSystem fs = new POIFSFileSystem();
            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
            // EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile, CipherAlgorithm.aes192, HashAlgorithm.sha384, -1, -1, null);

            Encryptor enc = info.getEncryptor();
            enc.confirmPassword(password);

            // Read in an existing OOXML file and write to encrypted output stream
            // don't forget to close the output stream otherwise the padding bytes aren't added
            try (OPCPackage opc = OPCPackage.open(excelFile, PackageAccess.READ_WRITE);
                 OutputStream os = enc.getDataStream(fs)) {
                opc.save(os);
            }

            // Write out the encrypted version
            FileOutputStream fos = new FileOutputStream(excelFile);
            fs.writeFilesystem(fos);
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
