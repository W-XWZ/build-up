package test.test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.UUID;

import org.jodconverter.OfficeDocumentConverter;
import org.jodconverter.office.DefaultOfficeManagerBuilder;
import org.jodconverter.office.ExternalOfficeManagerBuilder;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeManager;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;

public class LibreOfficeAndJodconverter {
    private static OfficeManager officeManager = null;
    private static final String dirPath = "D:/word/";
    private static final String LibreOfficeDirPath = "C:/Program Files/LibreOffice";
    public static void init() throws OfficeException {
        try {
            System.out.println("尝试连接已启动的服务...");
            ExternalOfficeManagerBuilder externalProcessOfficeManager = new ExternalOfficeManagerBuilder();
            externalProcessOfficeManager.setConnectOnStart(true);
            externalProcessOfficeManager.setPortNumber(8100);
            officeManager = externalProcessOfficeManager.build();
            officeManager.start();
            System.out.println("转换服务启动成功!");
        } catch (Exception e) {
            //命令方式：soffice -headless -accept="socket,host=127.0.0.1,port=8100;urp;" -nofirststartwizard
            System.out.println("启动新服务!");
            String libreOfficePath = LibreOfficeDirPath;
            // 此类在jodconverter-core中3版本中存在，在2.2.2版本中不存在
            DefaultOfficeManagerBuilder configuration = new DefaultOfficeManagerBuilder();
            // libreOffice的安装目录
            configuration.setOfficeHome(new File(libreOfficePath));
            // 设置端口号
            configuration.setPortNumbers(8100);
            // 设置任务执行超时为5分钟
            configuration.setTaskExecutionTimeout(1000 * 60 * 5L);
            // 设置任务队列超时为24小时
            configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);
 
            // 开启转换服务
            officeManager = configuration.build();
            officeManager.start();
            System.out.println("服务启动成功!");
        }
    }
    public static void desory() throws OfficeException {
        if (officeManager != null) {
            officeManager.stop();
        }
    }
    /** * 合并多个PDF，注意，源pdf中不能含有目的pdf，否则将合并失败*/
    public static boolean mergePdfFiles(String[] files, String newfile) {
        boolean retValue = false;
        Document document = null;
        try {
            document = new Document(new PdfReader(files[0]).getPageSize(1));
            PdfCopy copy = new PdfCopy(document, new FileOutputStream(newfile));
            document.open();
            for (int i = 0; i < files.length; i++) {
                PdfReader reader = new PdfReader(files[i]);
                int n = reader.getNumberOfPages();
                for (int j = 1; j <= n; j++) {
                    document.newPage();
                    PdfImportedPage page = copy.getImportedPage(reader, j);
                    copy.addPage(page);
                }
            }
            retValue = true;
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            document.close();
        }
        return retValue;
    }
    /**
     * 开启服务时耗时，需要安装
     * 并不是看到什么，就转化为什么样的。有偏移
     *
     * @param args
     */
    public static void main(String[] args) {
        try {
			init();
	        task();
	        desory();
		} catch (OfficeException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
    public static void task() throws OfficeException {
        String outputname = "output.pdf";
//        doDocToFdpLibre("test.docx", outputname);
//        doDocToFdpLibre("ppt.pptx", outputname);
        doDocToFdpLibre("4.xlsx", outputname);
//        doDocToFdpLibre("1.txt", outputname);
//        doDocToFdpLibre("不老梦.jpg", outputname);
    }
    public static String doDocToFdpLibre(String inputFileName, String outputFileName) throws OfficeException {
        File inputFile = new File("D:/word/" + inputFileName);
        System.out.println("libreOffice开始转换..............................");
        Long startTime = System.currentTimeMillis();
        OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
 
        File outputFile = new File(dirPath + outputFileName);
        if (outputFile.exists()) {
            String uuid1 = UUID.randomUUID().toString();
            File temp1 = new File(dirPath + uuid1 + ".pdf");
            outputFile.renameTo(temp1);
 
            String uuid = UUID.randomUUID().toString();
            File temp2 = new File(dirPath + uuid + ".pdf");
            converter.convert(inputFile, temp2);
 
            String[] files = {dirPath + uuid1 + ".pdf", dirPath + uuid + ".pdf"};
            String savepath = dirPath + outputFileName;
 
            if (mergePdfFiles(files, savepath)) {
                temp1.delete();
                temp2.delete();
            }
        } else {//DocumentFormat outputFormat = new DocumentFo
            converter.convert(inputFile, outputFile);
        }
        // 转换结束
        System.out.println("转换结束。。。。。");
        //转换时间
        long endTime = System.currentTimeMillis();
        long time = endTime - startTime;
        System.out.println("libreOffice转换所用时间为：" + time);
        return outputFile.getPath();
    }
}