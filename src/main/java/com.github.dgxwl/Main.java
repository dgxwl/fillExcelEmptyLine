package com.github.dgxwl;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;

public class Main extends JPanel {

    private JTextArea textArea;
    private JMenuItem openItem;
    private FileDialog pickFileDialog;
    private File excelFile;

    private void init() {  //初始化窗口界面
        JFrame frame = new JFrame("补全excel删除行");
        frame.setSize(560, 200);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setResizable(false);
        frame.setLocationRelativeTo(null);

        JMenu fileMenu = new JMenu("操作");  //操作菜单
        openItem = new JMenuItem("选择文件");  //菜单项
        fileMenu.add(openItem);

        JMenuBar bar = new JMenuBar();  //菜单栏
        bar.add(fileMenu);
        frame.setJMenuBar(bar);

        textArea = new JTextArea(25, 50);  //文本域
        textArea.setText("提示: 操作->选择待补全excel文件, 稍等几秒钟即可看到结果");
        textArea.setEditable(false);
        this.add(new JScrollPane(textArea));

        pickFileDialog = new FileDialog(frame, "选择文件", FileDialog.LOAD);  //文件选择窗口

        frame.add(this);
        frame.setVisible(true);
    }

    private void handle(File excelFile) {
        String suffix = excelFile.getName().substring(excelFile.getName().lastIndexOf('.') + 1);
        if (!"xls".equals(suffix) && !"xlsx".equals(suffix)) {
            textArea.setText("无法解析" + suffix + "格式文件");
            return ;
        }

        String destDir = excelFile.getParent();
        String destFileStr = destDir + File.separator + "行补全.xlsx";
        File destFile = new File(destFileStr);

        try (XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
             FileOutputStream fos = new FileOutputStream(destFile);
             XSSFWorkbook newWorkBook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.getSheetAt(0);
            //获取列数
            int columnNum = sheet.getRow(0).getPhysicalNumberOfCells();
            //获得总行数
            int rowNum = sheet.getPhysicalNumberOfRows();

            if (columnNum < 1) {
                textArea.setText("文件无内容");
                return ;
            }
            if (rowNum < 1) {
                textArea.setText("文件无内容");
                return ;
            }

            XSSFRow titleRow = sheet.getRow(0);

            //创建新文件
            newWorkBook.createSheet();
            XSSFSheet newSheetAt = newWorkBook.getSheetAt(0);
            XSSFRow newSheetRow1 = newSheetAt.createRow(0);

            //复制表头
            for (int i = 0; i < columnNum; i++) {
                XSSFCell cell = titleRow.getCell(i);
                newSheetRow1.createCell(i);
                XSSFCell newCell = newSheetRow1.getCell(i);
                newCell.setCellValue(cell.getStringCellValue());
            }

            //表头增加一列是否为删行, 用于区别原本存在但是空的行
            XSSFCell isDelTitleCell = newSheetRow1.createCell(columnNum);
            isDelTitleCell.setCellValue("是否为删行");

            //拿到开始的序号
            XSSFRow r1 = sheet.getRow(1);
            XSSFCell c1 = r1.getCell(0);
            int startSeq = (int) c1.getNumericCellValue();  //开始序号

            //遍历原表的行
            for (int i = 1; i < rowNum; i++) {
                //拿原表每行第一列的序号
                XSSFRow row = sheet.getRow(i);
                XSSFCell cell = row.getCell(0);
                int rowSeq = (int) cell.getNumericCellValue();

                //序号对不上, 有空缺
                while (startSeq < rowSeq) {
                    //新文件里创建空行
                    XSSFRow newRow = newSheetAt.createRow(startSeq);
                    XSSFCell newCell = newRow.createCell(0);
                    newCell.setCellValue(startSeq);
                    newRow.createCell(columnNum).setCellValue("Y");  //标记为删行
                    startSeq++;
                }

                //把原有行拷贝到新文件中
                XSSFRow newRow = newSheetAt.createRow(startSeq);
                for (int j = 0; j < columnNum; j++) {
                    XSSFCell newCell = newRow.createCell(j);
                    XSSFCell cj = row.getCell(j);
                    if (cj == null) {
                        continue;
                    }
                    newCell.setCellType(cj.getCellType());

                    switch (cell.getCellType()) {
                        case STRING:
                            newCell.setCellValue(cj.getStringCellValue());
                            break;
                        case NUMERIC:
                            newCell.setCellValue(cj.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            newCell.setCellValue(cj.getBooleanCellValue());
                            break;
                        case FORMULA:
                            newCell.setCellValue(cj.getCellFormula());
                            break;
                        case ERROR:
                            newCell.setCellValue(cj.getErrorCellValue());
                            break;
                        default:
                            break;
                    }
                }

                startSeq++;
            }

            newWorkBook.write(fos);

            //显示结果
            textArea.setText("操作成功！" + "\n已生成文件：\n" + destFileStr);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void handleEvent() {
        //监听打开选择文件菜单项
        openItem.addActionListener(e -> {
            pickFileDialog.setVisible(true);

            String dirPath = pickFileDialog.getDirectory();
            String fileName = pickFileDialog.getFile();

            if (dirPath == null || fileName == null) {
                return ;
            }

            excelFile = new File(dirPath, fileName);

            handle(excelFile);
        });
    }

    public static void main(String[] args) {
        Main main = new Main();
        main.init();
        main.handleEvent();
    }
}
