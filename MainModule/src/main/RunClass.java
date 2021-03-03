package main;

import java.awt.FlowLayout;
import java.io.File;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.filechooser.FileNameExtensionFilter;

public class RunClass {
    public RunClass() {
    }

    private void init() {
        this.setFrame();
        this.attachBtns(resources.MAIN_FRAME);
        this.show();
    }

    private void setFrame() {
        resources.MAIN_FRAME = new JFrame("Excel Braker");
        resources.MAIN_FRAME.setSize(700, 300);
        resources.MAIN_FRAME.setDefaultCloseOperation(3);
        resources.MAIN_FRAME.setLayout(new FlowLayout());
    }

    private void attachBtns(JFrame frame) {
        JButton getFileBtn = new JButton("Select File");

        getFileBtn.addActionListener((e) -> {
            JFileChooser fileChooser = new JFileChooser();
            FileNameExtensionFilter filter = new FileNameExtensionFilter("ExcelFiles", "xlsx", "xlsm", "xls");
            fileChooser.setFileFilter(filter);
            JFrame chooserFrame = new JFrame("Select File");
            int returnVal = fileChooser.showOpenDialog(chooserFrame);
            if (returnVal == 0) {
                resources.FILE_PATH = fileChooser.getSelectedFile().getPath();
                resources.EXCEL_FILE = new File(resources.FILE_PATH);
            }

        });

        JButton startConversionBtn = new JButton("Crack!");
        startConversionBtn.addActionListener((e) -> {
            if (resources.EXCEL_FILE.renameTo(new File(resources.EXCEL_FILE.getAbsoluteFile() + ".zip"))) {
                System.out.println("done");
                new ExcelBreaker(new File(resources.EXCEL_FILE.getAbsoluteFile() + ".zip"));
            }

        });

        frame.add(getFileBtn);
        frame.add(startConversionBtn);
    }

    private void show() {
        resources.MAIN_FRAME.setVisible(true);
    }

    public static void main(String[] args) {
        System.out.println("hello world");
        RunClass tmp = new RunClass();
        tmp.init();
    }
}
