package com.word.docx;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;

/**
 * Created by wyz on 15-11-11.
 * Date: 2015-11-11
 * Time: 下午7:31
 */
public class Main {

    public static void main(String[] args) throws Exception {


        // first line: The full path of rules file.
        // second line: The folder contains all the replaced files.
        List<String> config = readConfig();

        String rulesfile = config.get(0);  // "/home/wyz/Desktop/word/替换规则.txt";
        String dir = config.get(1);        // "/home/wyz/Desktop/word/程序及表单";


        // read the rules in file, and save in HashMap.
        HashMap<String, String> rules = readReplacementRulesFiles(rulesfile);
        rules.put("[", "");
        rules.put("]", "");


        // Reading the replaced files.
        List<String> srcFiles = readAllFiles(dir);


        // set the new folder.
        String old_dirname = dir.substring(dir.lastIndexOf(File.separator) + 1, dir.length());
        String new_dirname = old_dirname + "-替换后";


        int count = 0;


        // replace and export docx files.
        for (int i = 0; i < srcFiles.size(); ++i) {
            String srcFile = srcFiles.get(i);
            if (getExtensionName(srcFile)) {
                String destFile = srcFile.replaceFirst(old_dirname, new_dirname);

                // creat the new file
                File new_folder = new File(destFile.substring(0, destFile.lastIndexOf(File.separator)));
                new_folder.mkdirs();

                // replace and export
                exportdocx(srcFile, destFile, rules);
                count++;
            }
        }
        System.out.println("共处理了" + count + "个文件。");

    }

    /**
     * get the content of config.
     *
     * @return the path of rules file and the folder of replaced files.
     */
    public static List<String> readConfig() {
        System.out.println("正在读取当前目录下的config.txt文件.");

        List<String> readfile = new ArrayList<String>();
        readfile.clear();

        File file = null;
        BufferedReader reader = null;
        try {
            file = new File("config.txt");
            reader = new BufferedReader(new FileReader(file));
            String tempString = null;
            int line = 1;
            // 一次读入一行，直到读入null为文件结束
            while ((tempString = reader.readLine()) != null) {
                System.out.println("配置信息" + line + ": " + tempString + ".");
                readfile.add(tempString);
                line++;
            }
            reader.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (reader != null) {
                try {
                    reader.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
        }
        System.out.println("读取config.txt文件结束.");
        return readfile;
    }


    /**
     * get the rules.
     *
     * @param fileName the path of rules file.
     * @return the rules: key:source word, value:destine word
     */
    public static HashMap<String, String> readReplacementRulesFiles(String fileName) {
        System.out.println("正在读取替换规则文件.");
        HashMap<String, String> repRules = new HashMap<String, String>();
        repRules.clear();

        File file = null;
        BufferedReader reader = null;
        try {
            file = new File(fileName);
            reader = new BufferedReader(new FileReader(file));
            String tempString = null;

            // 一次读入一行，直到读入null为文件结束
            while ((tempString = reader.readLine()) != null) {
                String[] srctodes = tempString.split("=");
                repRules.put(srctodes[0].substring(1, srctodes[0].length() - 1), srctodes[1]);
            }
            reader.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (reader != null) {
                try {
                    reader.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
        }
        System.out.println("读取替换规则文件结束.");
        return repRules;
    }


    /**
     * get all files's list.
     *
     * @param srcPath the root folder of files.
     * @return the list of files
     */
    public static List<String> readAllFiles(final String srcPath) {
        System.out.println("正在读取所有需要替换的文件.");
        final List<String> files = new ArrayList<String>();

        Queue<String> que = new LinkedList<String>();
        que.clear();

        que.add(srcPath);
        while (!que.isEmpty()) {
            String ele = que.poll();
            final File file = new File(ele);

            if (file.isDirectory()) {
                final String[] filelist = file.list();
                for (final String subele : filelist) {
                    que.add(ele + File.separator + subele);
                }
            } else {
                //if(getExtensionName(file.getName()))
                files.add(file.getPath());
            }
        }
        System.out.println("读取所有需要替换的文件结束.");
        return files;
    }


    /**
     * replace function.
     *
     * @param srcFile  the path of the replaced file
     * @param destFile the export path of the new file
     * @param rules    the rules of replacement
     * @throws Exception
     */
    public static void exportdocx(String srcFile, String destFile, HashMap<String, String> rules) throws Exception {
        System.out.println("正在处理" + srcFile + ".");

        XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(srcFile));

        // replace paragraphs
        Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
        while (itPara.hasNext()) {
            XWPFParagraph paragraph = itPara.next();

            Iterator<String> iterator = rules.keySet().iterator();
            while (iterator.hasNext()) {
                String key = iterator.next();

                //System.out.println("getText:" + paragraph.getParagraphText());
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null) {
                        boolean isSetText = false;
                        if (text.indexOf(key) != -1) {
                            isSetText = true;
                            text = text.replace(key, rules.get(key));
                        }
                        if (isSetText) {
                            //参数0表示生成的文字是要从哪一个地方开始放置,设置文字从位置0开始,就可以把原来的文字全部替换掉了
                            run.setText(text, 0);
                        }
                    }
                }
            }
        }

        // replace tables
        Iterator<XWPFTable> it = document.getTablesIterator();
        while (it.hasNext()) {
            XWPFTable table = it.next();

            List<XWPFTableRow> rows = table.getRows();
            for (XWPFTableRow row : rows) {
                List<XWPFTableCell> cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {
                    List<XWPFParagraph> paragraphListTable = cell.getParagraphs();

                    for (XWPFParagraph paragraph : paragraphListTable) {
                        //String s = paragraph.getParagraphText();
                        Iterator<String> iterator = rules.keySet().iterator();
                        while (iterator.hasNext()) {
                            String key = iterator.next();
                            List<XWPFRun> runs = paragraph.getRuns();
                            for (XWPFRun run : runs) {
                                String text = run.getText(0);
                                if (text != null) {
                                    boolean isSetText = false;
                                    if (text.indexOf(key) != -1) {
                                        isSetText = true;
                                        text = text.replace(key, rules.get(key));
                                    }
                                    if (isSetText) {
                                        //参数0表示生成的文字是要从哪一个地方开始放置,设置文字从位置0开始,就可以把原来的文字全部替换掉了
                                        run.setText(text, 0);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        // replace headers paragraphs
        List<XWPFHeader> headers_paragraph = document.getHeaderList();
        for (XWPFHeader header : headers_paragraph) {
            List<XWPFParagraph> paragraphs = header.getParagraphs();
            for (XWPFParagraph paragraph : paragraphs) {
                //String s = paragraph.getParagraphText();
                Iterator<String> iterator = rules.keySet().iterator();
                while (iterator.hasNext()) {
                    String key = iterator.next();
                    List<XWPFRun> runs = paragraph.getRuns();
                    for (XWPFRun run : runs) {
                        String text = run.getText(0);
                        if (text != null) {
                            boolean isSetText = false;
                            if (text.indexOf(key) != -1) {
                                isSetText = true;
                                text = text.replace(key, rules.get(key));
                            }
                            if (isSetText) {
                                //参数0表示生成的文字是要从哪一个地方开始放置,设置文字从位置0开始,就可以把原来的文字全部替换掉了
                                run.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }

        // replace headers tables
        List<XWPFHeader> headers_table = document.getHeaderList();
        for (XWPFHeader header : headers_table) {
            List<XWPFTable> tables = header.getTables();
            for (XWPFTable table : tables) {
                List<XWPFTableRow> rows = table.getRows();
                for (XWPFTableRow row : rows) {
                    List<XWPFTableCell> cells = row.getTableCells();
                    for (XWPFTableCell cell : cells) {
                        List<XWPFParagraph> paragraphListTable = cell.getParagraphs();
                        for (XWPFParagraph paragraph : paragraphListTable) {
                            //String s = paragraph.getParagraphText();
                            Iterator<String> iterator = rules.keySet().iterator();
                            while (iterator.hasNext()) {
                                String key = iterator.next();
                                List<XWPFRun> runs = paragraph.getRuns();
                                for (XWPFRun run : runs) {
                                    String text = run.getText(0);
                                    if (text != null) {
                                        boolean isSetText = false;
                                        if (text.indexOf(key) != -1) {
                                            isSetText = true;
                                            text = text.replace(key, rules.get(key));
                                        }
                                        if (isSetText) {
                                            //参数0表示生成的文字是要从哪一个地方开始放置,设置文字从位置0开始,就可以把原来的文字全部替换掉了
                                            run.setText(text, 0);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        //System.out.println(document.toString());
        FileOutputStream outStream = new FileOutputStream(destFile);
        document.write(outStream);
        outStream.close();
    }


    /**
     * get the filename extension.
     *
     * @param filename the path of files
     * @return the extension of file
     */
    public static boolean getExtensionName(String filename) {
        if ((filename != null) && (filename.length() > 0)) {
            int dot = filename.lastIndexOf('.');
            if ((dot > -1) && (dot < (filename.length() - 1))) {
                return filename.substring(dot + 1).equals("docx");
            }
        }
        return false;
    }

}
